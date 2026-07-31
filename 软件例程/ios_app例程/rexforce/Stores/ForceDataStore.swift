//
//  ForceDataStore.swift
//  rexforce
//
//  Created by 陈学瀚 on 2026/7/28.
//

import Foundation
import Combine

/// 力值数据存储
///
/// 管理实时力值数据的环形缓冲区，最多保留 10000 个数据点。
/// 支持软件归零（Tare）和模式切换时自动清空。
@MainActor
final class ForceDataStore: ObservableObject {

    // MARK: - 常量
    static let maxPoints = 10000
    /// 中值滤波窗口大小，取奇数以便获得稳定中位数
    private static let medianWindowSize = 5
    /// 实时力值滤波系数，越小越平滑，越大响应越快
    private static let filterAlpha = 0.24

    // MARK: - Published 属性
    @Published var mode: DeviceMode = .unknown
    /// 触发 UI 刷新的计数器
    @Published var updateTick: Int = 0
    /// 当前主设备力值
    @Published var currentMainForce: Double = 0
    /// 当前副设备力值
    @Published var currentSecondaryForce: Double = 0
    /// 当前合力
    @Published var currentCombinedForce: Double = 0
    /// 总接收采样数
    @Published var totalSampleCount: Int = 0
    /// 采样率（Hz）
    @Published var sampleRate: Double = 0
    /// 归零状态
    @Published var isZeroed: Bool = false

    // MARK: - 数据缓冲区（非 Published，直接访问以提高性能）
    /// 主设备力值数组
    private(set) var mainForces: [Double] = []
    /// 副设备力值数组（单机模式为空）
    private(set) var secondaryForces: [Double] = []
    /// 合力数组（单机模式为空）
    private(set) var combinedForces: [Double] = []

    // MARK: - 私有属性
    /// 软件归零偏移
    private var zeroOffsetMain: Double = 0
    private var zeroOffsetSecondary: Double = 0
    /// 采样率计算
    private var rateCalcLastTime: CFAbsoluteTime = 0
    private var rateCalcLastCount: Int = 0
    /// EMA 滤波器状态
    private var filteredMainForce: Double?
    private var filteredSecondaryForce: Double?
    /// 中值滤波窗口
    private var recentMainInputs: [Double] = []
    private var recentSecondaryInputs: [Double] = []

    // MARK: - 初始化

    init() {
        preAllocateBuffers()
    }

    /// 预分配缓冲区容量
    private func preAllocateBuffers() {
        mainForces.reserveCapacity(Self.maxPoints + 100)
        secondaryForces.reserveCapacity(Self.maxPoints + 100)
        combinedForces.reserveCapacity(Self.maxPoints + 100)
        recentMainInputs.reserveCapacity(Self.medianWindowSize)
        recentSecondaryInputs.reserveCapacity(Self.medianWindowSize)
    }

    // MARK: - 数据写入

    /// 添加解析后的采样数据
    func appendSamples(_ result: FrameParser.ParseResult) {
        if result.mode != mode {
            clearData()
            mode = result.mode
        }

        let samples = result.samples
        let count = samples.count

        for sample in samples {
            let adjustedMain = sample.mainForce - zeroOffsetMain
            let filteredMain = applyCombinedFilter(
                input: adjustedMain,
                recentInputs: &recentMainInputs,
                previousEma: &filteredMainForce
            )
            mainForces.append(filteredMain)

            if let sec = sample.secondaryForce {
                let adjustedSec = sec - zeroOffsetSecondary
                let filteredSec = applyCombinedFilter(
                    input: adjustedSec,
                    recentInputs: &recentSecondaryInputs,
                    previousEma: &filteredSecondaryForce
                )
                secondaryForces.append(filteredSec)
                combinedForces.append(filteredMain + filteredSec)
                currentSecondaryForce = filteredSec
                currentCombinedForce = filteredMain + filteredSec
            } else {
                currentSecondaryForce = 0
                currentCombinedForce = filteredMain
            }

            currentMainForce = filteredMain
        }

        totalSampleCount += count
        trimBuffers()
        updateSampleRate()
        updateTick &+= 1
    }

    /// 组合滤波：先中值滤波去尖峰，再用 EMA 做低通平滑
    private func applyCombinedFilter(
        input: Double,
        recentInputs: inout [Double],
        previousEma: inout Double?
    ) -> Double {
        recentInputs.append(input)
        if recentInputs.count > Self.medianWindowSize {
            recentInputs.removeFirst(recentInputs.count - Self.medianWindowSize)
        }

        let medianFiltered = median(of: recentInputs)
        return applyLowPassFilter(input: medianFiltered, previous: &previousEma)
    }

    /// 计算窗口中的中位数
    private func median(of values: [Double]) -> Double {
        guard !values.isEmpty else { return 0 }
        let sorted = values.sorted()
        let middle = sorted.count / 2

        if sorted.count.isMultiple(of: 2) {
            return (sorted[middle - 1] + sorted[middle]) / 2
        }

        return sorted[middle]
    }

    /// 一阶指数低通滤波器
    private func applyLowPassFilter(input: Double, previous: inout Double?) -> Double {
        guard let last = previous else {
            previous = input
            return input
        }

        let filtered = last + Self.filterAlpha * (input - last)
        previous = filtered
        return filtered
    }

    /// 裁剪缓冲区到最大长度
    private func trimBuffers() {
        let overflow = mainForces.count - Self.maxPoints
        if overflow > 0 {
            mainForces.removeFirst(overflow)
            if !secondaryForces.isEmpty {
                secondaryForces.removeFirst(overflow)
            }
            if !combinedForces.isEmpty {
                combinedForces.removeFirst(overflow)
            }
        }
    }

    /// 更新采样率
    private func updateSampleRate() {
        let now = CFAbsoluteTimeGetCurrent()
        if rateCalcLastTime == 0 {
            rateCalcLastTime = now
            rateCalcLastCount = totalSampleCount
            return
        }

        let elapsed = now - rateCalcLastTime
        if elapsed >= 1.0 {
            let countDelta = totalSampleCount - rateCalcLastCount
            sampleRate = Double(countDelta) / elapsed
            rateCalcLastTime = now
            rateCalcLastCount = totalSampleCount
        }
    }

    // MARK: - 归零

    /// 软件归零：将当前读数设为零点
    func zero() {
        let recentCount = min(50, mainForces.count)
        if recentCount > 0 {
            let start = mainForces.count - recentCount
            let sum = mainForces[start...].reduce(0, +)
            zeroOffsetMain += sum / Double(recentCount)
        }

        if mode == .dual, secondaryForces.count >= recentCount {
            let start = secondaryForces.count - recentCount
            let sum = secondaryForces[start...].reduce(0, +)
            zeroOffsetSecondary += sum / Double(recentCount)
        }

        isZeroed = true
        recalculateWithOffset()
    }

    /// 使用新的零点偏移重新计算缓冲区
    private func recalculateWithOffset() {
        clearData(keepOffset: true)
    }

    // MARK: - 清空

    /// 清空所有数据
    /// - Parameter keepOffset: 是否保留归零偏移
    func clearData(keepOffset: Bool = false) {
        mainForces.removeAll(keepingCapacity: true)
        secondaryForces.removeAll(keepingCapacity: true)
        combinedForces.removeAll(keepingCapacity: true)
        recentMainInputs.removeAll(keepingCapacity: true)
        recentSecondaryInputs.removeAll(keepingCapacity: true)
        filteredMainForce = nil
        filteredSecondaryForce = nil
        currentMainForce = 0
        currentSecondaryForce = 0
        currentCombinedForce = 0
        totalSampleCount = 0
        sampleRate = 0
        rateCalcLastTime = 0
        rateCalcLastCount = 0
        updateTick &+= 1

        if !keepOffset {
            zeroOffsetMain = 0
            zeroOffsetSecondary = 0
            isZeroed = false
        }
    }

    /// 完全重置（断开连接时调用）
    func reset() {
        clearData(keepOffset: false)
        mode = .unknown
    }

    // MARK: - 统计信息

    /// 获取当前缓冲区中数据的统计范围
    func valueRange() -> (min: Double, max: Double) {
        if mainForces.isEmpty {
            return (-10, 10)
        }

        var minVal = Double.infinity
        var maxVal = -Double.infinity

        for v in mainForces {
            if v < minVal { minVal = v }
            if v > maxVal { maxVal = v }
        }

        if mode == .dual {
            for v in secondaryForces {
                if v < minVal { minVal = v }
                if v > maxVal { maxVal = v }
            }
            for v in combinedForces {
                if v < minVal { minVal = v }
                if v > maxVal { maxVal = v }
            }
        }

        if minVal == maxVal {
            minVal -= 1
            maxVal += 1
        }

        let range = maxVal - minVal
        let padding = range * 0.1
        return (minVal - padding, maxVal + padding)
    }

    /// 当前数据点数
    var pointCount: Int {
        mainForces.count
    }
}
