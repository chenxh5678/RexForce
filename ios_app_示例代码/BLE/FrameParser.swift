//
//  FrameParser.swift
//  rexforce
//
//  Created by 陈学瀚 on 2026/7/28.
//

import Foundation

/// BLE 数据帧解析器
///
/// 负责将 244 字节固定长度帧解析为 `ForceSample` 数组。
/// 支持单机模式（FUNC=0x03，80 个数据点）和双机模式（FUNC=0x05，40 组数据）。
struct FrameParser {

    /// 解析结果
    struct ParseResult {
        /// 首个数据点时间戳（ms）
        let timestamp: UInt16
        /// 设备模式
        let mode: DeviceMode
        /// 采样点列表
        let samples: [ForceSample]
    }

    /// 解码 3 字节重量点
    /// - Parameters:
    ///   - high: 高字节
    ///   - low: 低字节
    ///   - checksum: 校验字节
    /// - Returns: 重量值（kg），校验失败返回 nil
    static func decodePoint(high: UInt8, low: UInt8, checksum: UInt8) -> Double? {
        // 校验和 = (HIGH + LOW) & 0xFF
        guard (high &+ low) == checksum else { return nil }

        // 大端有符号 16 位整数
        let raw = (UInt16(high) << 8) | UInt16(low)
        let signed = Int16(bitPattern: raw)

        // 值 = 重量(kg) × 10
        return Double(signed) / 10.0
    }

    /// 解析 244 字节完整帧
    /// - Parameter frame: 帧数据（必须为 244 字节）
    /// - Returns: 解析结果，帧格式错误返回 nil
    static func parse(frame: Data) -> ParseResult? {
        guard frame.count == BLEConstants.frameLength else { return nil }

        let header = frame[0]
        let funcCode = frame[1]

        // 验证帧头和功能码
        guard header == BLEConstants.frameHeader,
              funcCode == BLEConstants.funcSingle || funcCode == BLEConstants.funcDual else {
            return nil
        }

        // 时间戳（大端，ms）
        let timestamp = (UInt16(frame[2]) << 8) | UInt16(frame[3])

        let mode = DeviceMode.from(funcCode: funcCode)
        var samples: [ForceSample] = []

        if funcCode == BLEConstants.funcSingle {
            // 单机模式：240 字节 = 80 个数据点 × 3 字节/点
            for i in stride(from: 0, to: BLEConstants.dataFieldLength, by: 3) {
                let base = 4 + i
                if let main = decodePoint(high: frame[base], low: frame[base + 1], checksum: frame[base + 2]) {
                    samples.append(ForceSample(mainForce: main, secondaryForce: nil))
                }
            }
        } else {
            // 双机模式：240 字节 = 40 组 × 6 字节/组
            for i in stride(from: 0, to: BLEConstants.dataFieldLength, by: 6) {
                let base = 4 + i
                let main = decodePoint(high: frame[base], low: frame[base + 1], checksum: frame[base + 2])
                let secondary = decodePoint(high: frame[base + 3], low: frame[base + 4], checksum: frame[base + 5])

                if let main = main, let secondary = secondary {
                    samples.append(ForceSample(mainForce: main, secondaryForce: secondary))
                }
            }
        }

        return ParseResult(timestamp: timestamp, mode: mode, samples: samples)
    }
}
