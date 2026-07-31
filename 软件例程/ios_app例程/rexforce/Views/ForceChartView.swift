//
//  ForceChartView.swift
//  rexforce
//
//  Created by 陈学瀚 on 2026/7/28.
//

import SwiftUI

/// 实时力值曲线图
///
/// 使用 Canvas 高性能渲染，支持最多 10000 个数据点实时滚动显示。
/// - 单机模式：显示 1 条力值曲线
/// - 双机模式：显示 3 条曲线（合力、主设备力、副设备力）
struct ForceChartView: View {

    @ObservedObject var dataStore: ForceDataStore

    // MARK: - 曲线颜色
    private let combinedColor = Color.blue
    private let mainColor = Color.orange
    private let secondaryColor = Color.purple
    private let singleColor = Color.blue

    // MARK: - 图表参数
    private let leftPadding: CGFloat = 56
    private let rightPadding: CGFloat = 16
    private let topPadding: CGFloat = 16
    private let bottomPadding: CGFloat = 32
    private let gridColor = Color.gray.opacity(0.2)
    private let axisColor = Color.secondary.opacity(0.5)

    var body: some View {
        GeometryReader { geometry in
            Canvas { context, size in
                // 显式引用 updateTick 确保 Canvas 随数据更新重绘
                let _ = dataStore.updateTick
                drawChart(context: context, size: size)
            }
            .background(Color(.systemBackground))
            .clipShape(RoundedRectangle(cornerRadius: 12))
        }
    }

    // MARK: - 绘图

    private func drawChart(context: GraphicsContext, size: CGSize) {
        let chartArea = CGRect(
            x: leftPadding,
            y: topPadding,
            width: size.width - leftPadding - rightPadding,
            height: size.height - topPadding - bottomPadding
        )

        // 获取数据范围
        let (yMin, yMax) = dataStore.valueRange()
        let pointCount = dataStore.pointCount

        // 绘制网格和坐标轴
        drawGrid(context: context, chartArea: chartArea, yMin: yMin, yMax: yMax)

        // 绘制 Y 轴标签
        drawYAxisLabels(context: context, chartArea: chartArea, yMin: yMin, yMax: yMax)

        // 绘制 X 轴标签
        drawXAxisLabels(context: context, chartArea: chartArea, pointCount: pointCount)

        // 绘制曲线
        guard pointCount > 1 else {
            drawEmptyState(context: context, chartArea: chartArea)
            return
        }

        let mode = dataStore.mode
        if mode == .single {
            drawCurve(
                context: context,
                chartArea: chartArea,
                data: dataStore.mainForces,
                color: singleColor,
                yMin: yMin,
                yMax: yMax,
                pointCount: pointCount
            )
        } else if mode == .dual {
            // 合力
            drawCurve(
                context: context,
                chartArea: chartArea,
                data: dataStore.combinedForces,
                color: combinedColor,
                yMin: yMin,
                yMax: yMax,
                pointCount: pointCount
            )
            // 主设备
            drawCurve(
                context: context,
                chartArea: chartArea,
                data: dataStore.mainForces,
                color: mainColor,
                yMin: yMin,
                yMax: yMax,
                pointCount: pointCount
            )
            // 副设备
            drawCurve(
                context: context,
                chartArea: chartArea,
                data: dataStore.secondaryForces,
                color: secondaryColor,
                yMin: yMin,
                yMax: yMax,
                pointCount: pointCount
            )
        }

        // 绘制图例
        drawLegend(context: context, size: size, mode: mode)

        // 绘制零线
        drawZeroLine(context: context, chartArea: chartArea, yMin: yMin, yMax: yMax)
    }

    // MARK: - 网格绘制

    private func drawGrid(context: GraphicsContext, chartArea: CGRect, yMin: Double, yMax: Double) {
        // 水平网格线（5 等分）
        let horizontalLines = 5
        for i in 0...horizontalLines {
            let y = chartArea.minY + chartArea.height * CGFloat(i) / CGFloat(horizontalLines)
            var path = Path()
            path.move(to: CGPoint(x: chartArea.minX, y: y))
            path.addLine(to: CGPoint(x: chartArea.maxX, y: y))
            context.stroke(path, with: .color(gridColor), lineWidth: 0.5)
        }

        // 垂直网格线（10 等分）
        let verticalLines = 10
        for i in 0...verticalLines {
            let x = chartArea.minX + chartArea.width * CGFloat(i) / CGFloat(verticalLines)
            var path = Path()
            path.move(to: CGPoint(x: x, y: chartArea.minY))
            path.addLine(to: CGPoint(x: x, y: chartArea.maxY))
            context.stroke(path, with: .color(gridColor), lineWidth: 0.5)
        }

        // 边框
        let rectPath = Path(roundedRect: chartArea, cornerRadius: 4)
        context.stroke(rectPath, with: .color(axisColor), lineWidth: 1)
    }

    // MARK: - Y 轴标签

    private func drawYAxisLabels(context: GraphicsContext, chartArea: CGRect, yMin: Double, yMax: Double) {
        let horizontalLines = 5
        let font = Font.system(size: 11, weight: .regular)

        for i in 0...horizontalLines {
            let ratio = 1.0 - Double(i) / Double(horizontalLines)
            let value = yMin + (yMax - yMin) * ratio
            let y = chartArea.minY + chartArea.height * CGFloat(i) / CGFloat(horizontalLines)

            let text = Text(String(format: "%.1f", value))
                .font(font)
                .foregroundColor(.secondary)

            context.draw(text, at: CGPoint(x: chartArea.minX - 8, y: y), anchor: .trailing)
        }

        // Y 轴单位
        let unitText = Text("kg")
            .font(.system(size: 11, weight: .medium))
            .foregroundColor(.secondary)
        context.draw(unitText, at: CGPoint(x: chartArea.minX - 8, y: chartArea.minY - 8), anchor: .bottomTrailing)
    }

    // MARK: - X 轴标签

    private func drawXAxisLabels(context: GraphicsContext, chartArea: CGRect, pointCount: Int) {
        let font = Font.system(size: 11, weight: .regular)
        let verticalLines = 10

        // X 轴显示时间（秒），采样率 1000Hz
        let totalSeconds = Double(pointCount) / 1000.0

        for i in 0...verticalLines {
            let ratio = Double(i) / Double(verticalLines)
            let x = chartArea.minX + chartArea.width * CGFloat(ratio)
            let timeValue = totalSeconds * ratio
            let label = String(format: "%.1fs", timeValue)

            let text = Text(label)
                .font(font)
                .foregroundColor(.secondary)

            context.draw(text, at: CGPoint(x: x, y: chartArea.maxY + 8), anchor: .top)
        }
    }

    // MARK: - 曲线绘制

    private func drawCurve(
        context: GraphicsContext,
        chartArea: CGRect,
        data: [Double],
        color: Color,
        yMin: Double,
        yMax: Double,
        pointCount: Int
    ) {
        guard data.count > 1 else { return }

        let count = data.count
        let yRange = yMax - yMin
        guard yRange > 0 else { return }

        var path = Path()

        // 数据点映射到画布坐标
        let stepX = chartArea.width / CGFloat(max(count - 1, 1))

        for i in 0..<count {
            let x = chartArea.maxX - stepX * CGFloat(count - 1 - i)
            let normalizedY = (data[i] - yMin) / yRange
            let clampedY = max(0, min(1, normalizedY))
            let y = chartArea.maxY - chartArea.height * CGFloat(clampedY)

            if i == 0 {
                path.move(to: CGPoint(x: x, y: y))
            } else {
                path.addLine(to: CGPoint(x: x, y: y))
            }
        }

        context.stroke(path, with: .color(color), lineWidth: 1.5)

        // 绘制当前值标记点
        if let lastValue = data.last {
            let normalizedY = (lastValue - yMin) / yRange
            let clampedY = max(0, min(1, normalizedY))
            let y = chartArea.maxY - chartArea.height * CGFloat(clampedY)
            let point = CGPoint(x: chartArea.maxX, y: y)

            let circle = Path(ellipseIn: CGRect(x: point.x - 3, y: point.y - 3, width: 6, height: 6))
            context.fill(circle, with: .color(color))
        }
    }

    // MARK: - 零线绘制

    private func drawZeroLine(context: GraphicsContext, chartArea: CGRect, yMin: Double, yMax: Double) {
        guard yMin < 0 && yMax > 0 else { return }

        let yRange = yMax - yMin
        let normalizedY = (0 - yMin) / yRange
        let y = chartArea.maxY - chartArea.height * CGFloat(normalizedY)

        var path = Path()
        path.move(to: CGPoint(x: chartArea.minX, y: y))
        path.addLine(to: CGPoint(x: chartArea.maxX, y: y))

        context.stroke(path, with: .color(Color.red.opacity(0.4)), style: StrokeStyle(lineWidth: 1, dash: [4, 4]))
    }

    // MARK: - 图例

    private func drawLegend(context: GraphicsContext, size: CGSize, mode: DeviceMode) {
        let font = Font.system(size: 12, weight: .medium)
        let items: [(String, Color)]

        switch mode {
        case .single:
            items = [("力值", singleColor)]
        case .dual:
            items = [
                ("合力", combinedColor),
                ("主设备", mainColor),
                ("副设备", secondaryColor)
            ]
        default:
            return
        }

        let itemWidth: CGFloat = 80
        let totalWidth = CGFloat(items.count) * itemWidth
        let startX = (size.width - totalWidth) / 2
        let y: CGFloat = 6

        for (index, item) in items.enumerated() {
            let x = startX + CGFloat(index) * itemWidth

            // 色块
            let rect = CGRect(x: x, y: y, width: 12, height: 12)
            let roundedRect = Path(roundedRect: rect, cornerRadius: 2)
            context.fill(roundedRect, with: .color(item.1))

            // 文字
            let text = Text(item.0)
                .font(font)
                .foregroundColor(.primary)
            context.draw(text, at: CGPoint(x: x + 18, y: y + 6), anchor: .leading)
        }
    }

    // MARK: - 空状态

    private func drawEmptyState(context: GraphicsContext, chartArea: CGRect) {
        let text = Text("等待数据...")
            .font(.system(size: 16, weight: .medium))
            .foregroundColor(.secondary)
        context.draw(text, at: CGPoint(x: chartArea.midX, y: chartArea.midY), anchor: .center)
    }
}
