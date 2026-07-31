//
//  ControlPanelView.swift
//  rexforce
//
//  Created by 陈学瀚 on 2026/7/28.
//

import SwiftUI
import CoreBluetooth

/// 左侧控制面板
///
/// 包含连接/断开按钮、归零按钮、设备信息显示和实时读数。
struct ControlPanelView: View {

    @ObservedObject var bleManager: BLEManager
    @ObservedObject var dataStore: ForceDataStore

    var body: some View {
        ScrollView {
            VStack(spacing: 20) {
                // 应用标题
                appHeader

                // 连接控制
                connectionSection

                // 实时读数
                readingsSection

                // 操作按钮
                actionsSection

                // 设备信息
                deviceInfoSection
            }
            .padding(.horizontal, 20)
            .padding(.vertical, 24)
        }
        .background(Color(.secondarySystemBackground))
        .frame(maxWidth: .infinity)
    }

    // MARK: - 应用标题

    private var appHeader: some View {
        VStack(spacing: 6) {
            Image(systemName: "scalemass")
                .font(.system(size: 36))
                .foregroundStyle(.blue)

            Text("RexForce")
                .font(.title2)
                .fontWeight(.bold)

            Text("测力数据采集")
                .font(.subheadline)
                .foregroundColor(.secondary)
        }
        .frame(maxWidth: .infinity)
        .padding(.bottom, 4)
    }

    // MARK: - 连接控制

    private var connectionSection: some View {
        VStack(spacing: 12) {
            connectionStatusBadge

            // 连接/断开按钮
            if bleManager.connectionState == .connected {
                Button(action: { bleManager.disconnect() }) {
                    Label("断开连接", systemImage: "antenna.radiowaves.left.and.right.slash")
                        .font(.headline)
                        .frame(maxWidth: .infinity)
                        .padding(.vertical, 14)
                }
                .buttonStyle(.borderedProminent)
                .tint(.red)
                .disabled(bleManager.connectionState == .disconnecting)
            } else if bleManager.connectionState == .scanning {
                Button(action: { bleManager.stopScanning() }) {
                    HStack {
                        ProgressView()
                            .tint(.white)
                        Text("正在搜索设备...")
                            .font(.headline)
                    }
                    .frame(maxWidth: .infinity)
                    .padding(.vertical, 14)
                }
                .buttonStyle(.borderedProminent)
                .tint(.orange)
            } else if bleManager.connectionState == .connecting {
                Button(action: {}) {
                    HStack {
                        ProgressView()
                            .tint(.white)
                        Text("正在连接...")
                            .font(.headline)
                    }
                    .frame(maxWidth: .infinity)
                    .padding(.vertical, 14)
                }
                .buttonStyle(.borderedProminent)
                .tint(.blue)
                .disabled(true)
            } else if bleManager.connectionState == .disconnecting {
                Button(action: {}) {
                    HStack {
                        ProgressView()
                            .tint(.white)
                        Text("正在断开...")
                            .font(.headline)
                    }
                    .frame(maxWidth: .infinity)
                    .padding(.vertical, 14)
                }
                .buttonStyle(.borderedProminent)
                .tint(.gray)
                .disabled(true)
            } else {
                Button(action: { bleManager.startScanning() }) {
                    Label("连接设备", systemImage: "antenna.radiowaves.left.and.right")
                        .font(.headline)
                        .frame(maxWidth: .infinity)
                        .padding(.vertical, 14)
                }
                .buttonStyle(.borderedProminent)
                .tint(.blue)
            }

            // 错误信息
            if case .error(let message) = bleManager.connectionState {
                Text(message)
                    .font(.caption)
                    .foregroundColor(.red)
                    .multilineTextAlignment(.center)
            }

            // 扫描到的设备列表
            if bleManager.connectionState == .scanning && !bleManager.discoveredDevices.isEmpty {
                VStack(spacing: 8) {
                    Text("发现设备")
                        .font(.caption)
                        .foregroundColor(.secondary)
                        .frame(maxWidth: .infinity, alignment: .leading)

                    ForEach(bleManager.discoveredDevices, id: \.identifier) { device in
                        Button(action: { bleManager.connect(to: device) }) {
                            HStack {
                                Image(systemName: "antenna.radiowaves.left.and.right")
                                    .foregroundColor(.blue)
                                Text(device.name ?? "未知设备")
                                    .font(.subheadline)
                                Spacer()
                                Image(systemName: "chevron.right")
                                    .font(.caption)
                                    .foregroundColor(.secondary)
                            }
                            .padding(.horizontal, 12)
                            .padding(.vertical, 10)
                            .background(Color(.tertiarySystemBackground))
                            .cornerRadius(8)
                        }
                        .buttonStyle(.plain)
                    }
                }
            }
        }
    }

    // MARK: - 连接状态徽标

    private var connectionStatusBadge: some View {
        HStack(spacing: 6) {
            Circle()
                .fill(statusColor)
                .frame(width: 8, height: 8)

            Text(statusText)
                .font(.caption)
                .foregroundColor(.secondary)
        }
        .padding(.horizontal, 12)
        .padding(.vertical, 6)
        .background(Color(.tertiarySystemBackground))
        .cornerRadius(20)
    }

    private var statusColor: Color {
        switch bleManager.connectionState {
        case .connected:    return .green
        case .scanning:     return .orange
        case .connecting:   return .blue
        case .disconnecting: return .gray
        case .error:        return .red
        default:            return .gray
        }
    }

    private var statusText: String {
        switch bleManager.connectionState {
        case .idle:          return "未连接"
        case .scanning:      return "搜索中"
        case .connecting:    return "连接中"
        case .connected:     return "已连接"
        case .disconnecting: return "断开中"
        case .disconnected:  return "已断开"
        case .error(let msg): return msg
        }
    }

    // MARK: - 实时读数

    private var readingsSection: some View {
        VStack(spacing: 12) {
            SectionHeader(title: "实时读数")

            if dataStore.mode == .dual {
                // 双机模式：三列读数
                HStack(spacing: 8) {
                    ReadingCard(
                        title: "合力",
                        value: dataStore.currentCombinedForce,
                        color: .blue
                    )
                    ReadingCard(
                        title: "主设备",
                        value: dataStore.currentMainForce,
                        color: .orange
                    )
                    ReadingCard(
                        title: "副设备",
                        value: dataStore.currentSecondaryForce,
                        color: .purple
                    )
                }
            } else {
                // 单机模式：单列读数
                ReadingCard(
                    title: "力值",
                    value: dataStore.currentMainForce,
                    color: .blue,
                    large: true
                )
            }
        }
    }

    // MARK: - 操作按钮

    private var actionsSection: some View {
        VStack(spacing: 12) {
            SectionHeader(title: "操作")

            // 归零按钮
            Button(action: { dataStore.zero() }) {
                Label("归零", systemImage: "arrow.counterclockwise.circle")
                    .font(.headline)
                    .frame(maxWidth: .infinity)
                    .padding(.vertical, 12)
            }
            .buttonStyle(.bordered)
            .tint(.blue)
            .disabled(bleManager.connectionState != .connected || dataStore.pointCount == 0)

            // 清空数据按钮
            Button(action: { dataStore.clearData(keepOffset: true) }) {
                Label("清空曲线", systemImage: "trash")
                    .font(.subheadline)
                    .frame(maxWidth: .infinity)
                    .padding(.vertical, 10)
            }
            .buttonStyle(.bordered)
            .tint(.secondary)
            .disabled(dataStore.pointCount == 0)
        }
    }

    // MARK: - 设备信息

    private var deviceInfoSection: some View {
        VStack(spacing: 8) {
            SectionHeader(title: "设备信息")

            InfoRow(label: "工作模式", value: dataStore.mode.displayName)
            InfoRow(label: "采样率", value: dataStore.sampleRate > 0 ? String(format: "%.0f Hz", dataStore.sampleRate) : "—")
            InfoRow(label: "数据点数", value: "\(dataStore.pointCount) / \(ForceDataStore.maxPoints)")
            InfoRow(label: "总采样数", value: "\(dataStore.totalSampleCount)")

            if dataStore.isZeroed {
                InfoRow(label: "归零状态", value: "已归零")
            }
        }
    }
}

// MARK: - 子视图

private struct SectionHeader: View {
    let title: String

    var body: some View {
        Text(title)
            .font(.subheadline)
            .fontWeight(.semibold)
            .foregroundColor(.secondary)
            .frame(maxWidth: .infinity, alignment: .leading)
    }
}

private struct ReadingCard: View {
    let title: String
    let value: Double
    let color: Color
    var large: Bool = false

    var body: some View {
        VStack(spacing: 4) {
            Text(title)
                .font(.caption2)
                .foregroundColor(.secondary)

            Text(String(format: "%.1f", value))
                .font(large ? .system(size: 36, weight: .bold, design: .rounded) : .title3.monospacedDigit())
                .fontWeight(.bold)
                .foregroundColor(color)

            Text("kg")
                .font(.caption2)
                .foregroundColor(.secondary)
        }
        .frame(maxWidth: .infinity)
        .padding(.vertical, large ? 16 : 12)
        .background(Color(.tertiarySystemBackground))
        .cornerRadius(12)
    }
}

private struct InfoRow: View {
    let label: String
    let value: String

    var body: some View {
        HStack {
            Text(label)
                .font(.caption)
                .foregroundColor(.secondary)
            Spacer()
            Text(value)
                .font(.caption)
                .fontWeight(.medium)
                .foregroundColor(.primary)
        }
        .padding(.horizontal, 12)
        .padding(.vertical, 8)
        .background(Color(.tertiarySystemBackground))
        .cornerRadius(8)
    }
}
