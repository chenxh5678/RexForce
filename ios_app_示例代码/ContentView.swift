//
//  ContentView.swift
//  rexforce
//
//  Created by 陈学瀚 on 2026/7/28.
//

import SwiftUI

/// 主界面
///
/// 横屏布局：左 1/3 为操作区，右 2/3 为曲线图。
struct ContentView: View {

    @StateObject private var bleManager = BLEManager()
    @StateObject private var dataStore = ForceDataStore()

    var body: some View {
        GeometryReader { geometry in
            HStack(spacing: 0) {
                // 左侧操作区（1/3）
                ControlPanelView(bleManager: bleManager, dataStore: dataStore)
                    .frame(width: geometry.size.width / 3)

                // 分割线
                Divider()

                // 右侧曲线图（2/3）
                ForceChartView(dataStore: dataStore)
                    .frame(maxWidth: .infinity)
                    .frame(maxHeight: .infinity)
                    .padding(12)
                    .background(Color(.systemBackground))
            }
        }
        .background(Color(.systemBackground))
        .onAppear {
            setupBLECallbacks()
        }
        .onDisappear {
            if bleManager.connectionState == .connected {
                bleManager.disconnect()
            }
        }
    }

    /// 配置 BLE 数据回调
    private func setupBLECallbacks() {
        bleManager.onDataReceived = { result in
            dataStore.appendSamples(result)
        }
    }
}

#Preview(traits: .landscapeLeft) {
    ContentView()
}
