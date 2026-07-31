//
//  ForceSample.swift
//  rexforce
//
//  Created by 陈学瀚 on 2026/7/28.
//

import Foundation

/// 解析后的单个采样点
struct ForceSample {
    /// 主设备力值（kg）
    let mainForce: Double
    /// 副设备力值（kg），单机模式下为 nil
    let secondaryForce: Double?

    /// 合力 = 主设备 + 副设备
    var combinedForce: Double {
        mainForce + (secondaryForce ?? 0)
    }

    /// 是否为双机模式数据
    var isDualMode: Bool {
        secondaryForce != nil
    }
}
