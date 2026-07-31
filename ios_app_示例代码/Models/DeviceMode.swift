//
//  DeviceMode.swift
//  rexforce
//
//  Created by 陈学瀚 on 2026/7/28.
//

import Foundation

/// 设备工作模式
enum DeviceMode: Equatable {
    /// 未知（未连接或未收到数据）
    case unknown
    /// 单机模式：仅发送主设备重量数据
    case single
    /// 双机模式：主设备与副设备数据合并发送
    case dual

    /// 功能码
    var funcCode: UInt8 {
        switch self {
        case .single: return 0x03
        case .dual:   return 0x05
        case .unknown: return 0x00
        }
    }

    /// 从功能码推断模式
    static func from(funcCode: UInt8) -> DeviceMode {
        switch funcCode {
        case 0x03: return .single
        case 0x05: return .dual
        default:   return .unknown
        }
    }

    /// 显示名称
    var displayName: String {
        switch self {
        case .unknown: return "未连接"
        case .single:  return "单机模式"
        case .dual:    return "双机模式"
        }
    }
}
