//
//  BLEConstants.swift
//  rexforce
//
//  Created by 陈学瀚 on 2026/7/28.
//

import Foundation
import CoreBluetooth

/// BLE 通信常量
enum BLEConstants {
    // MARK: - GATT UUID
    /// 服务 UUID
    static let serviceUUID = CBUUID(string: "0000FFE5-0000-1000-8000-00805F9A34FB")
    /// TX 特征（设备 → 客户端，Notify）
    static let txCharacteristicUUID = CBUUID(string: "0000FFE4-0000-1000-8000-00805F9A34FB")
    /// RX 特征（客户端 → 设备，Write）
    static let rxCharacteristicUUID = CBUUID(string: "0000FFE9-0000-1000-8000-00805F9A34FB")

    // MARK: - 帧格式
    /// 固定帧长度
    static let frameLength = 244
    /// 数据区长度
    static let dataFieldLength = 240
    /// 帧头 SLAVE
    static let frameHeader: UInt8 = 0x01
    /// 单机模式功能码
    static let funcSingle: UInt8 = 0x03
    /// 双机模式功能码
    static let funcDual: UInt8 = 0x05

    // MARK: - 设备名称
    /// BLE 广播名称前缀
    static let deviceNamePrefix = "Force-"

    // MARK: - 控制指令
    /// 恢复默认比例系数
    static let cmdRestoreDefault: [UInt8]     = [0x54, 0xF8]
    /// 保存当前为默认值
    static let cmdSaveDefault: [UInt8]        = [0x54, 0xCC]
    /// 请求温度原始值
    static let cmdRequestTemperature: [UInt8] = [0x54, 0xDD]
    /// 主动断开 BLE 连接
    static let cmdDisconnect: [UInt8]         = [0x54, 0xDE]
    /// 校准因子指令前缀
    static let cmdCalibratePrefix: [UInt8]    = [0x54, 0xAA]
}
