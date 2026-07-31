//
//  BLEManager.swift
//  rexforce
//
//  Created by 陈学瀚 on 2026/7/28.
//

import Foundation
import CoreBluetooth
import Combine

/// BLE 连接状态
enum BLEConnectionState: Equatable {
    case idle
    case scanning
    case connecting
    case connected
    case disconnecting
    case disconnected
    case error(String)
}

/// BLE 管理器
///
/// 负责 RexForce 测力主板的 BLE 扫描、连接、数据接收与指令发送。
@MainActor
final class BLEManager: NSObject, ObservableObject {

    // MARK: - Published 属性
    @Published var connectionState: BLEConnectionState = .idle
    @Published var discoveredDevices: [CBPeripheral] = []
    @Published var connectedDeviceName: String?
    @Published var deviceMode: DeviceMode = .unknown
    @Published var isScanning = false

    // MARK: - 数据回调
    /// 收到解析后的采样数据
    var onDataReceived: ((FrameParser.ParseResult) -> Void)?

    // MARK: - 私有属性
    private var centralManager: CBCentralManager!
    private var connectedPeripheral: CBPeripheral?
    private var connectedPeripheralDisplayName: String?
    private var txCharacteristic: CBCharacteristic?
    private var rxCharacteristic: CBCharacteristic?

    /// 帧接收缓冲区
    private var receiveBuffer = Data()
    /// 扫描超时定时器
    private var scanTimeoutTimer: Timer?
    /// 扫描超时时间（秒）
    private let scanTimeout: TimeInterval = 10

    // MARK: - 初始化

    override init() {
        super.init()
        centralManager = CBCentralManager(delegate: self, queue: .main)
    }

    // MARK: - 公开方法

    /// 开始扫描 Force-XXXX 设备
    func startScanning() {
        guard centralManager.state == .poweredOn else {
            connectionState = .error("蓝牙未开启")
            return
        }

        discoveredDevices = []
        isScanning = true
        connectionState = .scanning
        centralManager.scanForPeripherals(
            withServices: nil,
            options: [CBCentralManagerScanOptionAllowDuplicatesKey: false]
        )

        // 设置扫描超时
        scanTimeoutTimer?.invalidate()
        scanTimeoutTimer = Timer.scheduledTimer(withTimeInterval: scanTimeout, repeats: false) { [weak self] _ in
            guard let self else { return }
            Task { @MainActor [weak self] in
                guard let self else { return }
                if self.connectionState == .scanning {
                    self.stopScanning()
                    if self.discoveredDevices.isEmpty {
                        self.connectionState = .error("未找到 Force 设备，请确认设备已开机")
                    }
                }
            }
        }
    }

    /// 停止扫描
    func stopScanning() {
        centralManager.stopScan()
        scanTimeoutTimer?.invalidate()
        scanTimeoutTimer = nil
        isScanning = false
        if connectionState == .scanning {
            connectionState = .idle
        }
    }

    /// 连接指定设备
    func connect(to peripheral: CBPeripheral) {
        stopScanning()
        connectedPeripheral = peripheral
        connectedPeripheralDisplayName = peripheral.name
        peripheral.delegate = self
        connectionState = .connecting
        centralManager.connect(peripheral, options: nil)
    }

    /// 自动连接第一个发现的设备
    func connectToFirstDiscovered() {
        guard let first = discoveredDevices.first else {
            connectionState = .error("未发现设备")
            return
        }
        connect(to: first)
    }

    /// 断开连接
    ///
    /// 先发送主动断开指令，再取消连接。
    func disconnect() {
        guard let peripheral = connectedPeripheral else {
            connectionState = .idle
            return
        }

        connectionState = .disconnecting

        // 发送主动断开指令
        if let rxChar = rxCharacteristic {
            let data = Data(BLEConstants.cmdDisconnect)
            peripheral.writeValue(data, for: rxChar, type: .withResponse)
        }

        // 短暂延迟后取消连接，确保指令发送完成
        DispatchQueue.main.asyncAfter(deadline: .now() + 0.2) { [weak self] in
            self?.centralManager.cancelPeripheralConnection(peripheral)
        }
    }

    /// 发送校准因子
    /// - Parameters:
    ///   - measured: 当前测量值（kg）
    ///   - standard: 标准砝码值（kg）
    func sendCalibrationFactor(measured: Double, standard: Double) {
        guard let peripheral = connectedPeripheral,
              let rxChar = rxCharacteristic else { return }

        let factor = measured / standard
        let factorInt = UInt32((factor * 10000).rounded())

        var cmd = Data(BLEConstants.cmdCalibratePrefix)
        // 4 字节大端无符号整数
        cmd.append(UInt8((factorInt >> 24) & 0xFF))
        cmd.append(UInt8((factorInt >> 16) & 0xFF))
        cmd.append(UInt8((factorInt >> 8) & 0xFF))
        cmd.append(UInt8(factorInt & 0xFF))

        peripheral.writeValue(cmd, for: rxChar, type: .withResponse)
    }

    /// 恢复默认比例系数
    func restoreDefaultFactor() {
        sendCommand(BLEConstants.cmdRestoreDefault)
    }

    /// 保存当前为默认值
    func saveDefaultFactor() {
        sendCommand(BLEConstants.cmdSaveDefault)
    }

    // MARK: - 私有方法

    private func sendCommand(_ bytes: [UInt8]) {
        guard let peripheral = connectedPeripheral,
              let rxChar = rxCharacteristic else { return }
        peripheral.writeValue(Data(bytes), for: rxChar, type: .withResponse)
    }

    /// 处理接收到的原始 BLE 数据，累积并解析完整帧
    private func processReceivedData(_ data: Data) {
        receiveBuffer.append(data)

        while receiveBuffer.count >= BLEConstants.frameLength {
            // 查找帧头
            guard let headerIndex = receiveBuffer.firstIndex(of: BLEConstants.frameHeader) else {
                receiveBuffer.removeAll()
                break
            }

            // 移除帧头前的无效字节
            let invalidPrefixCount = receiveBuffer.distance(from: receiveBuffer.startIndex, to: headerIndex)
            if invalidPrefixCount > 0 {
                receiveBuffer.removeFirst(invalidPrefixCount)
            }

            // 检查是否有足够字节
            guard receiveBuffer.count >= BLEConstants.frameLength else { break }

            // 验证功能码
            let funcCodeIndex = receiveBuffer.index(receiveBuffer.startIndex, offsetBy: 1)
            let funcCode = receiveBuffer[funcCodeIndex]
            guard funcCode == BLEConstants.funcSingle || funcCode == BLEConstants.funcDual else {
                // 非有效功能码，跳过这个帧头字节继续搜索
                receiveBuffer.removeFirst(1)
                continue
            }

            // 提取完整帧
            let frame = receiveBuffer.prefix(BLEConstants.frameLength)
            receiveBuffer.removeFirst(BLEConstants.frameLength)

            // 解析帧
            if let result = FrameParser.parse(frame: Data(frame)) {
                if deviceMode != result.mode {
                    deviceMode = result.mode
                }
                onDataReceived?(result)
            }
        }
    }
}

// MARK: - CBCentralManagerDelegate

extension BLEManager: CBCentralManagerDelegate {

    func centralManagerDidUpdateState(_ central: CBCentralManager) {
        switch central.state {
        case .poweredOn:
            if connectionState == .error("蓝牙未开启") {
                connectionState = .idle
            }
        case .poweredOff:
            connectionState = .error("蓝牙未开启")
            isScanning = false
        case .unauthorized:
            connectionState = .error("未授权使用蓝牙")
        case .unsupported:
            connectionState = .error("设备不支持蓝牙")
        default:
            break
        }
    }

    func centralManager(_ central: CBCentralManager,
                        didDiscover peripheral: CBPeripheral,
                        advertisementData: [String: Any],
                        rssi RSSI: NSNumber) {
        let advertisedName = advertisementData[CBAdvertisementDataLocalNameKey] as? String
        let resolvedName = advertisedName ?? peripheral.name ?? ""
        guard resolvedName.localizedCaseInsensitiveContains("force") else { return }

        if connectedPeripheral?.identifier == peripheral.identifier {
            connectedPeripheralDisplayName = resolvedName
        }

        // 避免重复添加
        if !discoveredDevices.contains(where: { $0.identifier == peripheral.identifier }) {
            discoveredDevices.append(peripheral)
        }

        // 如果当前在扫描状态且只有一个设备，自动连接
        if connectionState == .scanning && discoveredDevices.count == 1 {
            connect(to: peripheral)
        }
    }

    func centralManager(_ central: CBCentralManager,
                        didConnect peripheral: CBPeripheral) {
        connectionState = .connected
        connectedDeviceName = connectedPeripheralDisplayName ?? peripheral.name
        receiveBuffer.removeAll()

        peripheral.discoverServices(nil)
    }

    func centralManager(_ central: CBCentralManager,
                        didFailToConnect peripheral: CBPeripheral,
                        error: Error?) {
        connectionState = .error("连接失败: \(error?.localizedDescription ?? "未知错误")")
        connectedPeripheral = nil
    }

    func centralManager(_ central: CBCentralManager,
                        didDisconnectPeripheral peripheral: CBPeripheral,
                        error: Error?) {
        connectedPeripheral = nil
        connectedPeripheralDisplayName = nil
        txCharacteristic = nil
        rxCharacteristic = nil
        connectedDeviceName = nil
        deviceMode = .unknown
        receiveBuffer.removeAll()

        if case .disconnecting = connectionState {
            connectionState = .disconnected
        } else if let error = error {
            connectionState = .error("连接断开: \(error.localizedDescription)")
        } else {
            connectionState = .disconnected
        }
    }
}

// MARK: - CBPeripheralDelegate

extension BLEManager: CBPeripheralDelegate {

    func peripheral(_ peripheral: CBPeripheral,
                    didDiscoverServices error: Error?) {
        if let error = error {
            connectionState = .error("发现服务失败: \(error.localizedDescription)")
            return
        }

        guard let services = peripheral.services, !services.isEmpty else {
            connectionState = .error("未找到任何服务")
            return
        }

        for service in services {
            peripheral.discoverCharacteristics(nil, for: service)
        }
    }

    func peripheral(_ peripheral: CBPeripheral,
                    didDiscoverCharacteristicsFor service: CBService,
                    error: Error?) {
        if let error = error {
            connectionState = .error("发现特征失败: \(error.localizedDescription)")
            return
        }

        for characteristic in service.characteristics ?? [] {
            if characteristic.uuid == BLEConstants.txCharacteristicUUID {
                txCharacteristic = characteristic
                // 订阅 Notify
                peripheral.setNotifyValue(true, for: characteristic)
            } else if characteristic.uuid == BLEConstants.rxCharacteristicUUID {
                rxCharacteristic = characteristic
            }
        }
    }

    func peripheral(_ peripheral: CBPeripheral,
                    didUpdateValueFor characteristic: CBCharacteristic,
                    error: Error?) {
        guard error == nil,
              characteristic.uuid == BLEConstants.txCharacteristicUUID,
              let data = characteristic.value else { return }

        processReceivedData(data)
    }

    func peripheral(_ peripheral: CBPeripheral,
                    didUpdateNotificationStateFor characteristic: CBCharacteristic,
                    error: Error?) {
        if let error = error {
            connectionState = .error("订阅通知失败: \(error.localizedDescription)")
        }
    }

    func peripheralIsReady(toSendWriteWithoutResponse peripheral: CBPeripheral) {
        // 可在此处理写就绪
    }
}
