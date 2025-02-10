########################################
#                                      #
#        Made by ATGXicefires          #
#                                      #
########################################

import win32api
from screeninfo import get_monitors
import ctypes
from ctypes import windll, c_int, Structure, POINTER, c_uint32, byref
import time
import sys
from PySide6.QtWidgets import (QApplication, QLabel, QVBoxLayout, QWidget)
from PySide6.QtCore import QTimer

class SystemTrayIcon(object):
    def __init__(self, parent):
        self.parent = parent
        self.hwnd = int(self.parent.winId())  # 取得視窗控制代碼
        self.message_id = win32api.RegisterWindowMessage("TaskbarCreated")
        self.create_tray_icon()

    def create_tray_icon(self):
        # 建立系統托盤圖示 (這部分比較複雜，需要查閱 Windows API 文件)
        pass

    def show_window(self):
        ctypes.windll.user32.ShowWindow(self.hwnd, 1)  # 1 代表 SW_SHOW
        self.parent.show()

    def hide_window(self):
        ctypes.windll.user32.ShowWindow(self.hwnd, 0)  # 0 代表 SW_HIDE
        self.parent.hide()

    def process_messages(self):
        # 處理 Windows 訊息 (這部分比較複雜，需要查閱 Windows API 文件)
        pass

# 新增 Windows API 結構和常數
class DISPLAYCONFIG_PATH_INFO(Structure):
    _fields_ = []  # 簡化版結構，實際使用時需完整定義

class DISPLAYCONFIG_MODE_INFO(Structure):
    _fields_ = []  # 簡化版結構，實際使用時需完整定義

# 載入必要的 Windows API 函數
SetDisplayConfig = windll.user32.SetDisplayConfig
GetDisplayConfigBufferSizes = windll.user32.GetDisplayConfigBufferSizes

# Windows 常數定義
QDC_ALL_PATHS = 1
SDC_TOPOLOGY_INTERNAL = 0x00000002
SDC_APPLY = 0x00000080

def get_refresh_rate():
    try:
        # 使用 win32api 獲取刷新率
        device = win32api.EnumDisplayDevices(None, 0)
        settings = win32api.EnumDisplaySettings(device.DeviceName, -1)
        refresh_rate = getattr(settings, 'DisplayFrequency')
        return refresh_rate
    except Exception as e:
        print(f"獲取刷新率時發生錯誤: {str(e)}")
        return None

def get_available_modes(device_name):
    modes = []
    i = 0
    try:
        while True:
            mode = win32api.EnumDisplaySettings(device_name, i)
            modes.append({
                'width': mode.PelsWidth,
                'height': mode.PelsHeight,
                'freq': mode.DisplayFrequency,
                'settings': mode
            })
            i += 1
    except:
        pass
    return modes

def set_refresh_rate(refresh_rate):
    try:
        # 1. 先設定桌面模式
        device = win32api.EnumDisplayDevices(None, 0)
        current_settings = win32api.EnumDisplaySettings(device.DeviceName, -1)
        
        new_settings = win32api.EnumDisplaySettings(device.DeviceName, -1)
        new_settings.PelsWidth = current_settings.PelsWidth
        new_settings.PelsHeight = current_settings.PelsHeight
        new_settings.BitsPerPel = current_settings.BitsPerPel
        new_settings.DisplayFrequency = refresh_rate
        
        # 2. 同時設定信號模式
        num_paths = c_uint32(0)
        num_modes = c_uint32(0)
        
        # 獲取所需緩衝區大小
        GetDisplayConfigBufferSizes(QDC_ALL_PATHS, byref(num_paths), byref(num_modes))
        
        # 應用新的顯示配置
        flags = SDC_TOPOLOGY_INTERNAL | SDC_APPLY
        result1 = win32api.ChangeDisplaySettings(new_settings, 0x00000001)
        result2 = SetDisplayConfig(0, None, 0, None, flags)
        
        # 驗證更改是否生效
        time.sleep(2)  # 等待設定生效
        current = get_refresh_rate()
        
        if current == refresh_rate and result1 == 0 and result2 == 0:
            print("桌面模式和信號模式都已成功更新")
            return True
        else:
            if current != refresh_rate:
                print(f"警告：刷新率設定未完全生效，當前為 {current}Hz")
            if result1 != 0:
                print(f"桌面模式設定失敗，錯誤碼：{result1}")
            if result2 != 0:
                print(f"信號模式設定失敗，錯誤碼：{result2}")
            return False
            
    except Exception as e:
        print(f"設置刷新率時發生錯誤: {str(e)}")
        return False

def main():
    app = QApplication(sys.argv)
    window = QWidget()
    layout = QVBoxLayout()

    label = QLabel("開始監控螢幕刷新率...\n按下 Ctrl+C 可終止程式")
    layout.addWidget(label)

    window.setLayout(layout)
    window.show()

    tray_icon = SystemTrayIcon(window)  # 建立系統托盤圖示

    def update_label():
        current_rate = get_refresh_rate()
        label.setText(f"當前螢幕刷新率: {current_rate} Hz")

        if current_rate != 60:
            label.setText(label.text() + "\n正在將刷新率設定為 60Hz...")
            if set_refresh_rate(60):
                label.setText(label.text() + "\n成功將螢幕刷新率設定為 60Hz\n設定已永久儲存")
            else:
                label.setText(label.text() + "\n設定刷新率失敗，請檢查是否有管理員權限")
                # 顯示所有可用的刷新率
                device = win32api.EnumDisplayDevices(None, 0)
                modes = get_available_modes(device.DeviceName)
                current_resolution = f"{modes[0]['width']}x{modes[0]['height']}"
                label.setText(label.text() + f"\n當前解析度 {current_resolution} 可用的刷新率:")
                for mode in modes:
                    if (mode['width'] == modes[0]['width'] and 
                        mode['height'] == modes[0]['height']):
                        label.setText(label.text() + f"\n{mode['freq']}Hz")
        else:
            label.setText(label.text() + "\n螢幕刷新率維持在 60Hz")

        label.setText(label.text() + "\n等待下次檢查...")

    timer = QTimer()
    timer.timeout.connect(update_label)
    timer.start(10000)  # 每10秒更新一次

    sys.exit(app.exec())

if __name__ == "__main__":
    main()