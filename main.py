import win32api
from screeninfo import get_monitors
import ctypes
import time

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

def set_refresh_rate(refresh_rate):
    try:
        device = win32api.EnumDisplayDevices(None, 0)
        settings = win32api.EnumDisplaySettings(device.DeviceName, -1)
        
        # 修改刷新率設定
        settings.DisplayFrequency = refresh_rate
        
        # 0x00000001 代表 CDS_UPDATEREGISTRY (永久更改)
        # 0x00000000 代表臨時更改
        result = win32api.ChangeDisplaySettings(settings, 0x00000001)
        
        if result == 0:  # DISP_CHANGE_SUCCESSFUL
            return True
        return False
    except Exception as e:
        print(f"設置刷新率時發生錯誤: {str(e)}")
        return False

def main():
    current_rate = get_refresh_rate()
    print(f"當前螢幕刷新率: {current_rate} Hz")
    
    if current_rate != 60:
        print("正在將刷新率設定為 60Hz...")
        if set_refresh_rate(60):
            print("成功將螢幕刷新率設定為 60Hz")
            print("設定已永久儲存")
        else:
            print("設定刷新率失敗")
    else:
        print("螢幕刷新率已經是 60Hz")

if __name__ == "__main__":
    main()
