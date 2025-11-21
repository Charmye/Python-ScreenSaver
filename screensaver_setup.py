import os
import time
import subprocess
import logging
import pyautogui
import pyperclip  # 用於複製貼上文字
from datetime import datetime
import sys

# --- 設定 Log 與 截圖路徑 ---
BASE_DIR = os.path.join(os.getcwd(), "AutoSetup_Result")
SCREENSHOT_DIR = os.path.join(BASE_DIR, "Screenshots")
LOG_FILE = os.path.join(BASE_DIR, "setup_log.txt")

if not os.path.exists(SCREENSHOT_DIR):
    os.makedirs(SCREENSHOT_DIR)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.FileHandler(LOG_FILE, encoding='utf-8'), logging.StreamHandler()]
)

def log(msg):
    logging.info(msg)

# --- 全 UI 自動化操作 ---
def automate_gui_steps():
    log("--- 開始執行 UI 自動化設定 ---")
    
    # 1. 準備文字到剪貼簿
    target_text = "日本船主責任相互保険組合"
    pyperclip.copy(target_text)
    log(f"已將文字複製到剪貼簿: {target_text}")

    # 2. 打開螢幕保護程式設定視窗
    log("正在開啟螢幕保護程式設定視窗...")
    subprocess.Popen("control desk.cpl,,@screensaver")
    time.sleep(2.5) # 等待視窗完全載入

    # ==========================
    # 階段一：設定時間為 5 分鐘
    # ==========================
    # 視窗開啟時，預設焦點通常在「螢幕保護裝置下拉選單」
    # 我們按 Tab 鍵移動焦點：
    # Tab 1次 -> 設定(T)
    # Tab 2次 -> 預覽(V)
    # Tab 3次 -> 等待時間(W)
    
    log("正在設定時間為 5 分鐘...")
    # 保險起見，先點一下視窗中間確保焦點在視窗上
    # (這裡用 pyautogui.press 比較保險，不依賴滑鼠座標)
    
    # 按一次選擇3D
    pyautogui.press('down')
    time.sleep(1)
    
    # 連按 3 次 Tab 到達時間輸入框
    pyautogui.press('tab', presses=3, interval=0.3)
    time.sleep(0.5)
    
    # 輸入 5
    pyautogui.typewrite('5')
    log("時間已輸入 5")
    time.sleep(0.5)

    # ==========================
    # 階段二：設定 3D 文字
    # ==========================
    # 現在焦點在時間框，我們需要回到「設定(T)」按鈕
    # 方法：按 Shift+Tab (反向 Tab) 2次 回到「設定(T)」
    
    log("正在進入詳細設定...")
    pyautogui.hotkey('shift', 'tab')
    time.sleep(0.2)
    pyautogui.hotkey('shift', 'tab')
    time.sleep(0.5)
    
    # 此時焦點應該在「設定(T)」按鈕上，按下 Space 或 Enter 進入
    pyautogui.press('space')
    time.sleep(2.0) # 等待子視窗彈出 (這一塊容易卡頓，設久一點)

    # 進入子視窗後
    log("正在輸入文字...")
    
    # 1. 切換到「自訂文字」(Alt+C)
    pyautogui.hotkey('alt', 'c')
    time.sleep(0.8) 
    
    # 2. 【關鍵修正】按一下 Tab 進入輸入框
    pyautogui.press('tab', presses=4, interval=0.3)
    time.sleep(0.5)
    
    # 3. 清空舊文字 (Ctrl+A -> Del) 並貼上 (Ctrl+V)
    pyautogui.hotkey('ctrl', 'a') 
    time.sleep(0.1)
    pyautogui.press('delete')
    time.sleep(0.1)
    
    pyautogui.hotkey('ctrl', 'v')
    log("文字已貼上。")
    time.sleep(1.0)

    # 4. 按下確定 (Enter)
    pyautogui.press('enter')
    time.sleep(1.5)
    log("已關閉詳細設定，回到主畫面。")

    # ==========================
    # 階段三：勾選「恢復時顯示登入畫面」
    # ==========================
    # 剛從子視窗回來，焦點通常會回到「設定(T)」按鈕
    # 恢復選項(R) 通常在時間的後面
    # 路徑：OK -> Tab -> 設定(T) -> Tab -> 預覽(V) -> Tab -> 時間 -> Tab -> 恢復(R)
    
    log("正在勾選鎖定畫面...")
    pyautogui.press('tab', presses=4, interval=0.3)
    
    # 這裡很難判斷目前是否已勾選。
    # 最保險的方式是用 Alt+R (如果支援快捷鍵) 
    # 或者直接按 Space (切換勾選狀態)。
    # 假設預設沒勾，按一次 Space。如果原本有勾，會被取消。
    # 比較穩妥是用 registry 補強，或是這裡假設環境是乾淨的。
    # 為了展示，我們先按 Alt+R (這通常會直接聚焦並切換)
    pyautogui.hotkey('alt', 'r') 
    time.sleep(0.5)

    # ==========================
    # 階段四：截圖與存檔
    # ==========================
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    screenshot_path = os.path.join(SCREENSHOT_DIR, f"ScreenSaver_Result_{timestamp}.png")
    
    # 嘗試截圖，如果還報錯，會被 catch 抓到
    try:
        pyautogui.screenshot(screenshot_path)
        log(f"已保存截圖: {screenshot_path}")
    except Exception as e:
        log(f"截圖失敗 (請確認已安裝 Pillow): {e}")

    log("UI 設定流程結束。")

if __name__ == "__main__":
    try:
        automate_gui_steps()
        
        print("\n========== 執行完畢 ==========")
        print("請檢查視窗設定，並查看截圖資料夾。")
        input("按 Enter 鍵退出...")
        
    except Exception as e:
        log(f"發生嚴重錯誤: {e}")
        input("發生錯誤，按 Enter 鍵退出...")
