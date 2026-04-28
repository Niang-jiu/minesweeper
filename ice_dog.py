import cv2
import numpy as np
import time
import os
import random
import mss
import pyautogui
import keyboard

# ==========================================
# 設定區
# ==========================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# 🔍 你說要給的「座標框」，自己把座標填好 (x, y, 寬, 高)
SEARCH_ROI = {"top": 800, "left": 400, "width": 120, "height": 120}

# 🖱️ 當 10 秒都沒看到按鈕時，要點擊的 A 座標
COORD_A = (433, 975)

# 🖼️ 圖檔名稱 (請你自己去截圖，把「兌獎」跟「再玩」的按鈕切下來存檔)
IMG_CLAIM = 'claim_btn.png'   # 兌獎按鈕
IMG_REPLAY = 'replay_btn2.png' # 再玩按鈕

# ==========================================
# ⚙️ 資源載入
# ==========================================
def load_template(filename):
    abs_path = os.path.join(SCRIPT_DIR, filename)
    if not os.path.exists(abs_path):
        print(f"🛑 喂！找不到圖片 {filename}，你是不是又忘記截圖了？")
        return None
    # 避免中文路徑爛掉的讀取法
    img_data = np.fromfile(abs_path, dtype=np.uint8)
    return cv2.imdecode(img_data, cv2.IMREAD_COLOR)

templates = {
    'claim': load_template(IMG_CLAIM),
    'replay': load_template(IMG_REPLAY)
}

if any(v is None for v in templates.values()):
    print("程式終止。先把圖檔準備好再來找我。")
    os._exit(0)

# ==========================================
# 🧠 核心邏輯
# ==========================================
def random_wait():
    """隨機等待 1 +- 0.5 秒 (0.5 ~ 1.5秒)"""
    wait_time = random.uniform(0.5, 1.5)
    time.sleep(wait_time)

def execute_fallback_sequence():
    """執行 10 秒沒看到按鈕的備用方案"""
    print(f"\n[系統] 已經 10 秒沒看到任何按鈕，執行輸入指令序列...")
    
    # 1. 點擊 A 座標
    print(f"👉 點擊 A 座標 {COORD_A}")
    pyautogui.click(COORD_A[0], COORD_A[1])
    random_wait()
    
    # 2. Ctrl + V
    print("📋 貼上 (Ctrl + V)")
    pyautogui.hotkey('ctrl', 'v')
    random_wait()
    
    # 3. Enter
    print("↩️ 按下 Enter")
    pyautogui.press('enter')
    random_wait()
    
    print("[系統] 序列執行完畢，繼續監控畫面。\n")

def main():
    print("🚀 啟動！要終止請隨時按 ESC。")
    
    sct = mss.mss()
    last_seen_time = time.time()
    
    while True:
        # 緊急煞車
        if keyboard.is_pressed('esc'):
            print("\n🛑 偵測到 ESC，收工！")
            break
            
        # 擷取你指定的區域
        sct_img = sct.grab(SEARCH_ROI)
        scr = np.array(sct_img)[:, :, :3]
        
        button_found = False
        
        # 依序尋找「兌獎」或「再玩」
        for btn_name, tmpl in templates.items():
            if scr.shape[0] < tmpl.shape[0] or scr.shape[1] < tmpl.shape[1]:
                continue
                
            res = cv2.matchTemplate(scr, tmpl, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
            
            if max_val > 0.85: # 信心度閥值
                button_found = True
                last_seen_time = time.time() # 更新最後看到的時間
                
                # 換算回全螢幕的絕對座標 (中心點)
                center_x = SEARCH_ROI["left"] + max_loc[0] + (tmpl.shape[1] // 2)
                center_y = SEARCH_ROI["top"] + max_loc[1] + (tmpl.shape[0] // 2)
                
                # 加入 +-3 的隨機座標偏移
                offset_x = random.randint(-3, 3)
                offset_y = random.randint(-3, 3)
                target_x = center_x + offset_x
                target_y = center_y + offset_y
                
                action_name = "兌獎" if btn_name == 'claim' else "再玩"
                print(f"🎯 找到 [{action_name}] 按鈕！準備點擊座標: ({target_x}, {target_y}) (偏移: {offset_x:>2}, {offset_y:>2})")
                
                # 看到圖後等 0.2 +- 0.05 秒 (0.15 ~ 0.25秒)
                time.sleep(random.uniform(0.15, 0.25))
                
                # 點擊加上偏移後的目標
                pyautogui.click(target_x, target_y)
                
                # 鼠標移到隨機避嫌位置
                rand_x = random.randint(200, 400)
                rand_y = random.randint(200, 400)
                pyautogui.moveTo(rand_x, rand_y)
                
                time.sleep(0.5) # 點完稍微等一下 UI 反應
                break # 點了一個就先跳出迴圈重新截圖，避免同畫面重複點
                
        # 檢查是不是已經超過 10 秒沒看到按鈕了
        if not button_found:
            time_since_last_seen = time.time() - last_seen_time
            if time_since_last_seen > 10.0:
                execute_fallback_sequence()
                # 執行完後重置計時器，不然它會瘋狂連續觸發
                last_seen_time = time.time()
                
        time.sleep(0.1) # 讓 CPU 喘口氣

if __name__ == "__main__":
    main()
