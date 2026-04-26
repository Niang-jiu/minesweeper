import cv2
import numpy as np
import time
import os
import keyboard
import mss
import win32gui
import win32api
import win32con
import threading
from multiprocessing import Process, freeze_support

# ==========================================
# 🎯 系統全域配置 & 啟動設定
# ==========================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

GAME_INSTANCES = [
    (366, 1091),   # 預設盤面 1
    #(2087, 1091),  # 預設盤面 2
]

GRID_W, GRID_H = 68, 38
TOTAL_MINES = 5

IMAGE_MAP = {
    '?': ['grid.png'], 
    'Flag': ['flag.png'], 
    '0': ['num_0.png'], 
    '1': ['num_1.png'], 
    '2': ['num_2.png'], 
    '3': ['num_3.png'], 
    '4': ['num_4.png'], 
    '5': ['num_5.png'],
    'REPLAY': ['replay_btn.png']
}

print(f"🕵️ 程式目前認為自己的家在： {SCRIPT_DIR}")
TEMPLATES = {}
for label, filenames in IMAGE_MAP.items():
    TEMPLATES[label] = [] 
    for filename in filenames:
        abs_path = os.path.join(SCRIPT_DIR, filename)
        if os.path.exists(abs_path):
            img_data = np.fromfile(abs_path, dtype=np.uint8)
            img = cv2.imdecode(img_data, cv2.IMREAD_COLOR)
            if img is not None: 
                TEMPLATES[label].append(img)

# ==========================================
# ⚙️ 基礎工具函數
# ==========================================
def safe_match_conf(src, tmpl):
    try:
        if src is None or tmpl is None: return 0.0
        if src.shape[0] < tmpl.shape[0] or src.shape[1] < tmpl.shape[1]: return 0.0
        return cv2.minMaxLoc(cv2.matchTemplate(src, tmpl, cv2.TM_CCOEFF_NORMED))[1]
    except Exception:
        return 0.0

def classify_cell_image(scr, default_label="?"):
    max_conf, best_label = 0.0, default_label
    for label, tmpl_list in TEMPLATES.items():
        if label == 'REPLAY': continue
        for tmpl in tmpl_list:
            val = safe_match_conf(scr, tmpl)
            if val > max_conf:
                max_conf, best_label = val, label
    threshold = (0.88 if best_label.isdigit() else (0.90 if best_label == 'Flag' else 0.85))
    return (best_label if max_conf >= threshold else default_label), max_conf

# ==========================================
# 🧠 正統旋轉門演算法生成器
# ==========================================
def revolving_door(n, k):
    if k == 0:
        return [[]]
    if n == k:
        return [list(range(n))]
    
    res = []
    res.extend(revolving_door(n-1, k))
    res.extend([c + [n-1] for c in reversed(revolving_door(n-1, k-1))])
    return res

# ==========================================
# 🎮 多開物件導向封裝 (Process-Safe)
# ==========================================
class MinesweeperGame:
    def __init__(self, game_id, base_x, base_y):
        self.game_id = game_id
        self.base_x = base_x
        self.base_y = base_y
        self.fixed_centers = [[(base_x + c * GRID_W, base_y + r * GRID_H) for c in range(5)] for r in range(5)]
        
        self.combos_queue = [] 
        self.combo_index = 0
        self.error_streak = 0
        self.local_sct = None
        self.last_seen_time = time.time() 
        self.progress_file = os.path.join(SCRIPT_DIR, f"progress_board_{self.game_id}.txt")

    def log(self, message):
        print(f"[盤面 {self.game_id}] {message}")

    def bg_click(self, pos, wait_after=0.0):
        x, y = int(pos[0]), int(pos[1])
        try:
            hwnd = win32gui.WindowFromPoint((x, y))
            if hwnd:
                client_pos = win32gui.ScreenToClient(hwnd, (x, y))
                lparam = win32api.MAKELONG(client_pos[0], client_pos[1])
                win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
                win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
                
                safe_lparam = win32api.MAKELONG(10, 10)
                win32gui.SendMessage(hwnd, win32con.WM_MOUSEMOVE, 0, safe_lparam)
        except Exception as e:
            self.log(f"點擊異常: {e}")
            
        if wait_after > 0:
            time.sleep(wait_after)

    def detect_cell_label(self, r, c):
        cx, cy = self.fixed_centers[r][c]
        monitor_box = {"top": int(cy-20), "left": int(cx-35), "width": 70, "height": 40}
        try:
            scr = np.array(self.local_sct.grab(monitor_box))[:, :, :3]
            label, _ = classify_cell_image(scr, default_label="?")
            return label
        except Exception:
            return "?"

    def force_set_state_no_toggle(self, r, c, target_label):
        while True:
            current_label = self.detect_cell_label(r, c)
            if current_label == target_label: 
                return True
            
            cx, cy = self.fixed_centers[r][c]
            self.bg_click((cx, cy), wait_after=0.1)
            
            start_t = time.time()
            while time.time() - start_t < 1.0:
                check_label = self.detect_cell_label(r, c)
                if check_label == target_label:
                    return True
                
                status, _ = self.get_grid_state()
                if status == "REPLAY":
                    return "REPLAY"
                time.sleep(0.05)
                    
            self.log(f"⚠️ 卡頓對抗：({r}, {c}) 尚未變為 {target_label}，重新點擊...")

    def get_grid_state(self):
        top_y = max(0, self.base_y - 100)
        left_x = max(0, self.base_x - 100)
        monitor = {"top": int(top_y), "left": int(left_x), "width": 600, "height": 500}
        
        try:
            img = self.local_sct.grab(monitor)
            game_roi = np.array(img)[:, :, :3]
        except Exception:
            return "ERROR", None
        
        if 'REPLAY' in TEMPLATES and TEMPLATES['REPLAY']:
            rep_conf = safe_match_conf(game_roi, TEMPLATES['REPLAY'][0])
            if rep_conf > 0.8:
                loc = cv2.minMaxLoc(cv2.matchTemplate(game_roi, TEMPLATES['REPLAY'][0], cv2.TM_CCOEFF_NORMED))[3]
                abs_x = left_x + loc[0] + 20
                abs_y = top_y + loc[1] + 10
                return "REPLAY", (abs_x, abs_y)

        grid = [["?" for _ in range(5)] for _ in range(5)]
        for r in range(5):
            for c in range(5):
                cx, cy = self.fixed_centers[r][c]
                rel_cx = int(cx - left_x)
                rel_cy = int(cy - top_y)
                
                roi = game_roi[max(0, rel_cy-20):rel_cy+20, max(0, rel_cx-35):rel_cx+35]
                if roi.size == 0: continue
                best_label, _ = classify_cell_image(roi, default_label="?")
                grid[r][c] = best_label
                
        return "GRID", grid

    def process_tick(self):
        status, data = self.get_grid_state()
        
        if status == "ERROR":
            self.error_streak += 1
            if self.error_streak % 5 == 0:
                self.log(f"⚠️ 畫面讀取失敗 ({self.error_streak} 次)，等待恢復...")
            return
            
        self.error_streak = 0
        
        if status == "REPLAY":
            self.log("🟢 結算畫面偵測！清除進度並重開局。")
            self.bg_click(data, wait_after=2.0)
            self.combos_queue = []
            self.combo_index = 0
            if os.path.exists(self.progress_file):
                try: os.remove(self.progress_file)
                except: pass
            return

        if status == "GRID":
            grid = data
            valid_cells = [(r, c) for r in range(5) for c in range(5) if grid[r][c] in ['?', 'Flag']]
            
            if len(valid_cells) < TOTAL_MINES:
                self.log("⚠️ 盤面未知格子太少，請確保在狀態良好的盤面上執行！")
                time.sleep(1)
                return

            if not self.combos_queue:
                self.log(f"🧠 產生旋轉門路徑... 目標格子數: {len(valid_cells)} 取 {TOTAL_MINES}")
                indices_combos = revolving_door(len(valid_cells), TOTAL_MINES)
                self.combos_queue = [[valid_cells[i] for i in combo] for combo in indices_combos]
                
                actual_flags = set((r, c) for r in range(5) for c in range(5) if grid[r][c] == 'Flag')
                loaded_from_file = False
                
                if os.path.exists(self.progress_file):
                    try:
                        with open(self.progress_file, "r") as f:
                            saved_idx = int(f.read().strip())
                            if 0 <= saved_idx < len(self.combos_queue):
                                self.combo_index = saved_idx
                                self.log(f"💾 讀取存檔成功！直接從第 {self.combo_index} 組繼續。")
                                loaded_from_file = True
                    except Exception as e:
                        self.log(f"⚠️ 讀取存檔失敗 ({e})，切換至視覺辨識。")
                
                if not loaded_from_file:
                    if len(actual_flags) > 0:
                        self.log("🔍 無存檔紀錄，啟動視覺反推：正在比對盤面尋找最吻合進度...")
                        best_idx = 0
                        min_diff = 999
                        for idx, combo in enumerate(self.combos_queue):
                            diff = len(actual_flags ^ set(combo))
                            if diff == 0:
                                best_idx = idx
                                break
                            if diff < min_diff:
                                min_diff = diff
                                best_idx = idx
                                
                        self.combo_index = best_idx
                        self.log(f"⏭️ 盤面辨識接關成功！鎖定進度為第 {self.combo_index} 組。")
                    else:
                        self.combo_index = 0
                        self.log(f"✅ 乾淨盤面！共 {len(self.combos_queue)} 種組合，開始爆破。")

            if self.combo_index < len(self.combos_queue):
                actual_flags = set((r, c) for r in range(5) for c in range(5) if grid[r][c] == 'Flag')
                target_flags = set(self.combos_queue[self.combo_index])

                to_pull = actual_flags - target_flags
                to_plant = target_flags - actual_flags

                if not to_pull and not to_plant:
                    self.combo_index += 1
                    try:
                        with open(self.progress_file, "w") as f:
                            f.write(str(self.combo_index))
                    except:
                        pass
                        
                    if self.combo_index % 500 == 0:
                        self.log(f"🏃 窮舉進度：已完成 {self.combo_index} / {len(self.combos_queue)} 組...")
                    return

                for r, c in to_pull:
                    if self.force_set_state_no_toggle(r, c, '?') == "REPLAY": return

                for r, c in to_plant:
                    if self.force_set_state_no_toggle(r, c, 'Flag') == "REPLAY": return

                self.log(f"⏳ 觀察盤面判定狀態 (最多等待 0.15 秒)...")
                check_start = time.time()
                hit_win = False
                
                while time.time() - check_start < 0.15:
                    status_check, _ = self.get_grid_state()
                    if status_check == "REPLAY":
                        hit_win = True
                        break
                    time.sleep(0.1)

                if hit_win:
                    self.log("🏆 成就達成，進入結算畫面！")
                    return

# ==========================================
# 🚀 145秒純後台定時點擊執行緒
# ==========================================
def auto_refresh_loop():
    """純獨立的定時點擊器 (純後台版)，每 145 秒點擊指定座標"""
    refresh_targets = [
       (94,63)
    ]
    
    time.sleep(145) 
    
    while True:
        print("\n🔄 [系統通知] 145秒已到，執行純後台定時點擊...")
        
        for i, (x, y) in enumerate(refresh_targets):
            if x == 0 and y == 0:
                continue
                
            try:
                hwnd = win32gui.WindowFromPoint((x, y))
                if hwnd:
                    client_pos = win32gui.ScreenToClient(hwnd, (x, y))
                    lparam = win32api.MAKELONG(client_pos[0], client_pos[1])
                    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
                    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
            except Exception as e:
                print(f"⚠️ 點擊異常: {e}")
            
            time.sleep(1.0) 
            
        print("✅ 重新載入完畢，繼續倒數 145 秒...")
        time.sleep(145) 

# ==========================================
# 🚀 多進程控制器
# ==========================================
def brain_worker(game_id, base_x, base_y):
    game = MinesweeperGame(game_id, base_x, base_y)
    with mss.mss() as sct:
        game.local_sct = sct 
        
        game.log("⏳ 啟動大腦進程，等待 3 秒讓遊戲畫面穩定...")
        time.sleep(3.0) 
        game.log("🟢 緩衝結束，正式開始截圖判斷！")
        
        while True:
            game.process_tick()
            time.sleep(0.05) 

# ==========================================
# 🏁 主程式進入點
# ==========================================
if __name__ == "__main__":
    freeze_support() 
    os.system('cls' if os.name == 'nt' else 'clear')
    
    print("☕ 程式已載入。[成就模式：終極完成版]")
    print("⚠️ 警告：請確保遊戲內的「旗子模式」已經手動開啟！程式不會幫你切換。")
    print("✅ 已恢復 145 秒獨立執行緒：後台定時點擊重載，對抗遊戲假死。")
    print("✅ 已啟用雙重接關機制：優先讀取文字檔，無檔案則自動『視覺反推盤面』接軌。")

    # 啟動定時點擊執行緒
    refresh_thread = threading.Thread(target=auto_refresh_loop, daemon=True)
    refresh_thread.start()

    processes = []
    for i, (x, y) in enumerate(GAME_INSTANCES):
        p = Process(target=brain_worker, args=(i+1, x, y))
        p.daemon = True 
        p.start()
        processes.append(p)

    try:
        while True:
            if keyboard.is_pressed('esc'):
                print("\n🛑 偵測到 ESC 鍵！主控端發送終止訊號。")
                for p in processes:
                    p.terminate()
                os._exit(0)
            time.sleep(0.1)
    except KeyboardInterrupt:
        print("\n🛑 程式手動終止。")
        for p in processes:
            p.terminate()
        os._exit(0)