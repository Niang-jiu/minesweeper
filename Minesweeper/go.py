import cv2
import numpy as np
import time
import os
import keyboard
import itertools
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
    #(366, 394),   # 預設盤面 3
    #(2087, 394),   # 預設盤面 4
]

GRID_W, GRID_H = 68, 38
TOTAL_MINES = 5
RISK_ZONES = [(3, 1), (4, 1)]

IMAGE_MAP = {
    '?': ['grid.png'], 
    'Flag': ['flag.png'], 
    '0': ['num_0.png'], 
    '1': ['num_1.png'], 
    '2': ['num_2.png'], 
    '3': ['num_3.png'], 
    '4': ['num_4.png'], 
    '5': ['num_5.png'],
    'REPLAY': ['replay_btn.png'], 
    'FLAG_OFF_BTN': ['flag_mode1.png'], 
    'FLAG_ON_BTN': ['flag_mode2.png']
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
        if label in ['REPLAY', 'FLAG_OFF_BTN', 'FLAG_ON_BTN']: continue
        for tmpl in tmpl_list:
            val = safe_match_conf(scr, tmpl)
            if val > max_conf:
                max_conf, best_label = val, label
    threshold = (0.88 if best_label.isdigit() else (0.90 if best_label == 'Flag' else 0.85))
    return (best_label if max_conf >= threshold else default_label), max_conf

# ==========================================
# 🧠 演算法核心
# ==========================================
def backtrack_solve(grid, current_mines, current_safes):
    unks = [(r, c) for r in range(5) for c in range(5) if grid[r][c] == '?' and (r,c) not in current_mines and (r,c) not in current_safes]
    real_flags = set((r, c) for r in range(5) for c in range(5) if grid[r][c] == 'Flag')
    all_known_mines = real_flags.union(current_mines)
    mines_needed = TOTAL_MINES - len(all_known_mines)
    if mines_needed < 0 or mines_needed > len(unks) or not unks or len(unks) > 25: return set(), set(), []
        
    valid_configs = []
    for combo in itertools.combinations(unks, mines_needed):
        assumed_mines = all_known_mines.union(set(combo))
        is_valid = True
        for r in range(5):
            for c in range(5):
                if grid[r][c] in '012345': 
                    num = int(grid[r][c])
                    nbs = [(r+dr, c+dc) for dr in [-1,0,1] for dc in [-1,0,1] if (dr!=0 or dc!=0) and 0<=r+dr<5 and 0<=c+dc<5]
                    if sum(1 for n in nbs if n in assumed_mines) != num:
                        is_valid = False; break
            if not is_valid: break
        if is_valid: valid_configs.append(set(combo))
            
    if not valid_configs: return set(), set(), []
    must_be_mines = set.intersection(*valid_configs)
    must_be_safes = set(unks) - set.union(*valid_configs)
    return must_be_mines, must_be_safes, valid_configs

def solve_logic(grid):
    mines, safes, valid_configs = set(), set(), [] 
    changed = True
    while changed:
        changed = False
        info = []
        for r in range(5):
            for c in range(5):
                if grid[r][c] in '012345': 
                    num = int(grid[r][c])
                    nbs = [(r+dr, c+dc) for dr in [-1,0,1] for dc in [-1,0,1] if (dr!=0 or dc!=0) and 0<=r+dr<5 and 0<=c+dc<5]
                    unks = [p for p in nbs if grid[p[0]][p[1]] == '?' and p not in safes and p not in mines]
                    flgs = sum(1 for p in nbs if grid[p[0]][p[1]] == 'Flag' or p in mines)
                    needed = num - flgs
                    if needed < 0: continue 
                    if needed == 0 and unks:
                        for p in unks: safes.add(p); changed = True
                    elif needed == len(unks) and unks:
                        for p in unks: mines.add(p); changed = True
                    if unks and needed > 0: info.append({'set': set(unks), 'needed': needed})

        for i in range(len(info)):
            for j in range(len(info)):
                if i == j: continue
                if info[i]['set'].issubset(info[j]['set']):
                    diff_set = info[j]['set'] - info[i]['set']
                    diff_need = info[j]['needed'] - info[i]['needed']
                    if diff_need < 0: continue
                    if diff_set:
                        if diff_need == 0:
                            for p in diff_set: 
                                if p not in safes: safes.add(p); changed = True
                        elif diff_need == len(diff_set):
                            for p in diff_set: 
                                if p not in mines: mines.add(p); changed = True

    real_flags = set((r, c) for r in range(5) for c in range(5) if grid[r][c] == 'Flag')
    all_mines = real_flags.union(mines)
    unks_total = [(r, c) for r in range(5) for c in range(5) if grid[r][c] == '?' and (r, c) not in all_mines and (r, c) not in safes]

    if len(all_mines) == TOTAL_MINES and unks_total:
        for p in unks_total: safes.add(p)
    elif len(all_mines) + len(unks_total) == TOTAL_MINES and unks_total:
        for p in unks_total: mines.add(p)

    conflict = safes.intersection(mines)
    if not conflict and not mines and not safes:
        m_ex, s_ex, v_conf = backtrack_solve(grid, mines, safes)
        valid_configs = v_conf 
        if m_ex or s_ex:
            mines.update(m_ex)
            safes.update(s_ex)

    if safes.intersection(mines): return None, None, []
    return mines, safes, valid_configs

# ==========================================
# 🎮 多開物件導向封裝 (Process-Safe)
# ==========================================
class MinesweeperGame:
    def __init__(self, game_id, base_x, base_y):
        self.game_id = game_id
        self.base_x = base_x
        self.base_y = base_y
        
        self.flag_btn = (base_x , base_y + 190)
        self.fixed_centers = [[(base_x + c * GRID_W, base_y + r * GRID_H) for c in range(5)] for r in range(5)]
        
        self.last_grid_str = ""
        self.did_cleanup = False 
        self.risk_zone_cleared = False 
        self.error_streak = 0
        self.local_sct = None

    def log(self, message):
        print(f"[盤面 {self.game_id}] {message}")

    def bg_click(self, pos, wait_after=0.0):
        """✨ 後台虛擬點擊，踹開游標防擋視線"""
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

    def is_flag_on(self):
        monitor_flag = {"top": int(self.flag_btn[1]-60), "left": int(self.flag_btn[0]-60), "width": 120, "height": 120}
        try:
            scr = np.array(self.local_sct.grab(monitor_flag))[:, :, :3]
            conf_on = safe_match_conf(scr, TEMPLATES['FLAG_ON_BTN'][0])
            conf_off = safe_match_conf(scr, TEMPLATES['FLAG_OFF_BTN'][0])
            return conf_on > conf_off
        except Exception:
            return False

    def set_flag_mode(self, want_on):
        while self.is_flag_on() != want_on:
            self.log(f"🔄 切換：{'開啟' if want_on else '關閉'}旗子模式")
            self.bg_click(self.flag_btn, wait_after=2.0)

    def detect_cell_label(self, r, c):
        cx, cy = self.fixed_centers[r][c]
        monitor_box = {"top": int(cy-20), "left": int(cx-35), "width": 70, "height": 40}
        try:
            scr = np.array(self.local_sct.grab(monitor_box))[:, :, :3]
            label, _ = classify_cell_image(scr, default_label="?")
            return label
        except Exception:
            return "?"

    def click_and_watch(self, r, c, old_label):
        cx, cy = self.fixed_centers[r][c]
        monitor_box = {"top": int(cy-20), "left": int(cx-35), "width": 70, "height": 40}
        
        self.bg_click((cx, cy), wait_after=0.0)
        
        start_t = time.time()
        check_interval = 0
        while True:
            scr = np.array(self.local_sct.grab(monitor_box))[:, :, :3]
            best_label, _ = classify_cell_image(scr, default_label=old_label)
            
            if best_label != old_label:
                # ✨ 格子變化後，稍微等一下動畫浮現，再確認一次是不是結算了
                time.sleep(0.15) 
                status, _ = self.get_grid_state()
                if status == "REPLAY": return "REPLAY"
                return True
                
            check_interval += 1
            if check_interval % 5 == 0:
                status, _ = self.get_grid_state()
                if status == "REPLAY":
                    return "REPLAY" # ✨ 收到結算畫面，直接丟出信號
                    
            if time.time() - start_t > 3.0: 
                return False 
            time.sleep(0.05)

    def force_set_state(self, r, c, target_label, max_retries=3):
        current_label = self.detect_cell_label(r, c)
        if current_label == target_label: return True
        
        if target_label == 'Flag':
            self.set_flag_mode(True)
            old_label = current_label if current_label != '?' else '?'
        elif target_label == '?':
            self.set_flag_mode(True if current_label == 'Flag' else False)
            old_label = current_label
        else:
            return False
            
        for attempt in range(max_retries):
            res = self.click_and_watch(r, c, old_label)
            if res == "REPLAY": 
                return "REPLAY" # ✨ 果斷中斷
            if res == True: 
                return True
            self.log(f"⚠️ 第 {attempt+1} 次確認失敗，重試中...")
            if self.detect_cell_label(r, c) == target_label:
                return True
        return False

    def get_grid_state(self):
        """✨ 擷取自己的盤面區域"""
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
                self.log(f"⚠️ 畫面讀取失敗 ({self.error_streak} 次)，退避...")
            return
            
        self.error_streak = 0
        
        if status == "REPLAY":
            self.log("🟢 點擊重開局...")
            self.bg_click(data, wait_after=2.0)
            self.set_flag_mode(False)
            self.did_cleanup = False
            self.risk_zone_cleared = False
            self.last_grid_str = "?" * 25
            return

        if status == "GRID":
            grid = data
            num_flags = sum(row.count('Flag') for row in grid)
            unks_left = [(r, c) for r in range(5) for c in range(5) if grid[r][c] == '?']

            if len(unks_left) == 25:
                self.set_flag_mode(False)

            mines, safes, valid_configs = solve_logic(grid) 
            all_known_mines = set((r, c) for r in range(5) for c in range(5) if grid[r][c] == 'Flag').union(mines if mines else set())
            
            if num_flags >= TOTAL_MINES and len(all_known_mines) >= TOTAL_MINES and not safes and unks_left:
                if self.did_cleanup: return
                self.log("⚠️ 旗子達上限，大掃除！")
                self.set_flag_mode(True)
                for r in range(5):
                    for c in range(5):
                        if grid[r][c] == 'Flag':
                            self.click_and_watch(r, c, 'Flag')
                self.set_flag_mode(False)
                self.did_cleanup = True 
                return

            if mines is None:
                current_grid_str = ''.join(''.join(row) for row in grid)
                if current_grid_str == self.last_grid_str:
                    self.log("⚠️ 畫面矛盾，等待特效...")
                self.last_grid_str = current_grid_str
                return
                
            self.last_grid_str = ''.join(''.join(row) for row in grid)
            unflagged_mines = sorted([m for m in mines if grid[m[0]][m[1]] != 'Flag'], key=lambda x: (x[0], x[1]))

            if safes:
                target = sorted(list(safes))[0]
                if target in RISK_ZONES and not self.risk_zone_cleared:
                    self.risk_zone_cleared = True; return 
                self.set_flag_mode(False)
                self.log(f"👉 安全區 ({target[0]}, {target[1]})")
                self.click_and_watch(target[0], target[1], '?')
                self.risk_zone_cleared = False 
                return

            elif unflagged_mines:
                target = unflagged_mines[0]
                if target in RISK_ZONES and not self.risk_zone_cleared:
                    self.risk_zone_cleared = True; return 
                self.set_flag_mode(True)
                self.log(f"🚩 插旗 ({target[0]}, {target[1]})")
                self.click_and_watch(target[0], target[1], '?')
                self.risk_zone_cleared = False 
                return
                
            elif valid_configs and 0 < len(valid_configs) < 10:
                self.log(f"🧪 差分窮舉！ {len(valid_configs)} 組...")
                hit_win, network_failed = False, False
                guessed_flags = set() 
                
                for i, combo in enumerate(valid_configs):
                    target_flags = set(combo)
                    flags_to_pull = guessed_flags - target_flags
                    flags_to_plant = target_flags - guessed_flags
                    
                    for r, c in flags_to_pull:
                        res = self.force_set_state(r, c, '?')
                        if res == "REPLAY": hit_win = True; network_failed = False; break
                        if not res: network_failed = True; break
                        guessed_flags.remove((r, c))
                    if network_failed or hit_win: break
                    
                    for r, c in flags_to_plant:
                        res = self.force_set_state(r, c, 'Flag')
                        if res == "REPLAY": hit_win = True; network_failed = False; break
                        if not res: network_failed = True; break
                        guessed_flags.add((r, c))
                    if network_failed or hit_win: break
                        
                    time.sleep(0.3)
                    s2, _ = self.get_grid_state()
                    if s2 == "REPLAY":
                        self.log("🎉 命中正確組合！")
                        hit_win = True; break
                            
                if network_failed or (not hit_win and guessed_flags):
                    self.log(f"⚠️ 復原 {len(guessed_flags)} 根實驗旗子...")
                    for r, c in list(guessed_flags):
                        self.force_set_state(r, c, '?') 
                self.risk_zone_cleared = False
                return

            else:
                valid_corners = [c for c in [(0,0), (0,4), (4,0), (4,4)] if grid[c[0]][c[1]] == '?']
                target = valid_corners[0] if valid_corners else None
                if not target:
                    unks = [(r, c) for r in range(5) for c in range(5) if grid[r][c] == '?']
                    if unks: target = unks[0]
                
                if target:
                    if target in RISK_ZONES and not self.risk_zone_cleared:
                        self.risk_zone_cleared = True; return
                    self.set_flag_mode(False)
                    self.log(f"🎲 盲猜 ({target[0]}, {target[1]})")
                    self.click_and_watch(target[0], target[1], '?')
                    self.risk_zone_cleared = False

# ==========================================
# 🚀 145秒純後台定時點擊執行緒
# ==========================================
def auto_refresh_loop():
    """純獨立的定時點擊器 (純後台版)，每 145 秒點擊指定座標"""
    # 👇 大少爺，把你要點的頻道或按鈕座標 (x, y) 填進來
    refresh_targets = [
        #(80,59),  # 區塊 1 座標
        #(1815,62),  # 區塊 2 座標
        #(80, 756),  # 區塊 3 (如果有就取消註解)
       #(1815, 756),  # 區塊 4 (如果有就取消註解)
       (94,63)
    ]
    
    # 先等第一個 145 秒再開始動作
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
                    print(f"👉 對區塊 {i+1} 座標 ({x}, {y}) 發送後台點擊。")
            except Exception as e:
                print(f"⚠️ 點擊異常: {e}")
            
            time.sleep(1.0) # 點完等一秒再換下一個區塊
            
        print("✅ 全區塊後台點擊完畢，繼續倒數 145 秒...")
        time.sleep(145) # 重新倒數

# ==========================================
# 🚀 多進程控制器
# ==========================================
def brain_worker(game_id, base_x, base_y):
    """🧠 獨立大腦進程：享有專屬的 CPU 核心與記憶體空間"""
    game = MinesweeperGame(game_id, base_x, base_y)
    with mss.mss() as sct:
        game.local_sct = sct 
        while True:
            game.process_tick()
            time.sleep(0.05) 

# ==========================================
# 🏁 主程式進入點
# ==========================================
if __name__ == "__main__":
    freeze_support() 
    os.system('cls' if os.name == 'nt' else 'clear')
    
    print("☕ 程式已載入。這下你不用擔心滑鼠被綁架，也不用怕它鞭屍了。")
    print("🚀 啟動！『真實多進程架構』已上膛。由作業系統完美分配核心。")

    # ✨ 啟動純後台定時點擊執行緒
    print("🚀 啟動純後台定時點擊執行緒...")
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