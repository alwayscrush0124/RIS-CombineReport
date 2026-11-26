"""CT 自動合併（放在 CT 檢查室用的小程式）

只有自動模式。看 RIS「登錄作業清單」上選的那間檢查室，同一位病人有多筆 CT 的：
  ‧ 報告合併單號（合併序號）不一樣 → Ctrl 選好那幾筆 → 右鍵「合併報告」
  ‧ 影像單號不一樣 → Ctrl 選好 → 右鍵「影像合併」（影像還沒到就隔「刷新時間」再試）
  ‧ 已經一樣的就不用做
不點開病人、不按開始檢查、不按檢查完成，所以檢查室不會變。
其他規則跟 ct_merge.py 的自動模式一樣：全部是 CT 才做、不碰不用合併的醫令（ct_merge_no_merge.txt，
預設 *(HMC)）、報告醫師不一樣的不碰、電腦閒置才動手、RIS 跳出沒看過的視窗就停。

畫面上只能改「刷新時間」：每隔幾分鐘重新查詢一次清單、再試一次影像合併。
檢查室照 RIS 清單上現在選的那一間（按「開始」時讀）。

要跟 ct_merge.py 放在同一個資料夾。
"""

import ctypes
import os
import sys
import threading
import tkinter as tk
from tkinter import messagebox, ttk

# 免安裝版的 Python（embeddable）不會自動把程式所在的資料夾加進來，ct_merge.py 要自己找
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ct_merge as C  # noqa: E402

TITLE = "CT 自動合併"
MUTEX_NAME = "Local\\CT_AUTO_MERGE_SINGLE"
_mutex = None


class RoomApp(C.AutoApp):
    MODE = "CT 檢查室"

    def __init__(self):
        tk.Tk.__init__(self)            # 不用手動版的介面，自己排
        self.title(TITLE)
        self.geometry("1100x680")
        self._init_state()
        # 這些在 CT 檢查室版固定，不給改
        self.room_var = tk.StringVar()
        self.ask_each = tk.BooleanVar(value=False)
        self.cross_room = tk.BooleanVar(value=False)
        self.idle_var = tk.StringVar(value=str(C.IDLE_SEC))

        top = ttk.Frame(self, padding=(8, 8, 8, 4))
        top.pack(fill="x")
        self.btn_go = ttk.Button(top, text="開始", width=10, command=self.on_start)
        self.btn_go.pack(side="left")
        self.btn_stop = ttk.Button(top, text="停止", width=10, command=self.on_stop, state="disabled")
        self.btn_stop.pack(side="left", padx=(6, 20))
        ttk.Label(top, text="刷新時間").pack(side="left")
        self.retry_var = tk.StringVar(value=str(C.IMAGE_RETRY_MIN))
        self.spin = ttk.Spinbox(top, from_=1, to=60, width=4, textvariable=self.retry_var)
        self.spin.pack(side="left", padx=4)
        ttk.Label(top, text="分鐘（重新查詢清單、再試一次影像合併）").pack(side="left")
        self.lbl_room = ttk.Label(top, text="檢查室：讀取中…", foreground="#0b57d0")
        self.lbl_room.pack(side="right")

        self._build_body("RIS 開在「登錄作業清單」、選好這間檢查室，按「開始」")
        self.after(300, self._show_room)

    # ---------- 檢查室：照 RIS 清單上選的 ----------
    def _read_room(self):
        self.ris.connect()
        return C.norm(self.ris.read_filters()["room"])

    def _show_room(self):
        """開程式時在右上角顯示 RIS 清單現在是哪一間（背景讀，不點任何東西）。"""
        def job():
            try:
                room = self._read_room()
                text = f"檢查室：{room}" if room else "檢查室：RIS 清單選的是「全部」，請選這間"
            except Exception as e:
                text = "檢查室：讀不到 RIS 的清單"
                self.q.put(("log", f"讀不到 RIS 的清單：{e}"))
            self.q.put(("call", lambda: self.lbl_room.config(text=text)))
        threading.Thread(target=job, daemon=True).start()

    def _room(self):
        """按「開始」時：檢查室照 RIS 清單上現在選的那一間。"""
        try:
            room = self._read_room()
        except Exception as e:
            messagebox.showinfo(TITLE, f"讀不到 RIS 的「登錄作業清單」：\n{e}\n\n"
                                "請把 RIS 切到「登錄作業清單」再按「開始」。", parent=self)
            return ""
        if not room:
            messagebox.showinfo(TITLE, "RIS「登錄作業清單」的檢查室是空白（全部檢查室）。\n"
                                "請先在 RIS 選這間 CT 檢查室，再按「開始」。", parent=self)
            return ""
        self.room_var.set(room)
        self.lbl_room.config(text=f"檢查室：{room}")
        return room

    # ---------- 畫面上的字 ----------
    def _set_running(self, on):
        self.btn_go.config(state="disabled" if on else "normal")
        self.spin.config(state="disabled" if on else "normal")
        if not on:
            self.btn_stop.config(state="disabled")

    def intro(self):
        return ("CT 自動合併：同一位病人的多筆 CT，報告合併序號不一樣就右鍵「合併報告」，"
                "影像單號不一樣就右鍵「影像合併」。不點開病人、不按開始檢查，檢查室不會變。")

    def confirm_text(self, room, cross):
        n = self._int(self.retry_var, C.IMAGE_RETRY_MIN, 1, 60)
        return (f"檢查室：{room}\n刷新時間：每 {n} 分鐘\n\n"
                "程式會自動（同一位病人有 2 筆以上 CT 才做）：\n"
                "  ‧ 報告合併序號不一樣 → 右鍵「合併報告」\n"
                f"  ‧ 影像單號不一樣 → 右鍵「影像合併」（影像還沒到就每 {n} 分鐘再試）\n"
                "不點開病人、不按開始檢查、不按檢查完成，檢查室不會變。\n"
                f"鍵盤滑鼠 {C.IDLE_SEC} 秒沒人動才會動手，有人在用就等。\n\n"
                "確定要開始嗎？")


def already_running():
    """同一台電腦只能開一個（兩個一起按 RIS 會打架）。"""
    global _mutex
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.CreateMutexW.restype = ctypes.c_void_p
    _mutex = k32.CreateMutexW(None, False, MUTEX_NAME)
    return ctypes.get_last_error() == 183          # ERROR_ALREADY_EXISTS


def main():
    if already_running():
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(TITLE, "「CT 自動合併」已經開著了，請看一下工作列。")
        root.destroy()
        return
    RoomApp().mainloop()


if __name__ == "__main__":
    main()
