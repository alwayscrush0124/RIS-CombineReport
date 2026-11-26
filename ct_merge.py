"""
CT 多筆合併小工具（操作 RIS 畫面，盯著「登錄作業清單」自動做）

同一位病人有多筆 CT 時，RIS 要做的事：
  第一段（病人一出現「已報到」就做）：
    點兩下 →（RIS 問「是否將同報告醫師，報告合併？」按「是」）→ 多筆登錄作業
    → 開始檢查 → 確認報告醫師、報告合併單號都一樣 → 影像合併（影像到了才會成功）
    → 點分頁的 X 離開
  第二段（狀態「檢查中」、影像還沒合併的，每隔幾分鐘試一次）：
    點兩下 → 確認報告 → 影像合併 → 點 X 離開
  「檢查完成」不按，留給技術員。

只在電腦閒置的時候動（鍵盤滑鼠一段時間沒人動、RIS 停在登錄作業清單），
有人在用 RIS 就等下一輪，不跟人搶。

這些按鈕按下去就直接生效（不像班表要自己存檔），所以：
  - 只處理清單上「全部都是 CT」的病人。有非 CT（X 光等）的整位不碰，留給人工：
    非 CT 不能合併，也不能被「開始檢查」換到 CT 檢查室。
  - 點兩下後先確認選到的真的是這位，才回答 RIS 的問句。
  - 打開的多筆登錄作業要跟清單對得上（病歷號相同、照會單號一筆不多一筆不少、
    全部是 CT），對不上就停，一個按鈕都不按。
  - 按了開始檢查，檢查室如果被換掉（RIS 登入的檢查室不對）立刻停止監看。
    勾「跨檢查室使用」時（在 portable 等別間幫忙）：離開後到換過去的那間清單，
    點選這幾筆（Ctrl 多選）拖到右邊檢查室表格的原檢查室，再查詢確認回來了。
    放開前核對選取和目標，不對就拖回原處放開（等於沒做）並停止。
  - 報告醫師、報告合併單號不一樣：停下來讓人在 RIS 改，改好按「確定」繼續。
    （醫師要怎麼統一的規則之後再定，先跑順。）
  - 每次點擊前確認：RIS 在最前面、沒有跳出視窗、要點的位置真的是清單／工具列／
    分頁的 X。
  - 跳出視窗：認得的照規則按；沒看過、只有一個按鈕的直接關掉；
    沒看過的問句（是/否、確定/取消…）停下來不按。
  - 任何意料之外的狀況：停止監看、跳視窗告訴你，不會自己繼續。

記錄寫在 ct_merge_log.txt（只記照會單號、醫令、醫師，不記病人姓名和病歷號）。
"""
import collections
import functools
import ctypes
from _ctypes import COMError
import os
import queue
import re
import sys
import threading
import time
import traceback
from ctypes import wintypes
from datetime import datetime

import tkinter as tk
from tkinter import messagebox, ttk

import comtypes.client  # noqa: F401  （pywinauto 的 IUIA 需要）
from pywinauto import keyboard, mouse
from pywinauto.controls.hwndwrapper import HwndWrapper
from pywinauto.controls.win32_controls import ComboBoxWrapper
from pywinauto.uia_defines import IUIA

BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
LOG_FILE = os.path.join(BASE_DIR, "ct_merge_log.txt")
# 不用合併的醫令（一行一個，跟清單上的醫令名稱一樣；大小寫、多餘空白不管；* 代表任意文字；
# ! 開頭是例外 = 這個要合併）。病人的 CT 只要有一筆符合，整位不碰（不開始檢查、不合併）。
# 檔案不在就用下面的預設建一個。
NO_MERGE_FILE = os.path.join(BASE_DIR, "ct_merge_no_merge.txt")
NO_MERGE_HEADER = (
    "# CT 多筆合併小工具：不用合併的醫令\n"
    "# 一行一個，跟清單上的醫令名稱一樣（大小寫、多餘空白不管）。# 開頭的是說明。\n"
    "# 病人的 CT 只要有一筆符合，整位都不碰（不開始檢查、不合併）。\n"
    "# * 代表任意文字：*(HMC) = 所有 (HMC) 結尾的醫令\n"
    "# ! 開頭是例外，這個醫令要合併，例如：!CT-Chest(HMC)\n"
)
NO_MERGE_DEFAULT = (
    # 2026-11-26 使用者：健檢的低劑量肺癌篩檢＋冠狀動脈鈣化分數不用合併 → 所有 (HMC) 都不合併，
    # 有要合併的之後再加例外
    "*(HMC)",
)

MAIN_TITLE_RE = re.compile(r".*放射線資訊管理系統.*")
LIST_FORM, LIST_GRID = "fRegistryList", "B"     # 登錄作業清單、它的表格
ROOM_GRID = "U"                                  # 清單右邊的檢查室表格（代碼／檢查室名稱），拖曳放下的地方
MULTI_FORM = "fMultiRegistry"                    # 多筆登錄作業
MULTI_TAB_RE = re.compile(r"多筆登錄")           # 上方分頁的標題
CT_RE = re.compile(r"^\s*CT", re.I)
READY = "已報到"
EXAMINING = "檢查中"

# 清單要用到的欄位（用欄位名稱找，不靠順序）
NEED_COLS = ("照會單號", "病歷號", "姓名", "狀態", "醫令名稱", "合併序號", "影像單號", "報告醫師")

# 等待上限（秒）：條件一成立就往下走，不是固定睡這麼久
OPEN_WAIT = 25        # 點兩下到多筆登錄作業打開（中間可能有問句視窗）
STEP_WAIT = 20        # 開始檢查到狀態變成檢查中
DOCTOR_WAIT = 6       # 開始檢查後等報告醫師帶進來
IMAGE_WAIT = 12       # 影像合併到影像儲存單號變成一樣（影像沒到時按了不會有反應）
IMAGE_POPUP_HOLD = 8  # 「影像資訊」視窗最多留幾秒再關（合併是在它開著的時候完成的）
IMAGE_INFO_AFTER = 5  # 影像合併好之後，等「影像資訊」跳出來最多幾秒
CLOSE_WAIT = 10       # 點 X 到多筆登錄作業關掉
POLL = 0.25
GUARD_WAIT = 0.5      # 點擊前的檢查沒過，最多再等這麼久（視窗剛好在重畫）
SCROLL_TRIES = 5      # 病人在清單畫面外時，最多捲幾次
QUERY_WAIT = 2.0      # 按「查詢」之後等清單更新
DRAG_GONE_WAIT = 8    # 放下後等那幾筆從清單上消失
DRAG_TRIES = 3        # 拖曳被核對擋下來（例如剛好清單自動更新）時，重新查詢再拖，最多幾次

# 監看
LOOP_SEC = 5          # 每隔幾秒看一次清單
IDLE_SEC = 20         # 鍵盤滑鼠多久沒人動才動手（畫面上可以改）
IMAGE_RETRY_MIN = 1   # 影像還沒到的病人，隔幾分鐘再試（畫面上可以改；使用者要一分鐘）
REOPEN_COOLDOWN = 60  # 同一位剛處理完，至少隔這麼久才再打開（等清單自動更新）

# RIS 開多筆登錄作業時的問句（都按「是」）：
#   點兩下後先問「是否多筆登錄?」（是/否/取消）→ 是 = 帶出這位的多筆登錄作業
#     （2026-11-25 實際跑第一次才看到；使用者手動做時很快按掉，觀察程式沒錄到）
#   接著可能問「是否將同報告醫師,報告合併?」→ 是 = 同報告醫師的報告合併（使用者確認過）
MULTI_Q_RE = re.compile(r"是否多筆登錄")
MERGE_Q_RE = re.compile(r"同報告醫師.*報告合併")
# 影像還沒到時按影像合併，RIS 會說「無影像,無法合併!」（確定）——看到就不用再等
NO_IMAGE_RE = re.compile(r"無影像.*無法合併")
# 改過報告醫師再按「報告合併」，RIS 依序問（使用者示範時都按「是」，2026-11-25 錄到）：
#   「資料未存檔,是否要存?」（是/否/取消）
#   「所選擇的資料分別屬於不同報告，是否確定合併報告?」（是/否）
#   「是否合併成新的報告?」（是/否）→ 使用者說按「是」
#     （示範時也有跳，觀察程式只錄到視窗沒錄到字；結果報告合併單號全部變成第一筆的照會單號）
SAVE_Q_RE = re.compile(r"資料未存檔.*是否要存")
DIFF_REPORT_Q_RE = re.compile(r"分別屬於不同報告.*確定合併報告")
NEW_REPORT_Q_RE = re.compile(r"是否合併成新的報告")
# 有影像時按「影像合併」，RIS 先問「是否確定合併影像?」（是/否）→ 是（2026-11-25 有影像的測試病人實測；
# 沒影像時不會問，直接說「無影像,無法合併!」）
IMAGE_Q_RE = re.compile(r"是否確定合併影像")
# 伺服器／資料庫錯誤：不能當成一般提醒按掉（2026-11-25 開始檢查時跳過
# 「[aServer.SaveObject][SaveEntity]...[IDBRIS.XRAYBASIC」，只有「確定」）→ 停下來留給人看
ERROR_RE = re.compile(r"\[aServer|SaveObject|Exception|ORA-\d|SQL|錯誤|失敗", re.I)
IMAGE_INFO_TITLE = "影像資訊"


def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def clean(s):
    """表格裡空白會顯示成 (null)，當成空的。"""
    s = norm(s)
    return "" if s.lower() in ("(null)", "null") else s


def oneline(s, n=200):
    s = re.sub(r"\s+", " ", s or "").strip()
    return s if len(s) <= n else s[:n] + "…"


def is_ct(order):
    return bool(CT_RE.match(order or ""))


def same(values):
    """全部一樣、而且不是空的。"""
    vs = {clean(v) for v in values}
    return len(vs) == 1 and "" not in vs


# 報告醫師要統一成「body」那幾筆的醫師（使用者 2026-11-25 定的規則）。
# body 自己就不一樣時（實際遇過：pelvis、abdomen 一位，chest 另一位），
# 依 abdomen → pelvis → chest 的順序，用排最前面的部位的醫師（使用者選的）。
BODY_PARTS = (("abdomen", re.compile(r"abdom", re.I)),
              ("pelvis", re.compile(r"pelvi", re.I)),
              ("chest", re.compile(r"chest", re.I)))
BODY_RE = re.compile(r"abdom|pelvi|chest", re.I)


def is_body(order):
    return bool(BODY_RE.search(order or ""))


def target_doctor(exams):
    """報告醫師要統一成誰：abdomen → pelvis → chest，有的第一個部位的醫師。

    回傳 (醫師, None)；決定不了回 (None, 原因)，這時候停下來請人處理。
    """
    body = [e for e in exams if is_body(e.get("order"))]
    if not body:
        return None, "這位沒有 abdomen／pelvis／chest 的檢查，不知道要統一成哪位醫師"
    if any(not clean(e.get("doctor")) for e in body):
        return None, "abdomen／pelvis／chest 有一筆的報告醫師是空白"
    for part, rx in BODY_PARTS:
        docs = [clean(e.get("doctor")) for e in exams if rx.search(e.get("order") or "")]
        if not docs:
            continue
        if not same(docs):
            return None, f"有好幾筆 {part}，報告醫師本身就不一樣"
        return docs[0], None
    return None, "找不到 body 的報告醫師"


class StopRun(Exception):
    """要停止監看（訊息給人看）。"""


class SkipPatient(Exception):
    """這一輪先不做這位（還沒動到任何東西），下一輪再看。"""


class ManualPatient(Exception):
    """使用者說這位自己處理，這次監看不再碰。"""


class NotReady(Exception):
    """RIS 畫面還沒準備好（例如登錄作業清單沒開），監看時等一下再看。"""


# RIS 忙（例如剛按「是否多筆登錄」在載入）時讀畫面會逾時，或元件剛好在重畫不見了：
# 等一下再讀，不是錯誤（2026-11-25 監看時實際遇到逾時，整個監看停掉）
UIA_BUSY_HRESULTS = (-2146233083,   # 0x80131505 逾時（RIS 沒空回應）
                     -2147220991)   # 0x80040201 元件不見了（正在重畫）
UIA_BUSY_WAIT = 20


def uia_retry(fn):
    @functools.wraps(fn)
    def wrap(*a, **k):
        end = time.time() + UIA_BUSY_WAIT
        while True:
            try:
                return fn(*a, **k)
            except COMError as e:
                if e.hresult not in UIA_BUSY_HRESULTS:
                    raise
                if time.time() >= end:
                    raise StopRun(f"RIS 一直沒有回應（讀畫面逾時超過 {UIA_BUSY_WAIT} 秒），"
                                  "程式停下來。請看一下 RIS 是不是當掉了。")
                time.sleep(0.5)
    return wrap


# ----------------------------------------------------------------------
# Win32
# ----------------------------------------------------------------------
_u32 = ctypes.WinDLL("user32", use_last_error=True)
_k32 = ctypes.WinDLL("kernel32", use_last_error=True)
WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
_u32.EnumWindows.argtypes = [WNDENUMPROC, wintypes.LPARAM]
_u32.EnumChildWindows.argtypes = [wintypes.HWND, WNDENUMPROC, wintypes.LPARAM]
_u32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
_u32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
_u32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
_u32.GetWindowThreadProcessId.restype = wintypes.DWORD
for _f in ("IsWindowVisible", "IsWindowEnabled", "IsWindow"):
    getattr(_u32, _f).argtypes = [wintypes.HWND]
    getattr(_u32, _f).restype = wintypes.BOOL
_u32.IsChild.argtypes = [wintypes.HWND, wintypes.HWND]
_u32.IsChild.restype = wintypes.BOOL
_u32.GetForegroundWindow.restype = wintypes.HWND
_u32.WindowFromPoint.argtypes = [wintypes.POINT]
_u32.WindowFromPoint.restype = wintypes.HWND
_u32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
_u32.GetAncestor.restype = wintypes.HWND
_u32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
_u32.PostMessageW.restype = wintypes.BOOL
_u32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
_k32.GetTickCount.restype = wintypes.DWORD
_u32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
_u32.SendMessageW.restype = ctypes.c_ssize_t
_u32.GetClientRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
_u32.ClientToScreen.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.POINT)]
_u32.GetCursorPos.argtypes = [ctypes.POINTER(wintypes.POINT)]
_u32.GetDoubleClickTime.restype = wintypes.UINT
GA_ROOT = 2
BM_CLICK = 0x00F5
# 下拉選單打開後的清單（ComboLBox）：只用來捲動、讀位置
LB_GETTOPINDEX, LB_GETCOUNT, LB_SETTOPINDEX, LB_GETITEMHEIGHT = 0x018E, 0x018B, 0x0197, 0x01A1
CB_SHOWDROPDOWN = 0x014F


class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [("cbSize", wintypes.UINT), ("dwTime", wintypes.DWORD)]


_u32.GetLastInputInfo.argtypes = [ctypes.POINTER(LASTINPUTINFO)]


def last_input_tick():
    """最後一次有鍵盤／滑鼠輸入的時間（開機後的毫秒）。"""
    li = LASTINPUTINFO()
    li.cbSize = ctypes.sizeof(li)
    _u32.GetLastInputInfo(ctypes.byref(li))
    return li.dwTime


def tick_now():
    return _k32.GetTickCount()


def text_of(h):
    b = ctypes.create_unicode_buffer(1024)
    _u32.GetWindowTextW(h, b, 1024)
    return b.value


def class_of(h):
    b = ctypes.create_unicode_buffer(256)
    _u32.GetClassNameW(h, b, 256)
    return b.value


def pid_of(h):
    p = wintypes.DWORD()
    _u32.GetWindowThreadProcessId(h, ctypes.byref(p))
    return p.value


def rect_of_hwnd(h):
    r = wintypes.RECT()
    _u32.GetWindowRect(h, ctypes.byref(r))
    return (r.left, r.top, r.right, r.bottom)


def top_windows():
    hs = []
    cb = WNDENUMPROC(lambda h, _l: hs.append(h) or True)
    _u32.EnumWindows(cb, 0)
    return hs


def child_windows(h):
    hs = []
    cb = WNDENUMPROC(lambda c, _l: hs.append(c) or True)
    _u32.EnumChildWindows(h, cb, 0)
    return hs


def window_desc(h):
    if not h:
        return "（沒有視窗）"
    if pid_of(h) == os.getpid():
        return "這支小工具自己的視窗"
    return f"「{text_of(h) or '沒有標題'}」({class_of(h)})"


def center(rect):
    l, t, r, b = rect
    return (l + r) // 2, (t + b) // 2


class Popup:
    def __init__(self, h, title, cls, texts, buttons):
        self.h, self.title, self.cls = h, title, cls
        self.texts = texts          # 視窗裡的文字
        self.buttons = buttons      # [(hwnd, 按鈕文字)]

    @property
    def body(self):
        return " ".join(self.texts)


# ----------------------------------------------------------------------
# UIA
# ----------------------------------------------------------------------
_iuia = IUIA()
_U = _iuia.UIA_dll
P_RECT, P_CT, P_NAME, P_EN, P_AID, P_HWND, P_OFF = 30001, 30003, 30005, 30010, 30011, 30020, 30022
P_VAL, P_LVAL, P_LSTATE = 30045, 30093, 30096
CT_BUTTON, CT_TEXT, CT_MENUITEM = 50000, 50020, 50011
_cache = _iuia.iuia.CreateCacheRequest()
for _p in (P_RECT, P_CT, P_NAME, P_EN, P_AID, P_HWND, P_OFF, P_VAL, P_LVAL, P_LSTATE):
    _cache.AddProperty(_p)
_TRUE = _iuia.iuia.CreateTrueCondition()
CELL_RE = re.compile(r"^(.+?) 資料列 (\d+)$")
ROW_RE = re.compile(r"^資料列 (\d+)$")
STATE_SELECTED = 0x2


def _by_aid(aid):
    return _iuia.iuia.CreatePropertyCondition(P_AID, aid)


def _rect(r):
    return (int(r.left), int(r.top), int(r.right), int(r.bottom))


def _value(e):
    """格子／欄位的值。不退回用名稱 —— WinForms 欄位的名稱是旁邊的標籤字。"""
    for p in (P_VAL, P_LVAL):
        try:
            v = e.GetCachedPropertyValue(p)
            if v not in (None, ""):
                return str(v)
        except Exception:
            pass
    return ""


def _cached(e, prop, default=None):
    try:
        return e.GetCachedPropertyValue(prop)
    except Exception:
        return default


class Ris:
    """讀 RIS 畫面、點 RIS 的按鈕。每一個會動到 RIS 的動作之前都先檢查。"""

    def __init__(self, log=print):
        self.log = log
        self.main = None
        self.pid = None
        self.grid_rect = None
        self.grid_hwnd = None
        self.image_info = ""        # 最近一次「影像資訊」視窗的內容
        self._image_seen = None     # 「影像資訊」視窗第一次看到的時間
        self.our_tick = 0           # 我們自己最後一次送滑鼠的時間（跟人的輸入分開）
        self.human_tick = 0         # 人最後一次動鍵盤滑鼠的時間
        # 「是否將同報告醫師，報告合併?」要按是還是否：回 True = 是（由 Runner 設定）
        self.merge_decider = lambda: True
        self.merge_answered = None  # 最近一次回答（True 是 / False 否 / None 還沒問）
        self.no_image = False       # 看到「無影像,無法合併!」

    # ---------- 連線 ----------
    def connect(self):
        self.main = None
        for h in top_windows():
            if MAIN_TITLE_RE.match(text_of(h)):
                self.main, self.pid = h, pid_of(h)
                break
        if not self.main:
            raise StopRun("找不到 RIS 主視窗，請確認 RIS 開著。")

    def _root(self):
        return _iuia.iuia.ElementFromHandle(self.main)

    def activate(self):
        """把 RIS 叫到最前面（剛按完小工具的按鈕／對話框時要做）。"""
        try:
            HwndWrapper(self.main).set_focus()
        except Exception:
            pass
        time.sleep(0.3)

    def ris_in_front(self):
        fg = _u32.GetForegroundWindow()
        return bool(fg) and (fg == self.main or _u32.GetAncestor(fg, GA_ROOT) == self.main)

    # ---------- 閒置 ----------
    def human_idle_sec(self):
        """人多久沒動鍵盤滑鼠了（我們自己送的滑鼠不算）。"""
        li = last_input_tick()
        if li > self.our_tick + 500:        # 不是我們送的 → 是人
            self.human_tick = max(self.human_tick, li)
        return ((tick_now() - self.human_tick) & 0xFFFFFFFF) / 1000.0

    def reset_idle(self):
        self.human_tick = last_input_tick()

    # ---------- 讀畫面 ----------
    def read_list(self, tries=4):
        """讀「登錄作業清單」。回傳 [dict]，每列含 NEED_COLS 和
        _n（列號）、_rect（醫令名稱那格）、_offscreen、_sel（是否被選取）。

        清單會自動更新：剛好在更新的時候讀，會讀到缺欄位的列（2026-11-25 實測）。
        缺欄位就隔一秒重讀，連續幾次都缺才當作畫面改版停下來。"""
        for k in range(tries):
            out, missing = self._read_list_once()
            if not missing:
                return out
            if k < tries - 1:
                time.sleep(1.0)
        raise StopRun(f"清單少了這些欄位：{'、'.join(missing)}（重讀 {tries} 次都一樣）。"
                      "RIS 畫面可能改版了，為安全起見停止。")

    @uia_retry
    def _read_list_once(self):
        form = self._root().FindFirst(_U.TreeScope_Descendants, _by_aid(LIST_FORM))
        if not form:
            raise NotReady("找不到「登錄作業清單」，請把 RIS 切到那個畫面。")
        grid = form.FindFirst(_U.TreeScope_Descendants, _by_aid(LIST_GRID))
        if not grid:
            raise NotReady("找不到「登錄作業清單」的表格。")
        self.grid_rect = _rect(grid.CurrentBoundingRectangle)
        self.grid_hwnd = grid.CurrentNativeWindowHandle or None
        arr = grid.FindAllBuildCache(_U.TreeScope_Descendants, _TRUE, _cache)
        rows, sel = {}, set()
        for i in range(arr.Length):
            e = arr.GetElement(i)
            name = e.CachedName or ""
            m = CELL_RE.match(name)
            if m:
                col, n = m.group(1), int(m.group(2))
                r = rows.setdefault(n, {"_n": n})
                r[col] = _value(e)
                if col == "醫令名稱":
                    r["_rect"] = _rect(e.CachedBoundingRectangle)
                    r["_offscreen"] = bool(_cached(e, P_OFF, False))
                continue
            m = ROW_RE.match(name)
            if m and int(_cached(e, P_LSTATE, 0) or 0) & STATE_SELECTED:
                sel.add(int(m.group(1)))
        out = []
        for n, r in sorted(rows.items()):
            if clean(r.get("照會單號")):
                r["_sel"] = n in sel
                out.append(r)
        missing = [c for c in NEED_COLS if any(c not in r for r in out)]
        return out, missing

    @uia_retry
    def read_filters(self):
        """清單上方的篩選：{room: 檢查室, status: 狀態}。狀態空白 = 全部。"""
        form = self._root().FindFirst(_U.TreeScope_Descendants, _by_aid(LIST_FORM))
        if not form:
            raise NotReady("找不到「登錄作業清單」，請把 RIS 切到那個畫面。")
        out = {}
        for aid, key in (("UNITPERFORMCODE", "room"), ("EXAM_STATUS", "status")):
            e = form.FindFirst(_U.TreeScope_Descendants, _by_aid(aid))
            try:
                out[key] = norm(e.GetCurrentPropertyValue(P_VAL)) if e else ""
            except Exception:
                out[key] = ""
        return out

    @uia_retry
    def _filter_controls(self):
        """清單上方：檢查室、狀態下拉選單的 hwnd，和「查詢」按鈕（rect, hwnd）。"""
        form = self._root().FindFirst(_U.TreeScope_Descendants, _by_aid(LIST_FORM))
        if not form:
            raise NotReady("找不到「登錄作業清單」，請把 RIS 切到那個畫面。")
        out = {}
        for aid in ("UNITPERFORMCODE", "EXAM_STATUS", "B_QUERY"):
            e = form.FindFirst(_U.TreeScope_Descendants, _by_aid(aid))
            if not e:
                raise StopRun(f"清單上找不到「{aid}」，RIS 畫面可能改版了。")
            out[aid] = (_rect(e.CurrentBoundingRectangle), e.CurrentNativeWindowHandle)
        return out

    @uia_retry
    def read_room_grid(self):
        """清單右邊的檢查室表格（拖曳放下的地方）。
        回傳 {"hwnd", "rect", "rows": [{code, name, rect, off, sel}]}。拖曳經過時，
        游標下面那一間會被選起來（2026-11-26 錄影）。"""
        form = self._root().FindFirst(_U.TreeScope_Descendants, _by_aid(LIST_FORM))
        if not form:
            raise NotReady("找不到「登錄作業清單」，請把 RIS 切到那個畫面。")
        g = form.FindFirst(_U.TreeScope_Descendants, _by_aid(ROOM_GRID))
        if not g:
            raise StopRun("清單右邊找不到檢查室表格，RIS 畫面可能改版了。")
        arr = g.FindAllBuildCache(_U.TreeScope_Descendants, _TRUE, _cache)
        rows, sel = {}, set()
        for i in range(arr.Length):
            e = arr.GetElement(i)
            name = e.CachedName or ""
            m = CELL_RE.match(name)
            if m:
                col, n = m.group(1), int(m.group(2))
                r = rows.setdefault(n, {})
                if col == "代碼":
                    r["code"] = clean(_value(e))
                elif col == "檢查室名稱":
                    r["name"] = norm(_value(e))
                    r["rect"] = _rect(e.CachedBoundingRectangle)
                    r["off"] = bool(_cached(e, P_OFF, False))
                continue
            m = ROW_RE.match(name)
            if m and int(_cached(e, P_LSTATE, 0) or 0) & STATE_SELECTED:
                sel.add(int(m.group(1)))
        out = [dict(r, sel=n in sel) for n, r in sorted(rows.items()) if r.get("name")]
        return {"hwnd": g.CurrentNativeWindowHandle or None,
                "rect": _rect(g.CurrentBoundingRectangle), "rows": out}

    def list_rooms(self):
        """清單「檢查室」下拉選單的選項（不含空白）和目前選的。"""
        ctl = self._filter_controls()
        cb = ComboBoxWrapper(ctl["UNITPERFORMCODE"][1])
        items = [norm(t) for t in cb.item_texts()]
        return [t for t in items if t], norm(cb.window_text())

    def query_list(self, room, status=""):
        """把清單篩選設成 檢查室＝room、狀態＝status（空白＝全部），按「查詢」取得新的清單。
        不會碰到任何病人。"""
        if not room:
            raise StopRun("沒有選檢查室。")
        ctl = self._filter_controls()
        for aid, want, label in (("UNITPERFORMCODE", room, "檢查室"), ("EXAM_STATUS", status, "狀態")):
            cb = ComboBoxWrapper(ctl[aid][1])
            items = [norm(t) for t in cb.item_texts()]
            if want not in items:
                raise StopRun(f"RIS 的「{label}」選單裡沒有「{want or '空白'}」。")
            if norm(cb.window_text()) != want:
                cb.select(items.index(want))
        rect, hwnd = ctl["B_QUERY"]
        x, y = center(rect)
        self.click(x, y, hwnd, "還沒按查詢。")
        time.sleep(QUERY_WAIT)
        got = self.read_filters()
        if got["room"] != room or got["status"] != status:
            raise StopRun(f"清單篩選沒有設成功（現在是檢查室「{got['room']}」、"
                          f"狀態「{got['status'] or '空白'}」）。")

    @uia_retry
    def read_form(self):
        """讀「多筆登錄作業」。沒開著回 None。

        回傳 {patno, room, status, tool_hwnd, tools: {按鈕: (rect, 能不能按)}, exams}
        每筆檢查：acc 照會單號、order 醫令、doctor 報告醫師、combine 報告合併單號、
                  image 影像儲存單號
        """
        form = self._root().FindFirst(_U.TreeScope_Descendants, _by_aid(MULTI_FORM))
        if not form:
            return None
        arr = form.FindAllBuildCache(_U.TreeScope_Descendants, _TRUE, _cache)
        f = {"patno": "", "room": "", "status": "", "tool_hwnd": None, "tools": {}, "exams": []}
        section = None
        fields = {"ORDERDESC": "order", "REPORTDR_VS": "doctor",
                  "COMBINE_SEQNO": "combine", "IMAGEACCESSNO": "image"}
        dr_pending = None   # 報告醫師下拉選單：後面跟著它的「開啟」箭頭按鈕
        for i in range(arr.Length):
            e = arr.GetElement(i)
            aid = e.CachedAutomationId or ""
            if aid == "ctrlReg":            # 每一筆檢查一個 ctrlReg
                f["exams"].append({"acc": "", "order": "", "doctor": "",
                                   "combine": "", "image": "", "dr": None})
                section = "exam"
                dr_pending = None
                continue
            if (dr_pending and e.CachedControlType == CT_BUTTON
                    and norm(e.CachedName) == "開啟"):
                r = _rect(e.CachedBoundingRectangle)
                cr = dr_pending["rect"]
                if cr[0] <= r[0] and r[2] <= cr[2] + 2 and cr[1] - 2 <= r[1] and r[3] <= cr[3] + 2:
                    dr_pending["arrow"] = r
                dr_pending = None
                continue
            if aid == "ToolStrip1":
                f["tool_hwnd"] = _cached(e, P_HWND) or None
                section = "tool"
                continue
            if aid == "Splitter1":
                section = None
                continue
            if section == "exam":
                ex = f["exams"][-1]
                if aid == "ACCESSNO":       # 這一個是標籤，內容就在名稱裡
                    ex["acc"] = clean(e.CachedName)
                elif aid in fields:
                    ex[fields[aid]] = clean(_value(e))
                    if aid == "REPORTDR_VS":
                        ex["dr"] = dr_pending = {"rect": _rect(e.CachedBoundingRectangle),
                                                 "hwnd": _cached(e, P_HWND) or None,
                                                 "arrow": None}
            elif section == "tool":
                ct = e.CachedControlType
                if ct == CT_BUTTON:
                    f["tools"][e.CachedName or ""] = (_rect(e.CachedBoundingRectangle),
                                                     bool(_cached(e, P_EN, True)))
                elif ct == CT_TEXT:
                    f["status"] = norm(e.CachedName)
            else:                           # 上方病人資料
                if aid == "PATNO":
                    f["patno"] = clean(_value(e))
                elif aid == "UNITPERFORMNAME":
                    f["room"] = clean(_value(e))
        return f

    @uia_retry
    def read_tabs(self):
        """上方的分頁：[{title, exit: (rect, hwnd)}]。"""
        panel = self._root().FindFirst(_U.TreeScope_Descendants, _by_aid("TITLE_PANEL"))
        if not panel:
            return []
        arr = panel.FindAllBuildCache(_U.TreeScope_Descendants, _TRUE, _cache)
        tabs, cur = [], None
        for i in range(arr.Length):
            e = arr.GetElement(i)
            aid = e.CachedAutomationId or ""
            if aid.endswith("_Title"):
                cur = {"title": "", "exit": None}
                tabs.append(cur)
            elif cur is not None and aid == "B_ACTIVE":
                cur["title"] = norm(e.CachedName)
            elif cur is not None and aid == "B_EXIT":
                cur["exit"] = (_rect(e.CachedBoundingRectangle), _cached(e, P_HWND) or None)
        return tabs

    # ---------- 跳出視窗 ----------
    def popups(self):
        out = []
        for h in top_windows():
            if h == self.main or not _u32.IsWindowVisible(h) or pid_of(h) != self.pid:
                continue
            cls = class_of(h)
            if re.search(r"tooltips|IME|Shadow|ComboLBox|DropDown", cls, re.I):
                continue
            l, t, r, b = rect_of_hwnd(h)
            if r - l < 20 or b - t < 20:
                continue
            texts, buttons = [], []
            for c in child_windows(h):
                if not _u32.IsWindowVisible(c):
                    continue
                s = text_of(c)
                if not s.strip():
                    continue
                if "button" in class_of(c).lower():
                    buttons.append((c, s))
                else:
                    texts.append(s)
            if not texts and not buttons and self._menu_items(h):
                continue                # 右鍵選單，不是跳出視窗
            out.append(Popup(h, text_of(h), cls, texts, buttons))
        return out

    # ---------- 右鍵選單（自動模式） ----------
    @staticmethod
    def _menu_items(h):
        """h 是右鍵選單的話回傳 {選項名稱: (rect, 能不能按)}，不是回 {}。"""
        try:
            arr = _iuia.iuia.ElementFromHandle(h).FindAllBuildCache(_U.TreeScope_Descendants, _TRUE, _cache)
        except Exception:
            return {}
        items = {}
        for i in range(arr.Length):
            e = arr.GetElement(i)
            if _cached(e, P_CT) == CT_MENUITEM:
                items[norm(e.CachedName)] = (_rect(e.CachedBoundingRectangle), bool(_cached(e, P_EN, True)))
        return items

    def context_menu(self):
        """RIS 目前開著的右鍵選單：(hwnd, {選項: (rect, 能按)})；沒有回 None。"""
        for h in top_windows():
            if h == self.main or not _u32.IsWindowVisible(h) or pid_of(h) != self.pid:
                continue
            if child_windows(h):
                continue
            items = self._menu_items(h)
            if items:
                return h, items
        return None

    def open_menu(self, x, y, target, what):
        """在 (x, y) 按右鍵（檢查跟 click 一樣），等右鍵選單出來。"""
        self._guard(x, y, target, what, strict=True)
        mouse.right_click(coords=(x, y))
        self.our_tick = last_input_tick()
        m = self.wait_until(self.context_menu, 3)
        if not m:
            raise StopRun(f"按了右鍵，選單沒有出來。{what}")
        return m

    def close_menu(self):
        """右鍵選單開著就按 Esc 關掉（什麼都不選）。"""
        if self.context_menu() and self.ris_in_front():
            keyboard.send_keys("{ESC}")
            self.our_tick = last_input_tick()
            self.wait_until(lambda: self.context_menu() is None, 2)

    def click_menu(self, name, what):
        """按右鍵選單的 name。按之前確認：選單開著、有這個選項、能按、
        那個位置上真的是這個選項（旁邊就是「檢查完成」「回復成已報到」，絕對不能按錯）。"""
        m = self.context_menu()
        if not m:
            raise StopRun(f"右鍵選單不見了。{what}")
        h, items = m
        problem = None
        if name not in items:
            problem = f"右鍵選單裡沒有「{name}」"
        elif not items[name][1]:
            problem = f"右鍵選單的「{name}」是灰的"
        else:
            x, y = center(items[name][0])
            try:
                at = norm(IUIA().iuia.ElementFromPoint(_U.tagPOINT(x, y)).CurrentName)
            except Exception:
                at = ""
            if at != name:
                problem = f"右鍵選單上要按的位置不是「{name}」（是「{at}」）"
            elif _u32.WindowFromPoint(wintypes.POINT(x, y)) != h:
                problem = "右鍵選單被別的東西擋住了"
            elif self.popups():
                problem = f"RIS 跳出了視窗「{self.popups()[0].title}」"
        if problem:
            self.close_menu()
            raise StopRun(f"{problem}，程式沒有按。{what}")
        mouse.click(coords=(x, y))
        self.our_tick = last_input_tick()

    def _press(self, p, pattern):
        """按跳出視窗的按鈕（pattern=None：唯一那個按鈕）。按不掉就停。"""
        btn = p.buttons[0] if pattern is None else next(
            (b for b in p.buttons if re.match(pattern, b[1])), None)
        if btn is None:
            raise StopRun(f"跳出視窗「{p.title}」找不到要按的按鈕（{pattern}）："
                          f"{oneline(p.body)}。程式停下來沒有按。")
        _u32.PostMessageW(btn[0], BM_CLICK, 0, 0)
        if self._gone(p.h, 2.0):
            return btn[1]
        # 視窗不在作用中時 BM_CLICK 可能沒效，改用滑鼠點（先確認那個位置真的是這個按鈕）
        x, y = center(rect_of_hwnd(btn[0]))
        hit = _u32.WindowFromPoint(wintypes.POINT(x, y))
        if hit == btn[0]:
            mouse.click(coords=(x, y))
            self.our_tick = last_input_tick()
            if self._gone(p.h, 2.0):
                return btn[1]
        raise StopRun(f"跳出視窗「{p.title}」的「{btn[1]}」按不掉，程式停下來了。")

    @staticmethod
    def _gone(h, timeout):
        end = time.time() + timeout
        while time.time() < end:
            if not _u32.IsWindow(h) or not _u32.IsWindowVisible(h):
                return True
            time.sleep(0.1)
        return False

    def handle_popups(self, expect=(), close_image=True):
        """處理 RIS 跳出來的視窗。expect：這個時候預期會出現的問句。

        認得的：
          「是否將同報告醫師，報告合併？」 只在 expect 有 merge_question 時按「是」
          「影像資訊」                     記下內容；close_image 時才按「離開」
        沒看過的：只有一個按鈕 → 按掉（使用者說跳出視窗都可以擋掉）
                  不只一個按鈕（是/否、確定/取消）→ 停，不猜
        """
        for p in self.popups():
            body = p.body
            if MULTI_Q_RE.search(body):
                if "multi_question" not in expect:
                    raise StopRun(f"在不預期的時候跳出「{oneline(body)}」，程式停下來沒有按。")
                b = self._press(p, r"^是")
                self.log(f"  RIS 問「{oneline(body)}」→ 按「{b}」")
            elif MERGE_Q_RE.search(body):
                if "merge_question" not in expect:
                    raise StopRun(f"在不預期的時候跳出「{oneline(body)}」，程式停下來沒有按。")
                # 報告醫師全部一樣才按「是」；有不一樣按「否」，之後改成一樣再按報告合併
                yes = bool(self.merge_decider())
                b = self._press(p, r"^是" if yes else r"^否")
                self.merge_answered = yes
                self.log(f"  RIS 問「{oneline(body)}」→ 按「{b}」"
                         + ("" if yes else "（報告醫師不一樣，改成一樣後再按報告合併）"))
            elif SAVE_Q_RE.search(body) or DIFF_REPORT_Q_RE.search(body) or NEW_REPORT_Q_RE.search(body):
                # 只有程式自己按「報告合併」的時候才回答
                if "report_merge" not in expect:
                    raise StopRun(f"在不預期的時候跳出「{oneline(body)}」，程式停下來沒有按。")
                b = self._press(p, r"^是")
                self.log(f"  RIS 問「{oneline(body)}」→ 按「{b}」")
            elif IMAGE_Q_RE.search(body):
                # 只有程式自己按「影像合併」的時候才回答
                if "image_merge" not in expect:
                    raise StopRun(f"在不預期的時候跳出「{oneline(body)}」，程式停下來沒有按。")
                b = self._press(p, r"^是")
                self.log(f"  RIS 問「{oneline(body)}」→ 按「{b}」")
            elif NO_IMAGE_RE.search(body) and len(p.buttons) == 1:
                b = self._press(p, None)
                self.no_image = True
                self.log(f"  RIS 說「{oneline(body)}」→ 按「{b}」")
            elif p.title == IMAGE_INFO_TITLE:
                if self._image_seen is None:
                    self._image_seen = time.time()
                    self.image_info = body
                    self.log(f"  影像資訊：{oneline(body)}")
                if close_image:
                    self._press(p, r"^離開")
                    self._image_seen = None
            elif ERROR_RE.search(body) or ERROR_RE.search(p.title or ""):
                # 使用者：不是「是／否」的視窗都按掉。錯誤訊息也按掉，但按掉之後停下來，
                # 因為剛才那個動作可能沒存成功，要人看一下這位
                if len(p.buttons) == 1:
                    b = self._press(p, None)
                    self.log(f"  RIS 跳出錯誤訊息「{p.title}」：{oneline(body, 300)}（按「{b}」）")
                    raise StopRun(f"RIS 跳出錯誤訊息：{oneline(body, 400)}\n"
                                  "已經按掉了。剛才那個動作可能沒有存成功，程式停下來，請你看一下這位。")
                raise StopRun(f"RIS 跳出錯誤訊息「{p.title}」：{oneline(body, 400)}\n"
                              "程式停下來，沒有按，請你看一下。")
            elif len(p.buttons) == 1:
                b = self._press(p, None)
                self.log(f"  擋掉跳出視窗「{p.title}」：{oneline(body)}（按「{b}」）")
            else:
                names = "、".join(t for _h, t in p.buttons) or "（找不到按鈕）"
                raise StopRun(f"RIS 跳出沒看過的視窗「{p.title}」：{oneline(body)}\n"
                              f"按鈕：{names}\n"
                              "程式停下來，沒有按任何按鈕，請你自己處理。")

    def image_popup_age(self):
        return 0 if self._image_seen is None else time.time() - self._image_seen

    # ---------- 等待 ----------
    def wait_until(self, pred, timeout, expect=(), close_image=True):
        """一直處理跳出視窗、一直問 pred()，成立就回傳它的值；逾時回 None。"""
        end = time.time() + timeout
        while True:
            self.handle_popups(expect, close_image)
            try:
                v = pred()
            except StopRun:
                raise
            except Exception:
                v = None            # RIS 正在重畫時偶爾讀失敗，下一輪再讀
            if v:
                return v
            if time.time() >= end:
                return None
            time.sleep(POLL)

    # ---------- 點擊 ----------
    def click_problem(self, x, y, target, strict=False):
        """要點 (x, y)，那個位置必須是 target 這個視窗（或它裡面的東西）。沒問題回 None。
        strict：一定要正好是 target 本身（清單表格用：捲軸是表格裡的子視窗，不能點到）。"""
        if not self.main or not _u32.IsWindow(self.main):
            return "找不到 RIS 主視窗（可能被關掉了）"
        pops = self.popups()
        if pops:
            return f"RIS 跳出了視窗「{pops[0].title}」"
        if not _u32.IsWindowEnabled(self.main):
            return "RIS 被對話框鎖住了"
        if not self.ris_in_front():
            fg = _u32.GetForegroundWindow()
            if fg and pid_of(fg) == os.getpid():
                return "OWN"        # 小工具自己在前面（例如剛按了按鈕），叫 RIS 回前面再試
            return f"目前在最前面的不是 RIS，而是 {window_desc(fg)}"
        if not target or not _u32.IsWindow(target):
            return "要點的東西不見了（畫面換了）"
        hit = _u32.WindowFromPoint(wintypes.POINT(x, y))
        ok = hit == target or (not strict and _u32.IsChild(target, hit))
        if not hit or not ok:
            return f"要點的位置不是預期的地方（現在是 {window_desc(hit)}）"
        return None

    def _guard(self, x, y, target, what, expect=(), strict=False):
        """動滑鼠之前的檢查；不行就 StopRun（什麼都還沒做）。"""
        self.handle_popups(expect)
        end = time.time() + GUARD_WAIT
        retried = False
        while True:
            why = self.click_problem(x, y, target, strict)
            if why == "OWN" and not retried:
                self.activate()
                retried = True
                continue
            if not why:
                return
            if time.time() >= end:
                raise StopRun(f"{'小工具自己的視窗在最前面' if why == 'OWN' else why}。{what}")
            time.sleep(0.1)

    def click(self, x, y, target, what, double=False, expect=(), strict=False, ctrl=False):
        """點 RIS 畫面上的 (x, y)。點之前確認 RIS 在最前面、沒跳出視窗、
        而且那個位置真的是 target。ctrl：按著 Ctrl 點（清單多選用）。"""
        self._guard(x, y, target, what, expect, strict)
        if ctrl:
            keyboard.send_keys("{VK_CONTROL down}")
            try:
                mouse.click(coords=(x, y))
            finally:
                keyboard.send_keys("{VK_CONTROL up}")
        elif double:
            mouse.double_click(coords=(x, y))
        else:
            mouse.click(coords=(x, y))
        self.our_tick = last_input_tick()

    def wait_human_idle(self, what, need=3, timeout=60):
        """拖曳這種會被人的滑鼠打斷的動作之前：等人放開滑鼠幾秒。"""
        end = time.time() + timeout
        while self.human_idle_sec() < need:
            if time.time() >= end:
                raise StopRun(f"一直有人在用鍵盤滑鼠。{what}")
            time.sleep(0.5)

    def drag(self, x1, y1, src, x2, ys, dst, aim, final, what, steps=20):
        """從 (x1, y1) 按住左鍵拖到目標。跟人的做法一樣：按下就開始拖，不用按 Ctrl。

        RIS 反白（= 要放下）的那一列跟游標的位置有偏差（2026-11-26 實測：反白的大約是
        游標往上 13～18 像素那一列），所以在目標那一列裡依序試 ys 這幾個高度，
        每到一個就問 aim()：回 None = RIS 反白的就是目標。對準之後再 final() 做最後核對，
        都過了才放開。任何一步不對就拖回起點放開（清單本身不接受放下，等於取消）。
        回傳 None（放下了）或取消的原因。"""
        self._guard(x1, y1, src, what, strict=True)
        for y in ys:
            why = self.click_problem(x2, y, dst, strict=True)
            if why:
                raise StopRun(f"要放下的位置不對：{why}。{what}")
        self.wait_human_idle(what)

        def glide(ax, ay, bx, by, n=steps):
            for i in range(1, n + 1):
                mouse.move(coords=(ax + (bx - ax) * i // n, ay + (by - ay) * i // n))
                self.our_tick = last_input_tick()
                time.sleep(0.03)

        mouse.press(coords=(x1, y1))
        self.our_tick = last_input_tick()
        released = False
        cy = ys[0]
        try:
            time.sleep(0.3)
            glide(x1, y1, x2, cy)
            why = "對不準"
            for y in ys:
                if y != cy:
                    glide(x2, cy, x2, y, n=2)
                    cy = y
                time.sleep(0.35)
                why = aim()
                if why is None:
                    break
            if why is None:
                why = final()       # 左邊選取最後才核對，核對完馬上放開（清單自動更新的空檔最短）
            if why is None:
                pt = wintypes.POINT()
                _u32.GetCursorPos(ctypes.byref(pt))
                if (pt.x, pt.y) != (x2, cy):
                    why = "游標被移動了（有人在動滑鼠）"
                elif _u32.WindowFromPoint(wintypes.POINT(x2, cy)) != dst:
                    why = "游標下面不是檢查室表格"
            if why:
                glide(x2, cy, x1, y1)
                time.sleep(0.3)
                mouse.release(coords=(x1, y1))
            else:
                mouse.release(coords=(x2, cy))
            released = True
            self.our_tick = last_input_tick()
            return why
        finally:
            if not released:        # 出錯：先回到起點（清單，不接受放下）再放開
                try:
                    mouse.move(coords=(x1, y1))
                    time.sleep(0.4)
                    mouse.release(coords=(x1, y1))
                    self.our_tick = last_input_tick()
                except Exception:
                    pass

    # ---------- 下拉選單 ----------
    def dropdown_list(self):
        """RIS 目前打開的下拉選單清單（ComboLBox）；沒有回 None。"""
        for h in top_windows():
            if (pid_of(h) == self.pid and _u32.IsWindowVisible(h)
                    and class_of(h).lower() == "combolbox"):
                return h
        return None

    def dropdown_point(self, lb, idx, text):
        """把第 idx 項捲進畫面，回傳它中間的螢幕座標。
        點之前用 UIA 確認那個位置的項目名稱真的是 text，不對就停。"""
        n = _u32.SendMessageW(lb, LB_GETCOUNT, 0, 0)
        if not 0 <= idx < n:
            raise StopRun(f"下拉選單只有 {n} 項，找不到第 {idx + 1} 項。")
        h = max(1, _u32.SendMessageW(lb, LB_GETITEMHEIGHT, 0, 0))
        rc = wintypes.RECT()
        _u32.GetClientRect(lb, ctypes.byref(rc))
        shown = max(1, (rc.bottom - rc.top) // h)
        top = _u32.SendMessageW(lb, LB_GETTOPINDEX, 0, 0)
        if not top <= idx < top + shown:
            _u32.SendMessageW(lb, LB_SETTOPINDEX, max(0, idx - shown // 2), 0)
            time.sleep(0.3)
            top = _u32.SendMessageW(lb, LB_GETTOPINDEX, 0, 0)
        if not top <= idx < top + shown:
            raise StopRun("下拉選單捲不到要選的那一項。")
        pt = wintypes.POINT((rc.right - rc.left) // 2, (idx - top) * h + h // 2)
        _u32.ClientToScreen(lb, ctypes.byref(pt))
        try:
            el = IUIA().iuia.ElementFromPoint(_U.tagPOINT(pt.x, pt.y))
            name = norm(el.CurrentName)
        except Exception:
            name = ""
        if clean(name) != clean(text):
            raise StopRun("下拉選單上要點的位置不是要選的那一項（位置算錯了），程式停下來沒有點。")
        return pt.x, pt.y

    def value_at(self, x, y):
        """畫面上 (x, y) 那一格的值（用 UIA 問那個位置是什麼）。"""
        try:
            el = IUIA().iuia.ElementFromPoint(_U.tagPOINT(x, y))
            for p in (P_VAL, P_LVAL):
                v = el.GetCurrentPropertyValue(p)
                if v not in (None, ""):
                    return norm(v)
        except Exception:
            pass
        return ""

    def close_dropdown(self, combo):
        """收起下拉選單，不改選擇。"""
        _u32.SendMessageW(combo, CB_SHOWDROPDOWN, 0, 0)

    def wheel(self, x, y, target, notches, what):
        """在 (x, y) 轉滑鼠滾輪（正數往上、負數往下），檢查跟 click 一樣。"""
        self._guard(x, y, target, what, strict=True)
        mouse.scroll(coords=(x, y), wheel_dist=notches)
        self.our_tick = last_input_tick()


# ----------------------------------------------------------------------
# 純邏輯（不碰 RIS，方便單獨測試）
# ----------------------------------------------------------------------
DONE_STATUSES = {"已檢查完成", "報告完成", "已通知", "不檢查"}


def active_rows(rows):
    """還沒做完的檢查（狀態不是已檢查完成／報告完成／已通知／不檢查）。"""
    return [r for r in rows if clean(r.get("狀態")) not in DONE_STATUSES]


def order_key(order):
    return norm(order).casefold()


def parse_no_merge(lines):
    """清單檔的每一行 → 規則 {exact: set, patterns: [regex], keep: set}。"""
    rules = {"exact": set(), "patterns": [], "keep": set()}
    for ln in lines:
        ln = ln.strip()
        if not ln or ln.startswith("#"):
            continue
        keep = ln.startswith("!")
        k = order_key(ln[1:] if keep else ln)
        if not k:
            continue
        if keep:
            rules["keep"].add(k)
        elif "*" in k:
            rules["patterns"].append(re.compile("^" + ".*".join(re.escape(p) for p in k.split("*")) + "$"))
        else:
            rules["exact"].add(k)
    return rules


def is_no_merge(order, rules):
    k = order_key(order)
    if k in rules["keep"]:
        return False
    return k in rules["exact"] or any(p.match(k) for p in rules["patterns"])


def rule_count(rules):
    return len(rules["exact"]) + len(rules["patterns"])


_no_merge_cache = {"mtime": None, "rules": None}


def no_merge_orders():
    """不用合併的醫令規則。每次都看檔案有沒有改過，改了就重讀。"""
    try:
        if not os.path.exists(NO_MERGE_FILE):
            with open(NO_MERGE_FILE, "w", encoding="utf-8") as f:
                f.write(NO_MERGE_HEADER + "\n".join(NO_MERGE_DEFAULT) + "\n")
        mt = os.path.getmtime(NO_MERGE_FILE)
        if mt != _no_merge_cache["mtime"]:
            with open(NO_MERGE_FILE, encoding="utf-8-sig") as f:
                _no_merge_cache["rules"] = parse_no_merge(f)
            _no_merge_cache["mtime"] = mt
        return _no_merge_cache["rules"]
    except OSError:
        return parse_no_merge(NO_MERGE_DEFAULT)


def group_patients(rows, no_merge=None):
    """把清單分成一位一位病人，判斷要做什麼。

    kind：
      start   全部 CT、至少 2 筆、全部已報到 → 開始檢查＋報告合併＋試影像合併
      merge   全部 CT、至少 2 筆、全部檢查中、報告或影像還沒合併 → 再試合併
      waiting 全部 CT、檢查中、都合併好了 → 等技術員按檢查完成，不用做
      nomerge CT 裡有「不用合併的醫令」（ct_merge_no_merge.txt）→ 整位不碰
      manual  有非 CT、或狀態混在一起 → 不碰，留給人工
      single  只有 1 筆 CT（或沒有 CT）→ 不用合併
    """
    if no_merge is None:
        no_merge = no_merge_orders()
    by = collections.OrderedDict()
    for r in rows:
        pn = clean(r.get("病歷號"))
        if pn:
            by.setdefault(pn, []).append(r)
    out = []
    for pn, rs in by.items():
        # 已經做完的檢查（例如早上照過的胸部 X 光）不會被開始檢查影響，不算進來；
        # 多筆登錄作業打開後還會核對裡面只有這幾筆 CT
        rs = active_rows(rs) or rs
        ct = [r for r in rs if is_ct(r.get("醫令名稱"))]
        non_ct = [r for r in rs if not is_ct(r.get("醫令名稱"))]
        statuses = {clean(r.get("狀態")) for r in rs}
        p = {
            "patno": pn,
            "name": clean(rs[0].get("姓名")),
            "accs": [clean(r.get("照會單號")) for r in ct],
            "orders": [clean(r.get("醫令名稱")) for r in ct],
            "doctors": sorted({clean(r.get("報告醫師")) for r in ct} - {""}),
            "status": "、".join(sorted(statuses)),
            "combined": len(ct) > 1 and same(r.get("合併序號") for r in ct),
            "imaged": len(ct) > 1 and same(r.get("影像單號") for r in ct),
        }
        p["key"] = tuple(sorted(p["accs"]))
        skip = [o for o in p["orders"] if is_no_merge(o, no_merge)]
        if len(ct) < 2:
            p["kind"], p["why"] = "single", ("只有 1 筆 CT，不用合併" if ct else "沒有 CT")
        elif skip:
            p["kind"] = "nomerge"
            p["why"] = f"有不用合併的醫令（{skip[0][:40]}），不碰"
        elif non_ct:
            p["kind"] = "manual"
            p["why"] = f"還有 {len(non_ct)} 筆非 CT（{clean(non_ct[0].get('醫令名稱'))[:20]}…），不碰，留給人工"
        elif statuses == {READY}:
            p["kind"], p["why"] = "start", "開始檢查＋合併"
        elif statuses == {EXAMINING}:
            if p["combined"] and p["imaged"]:
                p["kind"], p["why"] = "waiting", "都合併好了，等技術員按檢查完成"
            else:
                p["kind"] = "merge"
                p["why"] = "等影像，合併影像" if p["combined"] else "報告還沒合併"
        elif len(statuses) == 1:
            p["kind"], p["why"] = "other", f"狀態是「{p['status']}」，不用做"
        else:
            p["kind"], p["why"] = "manual", f"狀態混在一起（{p['status']}），不碰"
        out.append(p)
    return out


def pick_next(pats, memory, now, retry_sec, cooldown=REOPEN_COOLDOWN):
    """這一輪要處理哪一位。先做 start（新報到的），再做 merge（等影像的）。

    memory：{key: {"last": 上次處理的時間, "manual": True/False}}
    等影像的好幾位都到了重試時間：先試等最久的那位（輪流，不會一直試同一位）。
    """
    for kind, wait in (("start", cooldown), ("merge", retry_sec)):
        ready = []
        for p in pats:
            if p["kind"] != kind:
                continue
            m = memory.get(p["key"], {})
            if m.get("manual"):
                continue
            last = m.get("last", -1e9)
            if now - last < max(wait, cooldown):
                continue
            ready.append((last, len(ready), p))
        if ready:
            return min(ready)[2]
    return None


def identity_problems(form, patno, accs):
    """打開的多筆登錄作業，跟清單上這位病人對得上嗎？有問題回傳說明。"""
    probs = []
    if clean(form.get("patno")) != patno:
        probs.append("病歷號跟清單上選的那位不一樣")
    got = [e["acc"] for e in form["exams"]]
    if set(got) != set(accs) or len(got) != len(accs):
        more = sorted(set(got) - set(accs))
        less = sorted(set(accs) - set(got))
        s = "照會單號跟清單對不上"
        if more:
            s += f"（多了 {'、'.join(more)}）"
        if less:
            s += f"（少了 {'、'.join(less)}）"
        probs.append(s)
    non_ct = [e for e in form["exams"] if not is_ct(e["order"])]
    if non_ct:
        probs.append(f"裡面有非 CT 的檢查（{non_ct[0]['order'][:20]}…）")
    return probs


def report_problems(form):
    """報告醫師、報告合併單號是不是都一樣。"""
    probs = []
    docs = [e["doctor"] for e in form["exams"]]
    if not same(docs):
        cnt = collections.Counter(d or "（空白）" for d in docs)
        probs.append("報告醫師不一樣：" + "、".join(f"{d}×{n}" for d, n in cnt.items()))
    if not same(e["combine"] for e in form["exams"]):
        probs.append("報告合併單號不一樣（報告還沒合併）")
    return probs


def image_problems(form):
    if not same(e["image"] for e in form["exams"]):
        return ["影像儲存單號不一樣（影像還沒合併）"]
    return []


def form_summary(form):
    ex = form["exams"]
    return (f"{len(ex)} 筆  報告醫師 {'、'.join(sorted({e['doctor'] or '（空白）' for e in ex}))}"
            f"  報告合併單號 {'、'.join(sorted({e['combine'] or '（空白）' for e in ex}))}"
            f"  影像儲存單號 {'、'.join(sorted({e['image'] or '（空白）' for e in ex}))}")


# ----------------------------------------------------------------------
# 一位病人的流程
# ----------------------------------------------------------------------
class Runner:
    def __init__(self, ris, log, ask, allow_room_change=False, cross_room=False, home_room=""):
        self.ris = ris
        self.log = log
        self.ask = ask      # ask(title, body) -> True（確定）/ False（取消）；會等人回答
        # 測試腳本用：RIS 登入的不是 CT 檢查室也照做，不拖回去
        self.allow_room_change = allow_room_change
        # 跨檢查室使用：在別間（例如 portable）幫忙合併。開始檢查會把檢查換到這台登入的
        # 檢查室，離開多筆登錄作業後再從清單拖回 home_room（使用者 2026-11-26 示範的做法）
        self.cross_room = cross_room
        self.home_room = home_room
        self.login_room = ""        # 開始檢查後檢查被換去的那間（= 這台 RIS 登入的檢查室）
        self.result = {}
        ris.merge_decider = self._decide_merge

    def _decide_merge(self):
        """RIS 問「是否將同報告醫師，報告合併?」：畫面上的報告醫師全部一樣才回答是。"""
        try:
            f = self.ris.read_form()
        except Exception:
            f = None
        if not f or not f["exams"]:
            self.log("  （讀不到報告醫師，當作不一樣）")
            return False
        docs = [e["doctor"] for e in f["exams"]]
        if same(docs):
            return True
        cnt = collections.Counter(d or "（空白）" for d in docs)
        self.log("  報告醫師：" + "、".join(f"{d}×{n}" for d, n in cnt.items()))
        return False

    def _form_if(self, cond):
        f = self.ris.read_form()
        return f if f and cond(f) else None

    def _settled_form(self, expect):
        """等多筆登錄作業內容穩定（連續兩次讀到一樣的檢查）。"""
        prev, end = None, time.time() + 6
        while True:
            self.ris.handle_popups(expect)
            f = self.ris.read_form()
            key = f and (f["patno"], tuple(e["acc"] for e in f["exams"]))
            if f and f["exams"] and key == prev and all(e["acc"] for e in f["exams"]):
                return f
            if time.time() >= end:
                return f
            prev = key
            time.sleep(0.4)

    def set_doctor(self, acc, doctor):
        """把照會 acc 那一筆的報告醫師改成 doctor，跟人的做法一樣：
        點開報告醫師下拉選單 → 捲到那位醫師 → 點下去。改完核對。"""
        def exam_of(f):
            if len(f["exams"]) != len({e["acc"] for e in f["exams"]}):
                return None         # RIS 正在重畫（會短暫多出一份），等它好
            return next((e for e in f["exams"] if e["acc"] == acc), None)

        form = self.ris.wait_until(lambda: self._form_if(exam_of), 5)
        if not form:
            raise StopRun(f"找不到照會 {acc} 那一筆檢查。")
        dr = exam_of(form)["dr"] or {}
        combo = dr.get("hwnd")
        if not combo:
            raise StopRun(f"找不到照會 {acc} 的報告醫師下拉選單。")
        items = [clean(t) for t in ComboBoxWrapper(combo).item_texts()]
        if doctor not in items:
            raise StopRun(f"照會 {acc} 的報告醫師下拉選單裡，找不到要改成的那位醫師。")
        idx = items.index(doctor)
        if dr.get("arrow"):
            x, y = center(dr["arrow"])
        else:                       # 找不到箭頭按鈕：點選單右邊
            l, t, r, b = dr["rect"]
            x, y = r - 8, (t + b) // 2
        self.ris.click(x, y, combo, f"還沒改照會 {acc} 的報告醫師。")
        lb = self.ris.wait_until(self.ris.dropdown_list, 3)
        if not lb:
            raise StopRun(f"點了照會 {acc} 的報告醫師下拉選單，選單沒有打開。")
        try:
            px, py = self.ris.dropdown_point(lb, idx, doctor)
            self.ris.click(px, py, lb, f"還沒改照會 {acc} 的報告醫師。", strict=True)
        except Exception:
            self.ris.close_dropdown(combo)
            raise
        ok = self.ris.wait_until(
            lambda: self._form_if(lambda f: (exam_of(f) or {}).get("doctor") == doctor), 5)
        if not ok:
            self.ris.close_dropdown(combo)
            raise StopRun(f"照會 {acc} 的報告醫師沒有變成要的醫師，程式停下來了（還沒存檔）。\n"
                          "請你看一下那一筆的報告醫師。")

    def _press_tool(self, form, name, expect=()):
        rect, enabled = form["tools"].get(name, (None, False))
        if not rect:
            raise StopRun(f"找不到「{name}」按鈕。")
        if not enabled:
            raise StopRun(f"「{name}」按鈕是灰的，按不了。")
        x, y = center(rect)
        self.ris.click(x, y, form["tool_hwnd"], f"還沒按「{name}」。", expect=expect)

    # ---------- 步驟 ----------
    def precheck(self, pat):
        """動手之前：清單上這位還在、看得到、RIS 在清單畫面。沒動到任何東西。
        回傳要點兩下的那一列。"""
        if self.ris.read_form() is not None:
            raise SkipPatient("「多筆登錄作業」開著（有人在用）")
        try:
            flt = self.ris.read_filters()
        except NotReady as e:
            raise SkipPatient(str(e))
        if not flt["room"]:
            # 檢查室空白 = 全部檢查室：別的 CT 室的病人也會被開始檢查、換到這間
            raise StopRun("清單的「檢查室」篩選是空白（全部檢查室），這樣會動到別間檢查室的病人。\n"
                          "請把檢查室設成這台負責的 CT 檢查室再開始監看。")
        for attempt in range(SCROLL_TRIES + 1):
            try:
                rows = self.ris.read_list()
            except NotReady as e:
                raise SkipPatient(str(e))
            mine = active_rows([r for r in rows if clean(r.get("病歷號")) == pat["patno"]])
            mine_ct = [r for r in mine if clean(r.get("照會單號")) in pat["accs"]]
            if (sorted(clean(r["照會單號"]) for r in mine_ct) != sorted(pat["accs"])
                    or len(mine) != len(mine_ct)):
                raise SkipPatient("清單上這位的檢查變了，下一輪重新判斷")
            gl, gt, gr, gb = self.ris.grid_rect

            def shown(r):
                rc = r.get("_rect")
                return bool(rc and not r.get("_offscreen") and rc[2] > rc[0]
                            and gl <= rc[0] and rc[2] <= gr and gt <= rc[1] and rc[3] <= gb)

            for r in mine_ct:
                if shown(r):
                    x, y = center(r["_rect"])
                    why = self.ris.click_problem(x, y, self.ris.grid_hwnd, strict=True)
                    if why == "OWN":
                        self.ris.activate()
                        why = self.ris.click_problem(x, y, self.ris.grid_hwnd, strict=True)
                    if why and attempt < SCROLL_TRIES and "預期" in why:
                        break       # 可能被捲軸擋住一半：再捲一點
                    if why:
                        raise SkipPatient(f"現在不能點清單：{why}")
                    return r
            if attempt == SCROLL_TRIES:
                break
            # 這位在畫面外：用滾輪把他捲進來（每格大約 3 列）
            vis = [r["_n"] for r in rows if shown(r)]
            want = [r["_n"] for r in mine_ct]
            if not vis:
                raise SkipPatient("清單上一列都看不到")
            if min(want) >= max(vis):
                notches = -min(10, (min(want) - max(vis) + 2) // 3 + 1)
            elif max(want) <= min(vis):
                notches = min(10, (min(vis) - max(want) + 2) // 3 + 1)
            else:
                raise SkipPatient("這位在清單中間卻點不到（可能被擋住）")
            if attempt == 0:
                self.log(f"  清單上要捲動才看得到這位，往{'下' if notches < 0 else '上'}捲")
            cx, cy = center(self.ris.grid_rect)
            self.ris.wheel(cx, cy, self.ris.grid_hwnd, notches, "還沒捲動清單。")
            time.sleep(0.6)
        raise SkipPatient(f"捲了 {SCROLL_TRIES} 次清單還是看不到這位")

    def open_patient(self, pat, target):
        """點兩下打開多筆登錄作業，確認打開的就是這位、全部是 CT。"""
        x, y = center(target["_rect"])
        self.ris.click(x, y, self.ris.grid_hwnd, "還沒點兩下，什麼都沒動。", double=True, strict=True)
        # 回答 RIS 的問句之前，先確認點到的真的是這位（清單會自動更新，列可能剛好移動）
        end = time.time() + OPEN_WAIT
        sel_end = time.time() + 3       # RIS 處理點擊需要一點時間，剛點完可能還是舊的選取
        verified = False
        while True:
            if not verified:
                try:
                    sel = [r for r in self.ris.read_list() if r["_sel"]]
                except NotReady:
                    sel = []
                if sel and all(clean(r.get("照會單號")) in pat["accs"] for r in sel):
                    verified = True
                elif time.time() >= sel_end:
                    raise StopRun("點兩下之後，清單上選到的不是這位病人（清單可能剛好更新）。"
                                  "程式沒有回答 RIS 的任何問題，請你看一下 RIS 畫面。")
            if verified:
                self.ris.handle_popups({"multi_question", "merge_question"})
            f = self.ris.read_form()
            if f and verified:
                break
            if time.time() >= end:
                raise StopRun(f"點兩下之後 {OPEN_WAIT} 秒都沒有打開「多筆登錄作業」"
                              + ("" if verified else "（也確認不了選到哪一位）") + "。")
            time.sleep(POLL)
        form = self._settled_form({"merge_question"})
        probs = identity_problems(form, pat["patno"], pat["accs"])
        if probs:
            raise StopRun("打開的「多筆登錄作業」跟清單對不上：" + "；".join(probs)
                          + "\n程式沒有按任何按鈕，請你自己確認。")
        self.log(f"  多筆登錄作業：{len(form['exams'])} 筆，都是 CT，病歷號和照會單號都對得上"
                 f"（{form['status']}）")
        return form

    def start_exam(self, form):
        st = form["status"]
        if EXAMINING in st:
            return form
        if READY not in st:
            raise StopRun(f"「多筆登錄作業」的狀態是「{st}」，不是{READY}或{EXAMINING}，"
                          "程式沒有按任何按鈕。")
        room = form["room"]
        # 跨檢查室：多筆登錄作業上方的檢查室（開始檢查前）就是之後要丟回去的那間，
        # 要跟小工具左上角選的一樣；開始檢查後它會跳成這台登入的檢查室，就去那間撈回來
        if self.cross_room and norm(room) != norm(self.home_room):
            raise StopRun(f"「多筆登錄作業」上方的檢查室是「{room or '空白'}」，"
                          f"跟小工具選的「{self.home_room}」不一樣。\n程式沒有按開始檢查，請你看一下。")
        self.ris.merge_answered = None
        self._press_tool(form, "開始檢查", expect={"merge_question"})
        form = self.ris.wait_until(lambda: self._form_if(lambda f: EXAMINING in f["status"]),
                                   STEP_WAIT, expect={"merge_question"})
        if not form:
            raise StopRun(f"按了「開始檢查」，但 {STEP_WAIT} 秒內狀態沒有變成「{EXAMINING}」。")
        self.log(f"  開始檢查 → {form['status']}")
        if room and form["room"] and form["room"] != room and self.cross_room:
            self.login_room = form["room"]
            self.log(f"  檢查室從「{room}」換到「{form['room']}」"
                     f"（跨檢查室使用：離開後拖回「{self.home_room or room}」）")
        elif room and form["room"] and form["room"] != room:
            if not self.allow_room_change:
                raise StopRun(f"按了「開始檢查」之後，檢查室從「{room}」變成「{form['room']}」。\n"
                              "RIS 登入的檢查室可能不對，程式停止監看，以免把更多病人換到別的檢查室。\n"
                              "這位病人請你自己確認。")
            self.log(f"  ※ 檢查室從「{room}」變成「{form['room']}」（測試模式，照做）")
        # 開始檢查之後 RIS 會問「是否將同報告醫師，報告合併?」（實測：狀態變了才跳出來），
        # 等它一下；報告本來就合併好的可能不會問
        self.ris.wait_until(lambda: self.ris.merge_answered is not None, 3,
                            expect={"merge_question"})
        return self.ris.read_form() or form

    def ensure_reports(self, pat, form):
        """報告醫師、報告合併單號要都一樣。

        醫師不一樣：改成 abdomen／pelvis／chest 那幾筆的醫師（決定不了就請人改），
                    再由程式按「報告合併」。
        醫師一樣、報告合併單號不一樣：程式按「報告合併」。
        """
        q = {"merge_question"}      # 問句可能晚一點才跳出來
        qm = q | {"report_merge"}   # 按報告合併時 RIS 會問「資料未存檔」「不同報告確定合併」
        form = (self.ris.wait_until(
            lambda: self._form_if(lambda f: all(e["doctor"] for e in f["exams"])), DOCTOR_WAIT,
            expect=q) or self.ris.read_form())
        merged_once = False
        changed_once = False
        asked_once = False
        while True:
            self.ris.handle_popups(q)
            form = self.ris.read_form()
            if form is None:
                raise StopRun("「多筆登錄作業」被關掉了。")
            probs = identity_problems(form, pat["patno"], pat["accs"])
            if probs:
                raise StopRun("「多筆登錄作業」的內容變了：" + "；".join(probs))
            docs_ok = same(e["doctor"] for e in form["exams"])
            comb_ok = same(e["combine"] for e in form["exams"])
            if docs_ok and comb_ok:
                self.log("  報告醫師、報告合併單號都一樣")
                self.result["report"] = "✔ 合併成功" if (merged_once or self.ris.merge_answered) else "✔ 已合併"
                # 改過醫師、但報告本來就合併好（不用按報告合併）：醫師還沒存，按「存檔」
                # （2026-11-25 測試病人實測：回已報到不會拆開報告合併，這種情況真的會發生）
                touched = changed_once or asked_once or self.ris.merge_answered is False
                if touched and form["tools"].get("存檔", (None, False))[1]:
                    self._press_tool(form, "存檔", expect=qm)
                    if not self.ris.wait_until(
                            lambda: self._form_if(lambda f: not f["tools"].get("存檔", (0, 0))[1]),
                            STEP_WAIT, expect=qm):
                        raise StopRun(f"按了「存檔」，{STEP_WAIT} 秒內「存檔」沒有變灰（可能沒存成功）。")
                    self.log("  按「存檔」→ 存好了")
                    form = self.ris.read_form() or form
                return form
            if docs_ok:
                # 醫師一樣了，報告還沒合併 → 按下方的「報告合併」
                if merged_once:
                    raise StopRun("按了「報告合併」，但報告合併單號還是不一樣。")
                self._press_tool(form, "報告合併", expect=qm)
                if not self.ris.wait_until(
                        lambda: self._form_if(lambda f: same(e["combine"] for e in f["exams"])),
                        STEP_WAIT, expect=qm):
                    raise StopRun(f"按了「報告合併」，{STEP_WAIT} 秒內報告合併單號沒有變成一樣。")
                self.log("  按「報告合併」→ 報告合併單號都一樣")
                merged_once = True
                continue
            # 醫師不一樣：改成 body（abdomen／pelvis／chest）的醫師
            target, why = target_doctor(form["exams"])
            if target and not changed_once:
                todo = [e for e in form["exams"] if clean(e["doctor"]) != target]
                self.log(f"  報告醫師不一樣：{len(todo)} 筆改成跟 abdomen／pelvis／chest 一樣")
                for e in todo:
                    self.set_doctor(e["acc"], target)
                    self.log(f"    {e['acc']}（{e['order'][:40]}）改好了")
                changed_once = True
                continue
            # 決定不了（或改了還是不一樣）：請人改
            probs = report_problems(form) + ([why] if why else [])
            if changed_once:
                probs.append("程式改過一次了，還是不一樣")
            self.log("  需要你處理：" + "；".join(probs))
            ok = self.ask("報告醫師要你處理",
                          f"{pat['name']}：\n" + "；\n".join(probs) + "\n\n"
                          "請到 RIS 的「多筆登錄作業」把報告醫師改成一樣。\n"
                          "（報告合併不用按，改好之後程式會按。）\n\n"
                          "改好了按「確定」，程式會重新核對再繼續；\n"
                          "按「取消」：這位你自己處理，程式這次不會再碰他（視窗留著給你）。")
            if not ok:
                raise ManualPatient("你選擇自己處理")
            asked_once = True
            self.ris.reset_idle()
            self.ris.activate()

    def merge_images(self, form):
        """按影像合併。成功回 True；影像還沒到（按了沒反應）回 False。"""
        if not image_problems(form):
            self.log("  影像儲存單號本來就一樣，不用合併")
            self.result["image"] = "✔ 已合併"
            return True
        self.ris.image_info = ""
        self.ris._image_seen = None
        self.ris.no_image = False
        qi = {"image_merge"}
        self._press_tool(form, "影像合併", expect=qi)
        end = time.time() + IMAGE_WAIT
        done = False
        done_at = None
        while True:
            # 「影像資訊」視窗開著的時候合併還在進行，先別關：合併好了或放了夠久才關
            self.ris.handle_popups(qi, close_image=done or
                                   self.ris.image_popup_age() > IMAGE_POPUP_HOLD)
            try:
                f = self.ris.read_form()
            except Exception:
                f = None
            if f and not image_problems(f) and not done:
                done, done_at = True, time.time()
            # 有影像時「影像資訊」是合併好之後才跳（2026-11-25 實測）：等它出來、記下來、關掉
            if done and self.ris.image_info and self.ris.image_popup_age() == 0:
                break
            if done and time.time() - done_at > IMAGE_INFO_AFTER:
                self.ris.handle_popups(qi, close_image=True)
                break
            if self.ris.no_image and not done:      # RIS 明講沒影像，不用再等
                break
            if time.time() >= end:
                self.ris.handle_popups(qi, close_image=True)
                break
            time.sleep(POLL)
        if done:
            info = self.ris.image_info
            self.log("  影像合併 → 影像儲存單號都一樣"
                     + (f"（影像資訊：{oneline(info, 160)}）" if info else ""))
            self.result["image"] = "✔ 合併成功"
        elif self.ris.no_image:
            self.log("  影像還沒到（RIS 說無影像），之後再試")
            self.result["image"] = "✗ 無影像"
        else:
            info = self.ris.image_info
            self.log("  影像合併沒有反應，影像可能還沒到"
                     + (f"（影像資訊：{oneline(info, 120)}）" if info else ""))
            self.result["image"] = "✗ 沒合併成功"
        return done

    def close_form(self):
        """點上方「多筆登錄作業」分頁的 X 離開（使用者確認過不會改到東西）。"""
        tab = next((t for t in self.ris.read_tabs() if MULTI_TAB_RE.search(t["title"])), None)
        if not tab or not tab["exit"] or not tab["exit"][1]:
            raise StopRun("找不到「多筆登錄作業」分頁的 X，沒辦法離開。請你自己關掉。")
        rect, hwnd = tab["exit"]
        # RIS 有資料沒存時，分頁的 X 會變灰（實測「公告事項」有未存資料時就是這樣）
        if not self.ris.wait_until(lambda: _u32.IsWindowEnabled(hwnd), 3):
            raise StopRun("「多筆登錄作業」分頁的 X 是灰的，RIS 不讓關（可能有資料沒存）。\n"
                          "程式停下來，請你自己看一下。")
        x, y = center(rect)
        self.ris.click(x, y, hwnd, "還沒關掉多筆登錄作業。")
        if not self.ris.wait_until(lambda: self.ris.read_form() is None, CLOSE_WAIT):
            raise StopRun("點了分頁的 X，但「多筆登錄作業」沒有關掉。")
        self.log("  點 X 離開多筆登錄作業")

    # ---------- 跨檢查室：拖回原本的檢查室 ----------
    def return_home(self, pat):
        """離開多筆登錄作業後：重新查詢原本檢查室的清單，這位的檢查不在了（被開始檢查換走）
        就去換過去的那間拖回來。"""
        home = self.home_room
        if not home:
            raise StopRun("跨檢查室使用：不知道原本是哪一間檢查室。")
        self.ris.query_list(home, "")
        here = {clean(r.get("照會單號")) for r in self.ris.read_list()}
        missing = [a for a in pat["accs"] if a not in here]
        if not missing:
            return
        away = self.login_room
        if not away or norm(away) == norm(home):
            raise StopRun(f"處理完之後，照會 {'、'.join(missing)} 不在「{home}」的清單上，"
                          "也不知道被換到哪一間。\n程式停止監看，請你確認它們在哪一間。")
        self.move_back(missing, home, away)

    def _show_rows(self, accs, room):
        """清單上（目前查詢的是 room）找出這幾筆，捲到全部都看得到、點得到。回傳 {照會: 列}。"""
        for attempt in range(SCROLL_TRIES + 1):
            rows = self.ris.read_list()
            mine = {clean(r.get("照會單號")): r for r in rows if clean(r.get("照會單號")) in accs}
            lost = [a for a in accs if a not in mine]
            if lost:
                raise StopRun(f"「{room}」的清單上找不到照會 {'、'.join(lost)}，沒辦法拖回去。\n"
                              "請你自己找一下它們在哪一間。")
            gl, gt, gr, gb = self.ris.grid_rect

            def shown(r):
                rc = r.get("_rect")
                return bool(rc and not r.get("_offscreen") and rc[2] > rc[0]
                            and gl <= rc[0] and rc[2] <= gr and gt <= rc[1] and rc[3] <= gb)

            blocked = False
            if all(shown(r) for r in mine.values()):
                for r in mine.values():
                    x, y = center(r["_rect"])
                    why = self.ris.click_problem(x, y, self.ris.grid_hwnd, strict=True)
                    if why == "OWN":
                        self.ris.activate()
                        why = self.ris.click_problem(x, y, self.ris.grid_hwnd, strict=True)
                    if why and attempt < SCROLL_TRIES and "預期" in why:
                        blocked = True      # 被捲軸擋住一半：再捲一點
                        break
                    if why:
                        raise StopRun(f"現在不能點清單：{why}。還沒拖回去。")
                if not blocked:
                    return mine
            if attempt == SCROLL_TRIES:
                break
            vis = [r["_n"] for r in rows if shown(r)]
            want = [r["_n"] for r in mine.values()]
            if not vis:
                raise StopRun("清單上一列都看不到，還沒拖回去。")
            if max(want) - min(want) >= len(vis):
                raise StopRun("要拖的幾筆在清單上隔太遠，沒辦法同時出現在畫面上。還沒拖回去。")
            if max(want) >= max(vis):
                notches = -min(10, (max(want) - max(vis) + 2) // 3 + 1)
            else:
                notches = min(10, (min(vis) - min(want) + 2) // 3 + 1)
            cx, cy = center(self.ris.grid_rect)
            self.ris.wheel(cx, cy, self.ris.grid_hwnd, notches, "還沒捲動清單。")
            time.sleep(0.6)
        raise StopRun(f"捲了 {SCROLL_TRIES} 次清單還是沒辦法讓要拖的幾筆都看得到。還沒拖回去。")

    def _selected(self):
        return {clean(r.get("照會單號")) for r in self.ris.read_list(tries=1) if r["_sel"]}

    def _select_rows(self, accs, room):
        """跟人一樣：點第一筆，其他按著 Ctrl 點。確認選到的剛好是這幾筆。"""
        gap = _u32.GetDoubleClickTime() / 1000.0 + 0.2     # 不能被當成點兩下（會打開多筆登錄作業）
        for attempt in range(2):
            rows = self._show_rows(accs, room)
            for i, a in enumerate(accs):
                x, y = center(rows[a]["_rect"])
                self.ris.click(x, y, self.ris.grid_hwnd, "還沒選取要拖的檢查。",
                               strict=True, ctrl=i > 0)
                time.sleep(gap)
            if self.ris.wait_until(lambda: self._selected() == set(accs), 3):
                return rows
            if attempt == 0:
                self.log("  清單選取跟要的不一樣（清單可能剛好更新），重選一次")
        raise StopRun(f"用 Ctrl 選不到剛好照會 {'、'.join(accs)} 這幾筆。還沒拖回去，"
                      f"它們還在「{room}」。")

    def move_back(self, accs, home, away):
        """把 accs 從 away 的清單拖到右邊檢查室表格的 home（使用者 2026-11-26 示範）：
        查詢 away → 點第一筆、Ctrl 點其他 → 按住拖到右邊的 home 放開（RIS 不會再問）
        → 查詢 home 確認回來了。放開前核對選取和目標，不對就拖回原處放開（等於沒做）。"""
        self.log(f"  跨檢查室：照會 {'、'.join(accs)} 在「{away}」，拖回「{home}」")
        want = set(accs)

        def aim():
            if self.ris.popups():
                return "RIS 跳出了視窗"
            try:
                for k in range(2):          # 連續兩次都是目標才算（RIS 反白要一點時間）
                    on = [r["name"] for r in self.ris.read_room_grid()["rows"] if r["sel"]]
                    if [norm(n) for n in on] != [norm(home)]:
                        return f"右邊反白的是「{'、'.join(on) or '沒有'}」，不是「{home}」"
                    if k == 0:
                        time.sleep(0.2)
                return None
            except Exception as e:
                return f"讀不到右邊的檢查室表格（{type(e).__name__}）"

        def final():
            if self.ris.popups():
                return "RIS 跳出了視窗"
            try:
                sel = self._selected()
            except Exception as e:
                return f"讀不到左邊清單（{type(e).__name__}）"
            if sel != want:
                return f"左邊清單選取的不是剛好這幾筆（現在選了 {len(sel)} 筆）"
            return None

        # 直接拖（使用者要快）。清單每 60 秒自動更新、選取會被重設（2026-11-26 實測），
        # 剛好碰到時放開前的核對會擋下來 → 重新查詢、重選、再拖；剛更新完，下一次不會那麼快
        for attempt in range(1, DRAG_TRIES + 1):
            self.ris.query_list(away, "")
            rows = self._select_rows(accs, away)
            before = {clean(r.get("照會單號")) for r in self.ris.read_list()}
            grid = self.ris.read_room_grid()
            tgt = next((r for r in grid["rows"] if norm(r["name"]) == norm(home)), None)
            if not tgt:
                raise StopRun(f"右邊的檢查室表格裡找不到「{home}」。還沒拖回去。")
            gl, gt, gr, gb = grid["rect"]
            rc = tgt["rect"]
            if tgt["off"] or not (gl <= rc[0] and rc[2] <= gr and gt <= rc[1] and rc[3] <= gb):
                raise StopRun(f"右邊的檢查室表格上看不到「{home}」（要捲動）。還沒拖回去。")
            sx, sy = center(rows[accs[0]]["_rect"])
            tx, ty = center(rc)
            if norm(self.ris.value_at(tx, ty)) != norm(home):
                raise StopRun(f"右邊檢查室表格上要放下的位置不是「{home}」（位置算錯了），程式沒有拖。")
            # 只在目標那一列裡面試（游標下面那一列、RIS 反白的那一列都要是目標）：先往下再往上
            top, bot = rc[1], rc[3]
            ys = list(range(ty, bot - 1, 3)) + list(range(ty - 3, top, -3))
            why = self.ris.drag(sx, sy, self.ris.grid_hwnd, tx, ys, grid["hwnd"], aim, final,
                                "還沒拖回去。")
            if not why:
                break
            if attempt < DRAG_TRIES:
                self.log(f"  拖曳沒有放下（{why}），重新查詢清單再拖一次")
        else:
            raise StopRun(f"拖了 {DRAG_TRIES} 次都沒放下，最後一次：{why}。\n"
                          f"照會 {'、'.join(accs)} 還在「{away}」，請你自己拖回「{home}」。")
        # 放下後 RIS 會自己把它們從清單拿掉（錄影：約 1 秒）
        self.ris.wait_until(
            lambda: not ({clean(r.get("照會單號")) for r in self.ris.read_list(tries=1)} & want),
            DRAG_GONE_WAIT)
        self.ris.query_list(away, "")
        after = {clean(r.get("照會單號")) for r in self.ris.read_list()}
        others = sorted((before - after) - want)
        if others:
            raise StopRun(f"放下之後，除了這位的檢查，照會 {'、'.join(others)} 也從「{away}」的清單不見了，"
                          f"可能被一起拖到「{home}」。\n程式停止監看，請你馬上確認，必要時拖回「{away}」。")
        self.ris.query_list(home, "")
        back = {clean(r.get("照會單號")): r for r in self.ris.read_list()
                if clean(r.get("照會單號")) in want}
        bad = [a for a in accs if a not in back or clean(back[a].get("檢查室")) != tgt["code"]]
        if bad:
            raise StopRun(f"拖曳之後，照會 {'、'.join(bad)} 沒有回到「{home}」的清單上。\n"
                          "程式停止監看，請你確認它們現在在哪一間。")
        sts = "、".join(sorted({clean(r.get("狀態")) for r in back.values()}))
        self.log(f"  拖回「{home}」了：{len(accs)} 筆都回到清單上（{sts}）")

    def process(self, pat):
        """處理一位。回傳 'merged'（影像也合併好了）或 'waiting'（等影像）。
        self.result：這次報告合併、影像合併的結果（給畫面中間的表格顯示）。"""
        t0 = time.time()
        self.result = {}
        self.ris.merge_answered = None      # 上一位的回答不要帶過來
        target = self.precheck(pat)
        form = self.open_patient(pat, target)
        form = self.start_exam(form)
        form = self.ensure_reports(pat, form)
        merged = self.merge_images(form)
        form = self.ris.read_form()
        if form is None:
            raise StopRun("「多筆登錄作業」被關掉了。")
        probs = identity_problems(form, pat["patno"], pat["accs"]) + report_problems(form)
        if probs:
            raise StopRun("離開前的核對沒過：" + "；".join(probs))
        self.log("  核對：" + form_summary(form))
        self.close_form()
        if self.cross_room:
            self.return_home(pat)
            self.log(f"  V {'影像也合併好了' if merged else '報告合併好了，等影像'}"
                     f"（{time.time() - t0:.1f} 秒）；檢查完成留給技術員")
            return "merged" if merged else "waiting"
        # 開始檢查之後，這位要還在清單上。不見了可能是被換到別的檢查室，
        # 也可能是清單的「狀態」篩選不顯示檢查中 —— 兩種都要停：前者會把更多病人
        # 換錯檢查室，後者程式之後看不到這位、沒辦法再合併影像。
        try:
            left = {clean(r.get("照會單號")) for r in self.ris.read_list()}
        except NotReady:
            left = None
        if left is not None and not set(pat["accs"]) & left and self.allow_room_change:
            self.log("  ※ 這位從清單上不見了（測試模式：檢查室被換掉是預期的）")
        elif left is not None and not set(pat["accs"]) & left:
            raise StopRun("處理完之後，這位從清單上不見了。可能是：\n"
                          "  1. RIS 登入的檢查室不對，開始檢查把他換到別的檢查室\n"
                          "  2. 清單的「狀態」篩選不顯示「檢查中」，程式之後看不到他\n"
                          "程式停止監看，請你確認這兩件事。")
        self.log(f"  V {'影像也合併好了' if merged else '報告合併好了，等影像'}"
                 f"（{time.time() - t0:.1f} 秒）；檢查完成留給技術員")
        return "merged" if merged else "waiting"


# ----------------------------------------------------------------------
# 介面
# ----------------------------------------------------------------------
def auto_plan(pats):
    """自動模式：同樣的病人分類（全部 CT、至少 2 筆、沒有 HMC、沒有非 CT、狀態一致），
    但不開始檢查 —— 報告沒合併就右鍵合併報告，影像沒合併就右鍵影像合併（等影像的每隔幾分鐘再試）。
    kind 換成：start = 要合併報告（馬上做）、merge = 只差影像（照重試間隔）、waiting = 都好了。"""
    out = []
    for p in pats:
        q = dict(p)
        if p["kind"] in ("start", "merge"):
            if not p["combined"]:
                q["kind"] = "start"
                q["why"] = "右鍵合併報告" + ("、影像合併" if not p["imaged"] else "")
            elif not p["imaged"]:
                q["kind"], q["why"] = "merge", "右鍵影像合併（等影像，每隔幾分鐘再試）"
            else:
                q["kind"], q["why"] = "waiting", "報告、影像都合併好了"
        out.append(q)
    return out


class AutoRunner(Runner):
    """自動模式：不點開病人、不按開始檢查（檢查室不會變）。在清單上點這位的第一筆 CT、
    Ctrl 點其他幾筆 → 右鍵 →「合併報告」／「影像合併」（使用者 2026-11-26 說的快速做法）。
    規則跟手動模式一樣：全部 CT、沒有 HMC、報告醫師要一樣；跳出視窗照同樣的規則回答。"""

    def __init__(self, ris, log, ask, home_room):
        super().__init__(ris, log, ask, home_room=home_room)
        self._docs = []
        ris.merge_decider = lambda: same(self._docs)    # 「同報告醫師報告合併?」：醫師都一樣才是

    def _rows_now(self, pat, after=False):
        """清單上這位的 CT 還是這幾筆、還沒做完。after：已經按過選單了，對不上就停。"""
        rows = self.ris.read_list()
        mine_all = active_rows([r for r in rows if clean(r.get("病歷號")) == pat["patno"]])
        mine = [r for r in mine_all if clean(r.get("照會單號")) in pat["accs"]]
        if sorted(clean(r["照會單號"]) for r in mine) != sorted(pat["accs"]) or len(mine_all) != len(mine):
            if after:
                raise StopRun("按完之後重新查詢，清單上這位的 CT 跟原本對不上（檢查室可能被換了）。\n"
                              "程式停下來，請你看一下這位。")
            raise SkipPatient("清單上這位的檢查變了，下一輪重新判斷")
        return mine

    def _menu(self, pat, item):
        """選好這位的 CT → 右鍵 → 確認選取還是這幾筆 → 按 item。"""
        accs = pat["accs"]
        for attempt in range(1, DRAG_TRIES + 1):
            rows = self._select_rows(accs, self.home_room)
            x, y = center(rows[accs[0]]["_rect"])
            self.ris.open_menu(x, y, self.ris.grid_hwnd, f"還沒按「{item}」。")
            sel = self._selected()
            if sel == set(accs):
                self.ris.click_menu(item, f"還沒按「{item}」。")
                return
            self.ris.close_menu()
            if attempt < DRAG_TRIES:
                self.log(f"  右鍵後選取跑掉了（選了 {len(sel)} 筆，清單可能剛好更新），重選一次")
        raise StopRun(f"選了 {DRAG_TRIES} 次，右鍵時選取都不是剛好這幾筆，程式沒有按「{item}」。")

    def _settle(self, expect, quiet, timeout=30):
        """按了選單之後：照規則回答 RIS 的問句，直到連續 quiet 秒都沒有視窗。
        「影像資訊」留著 IMAGE_POPUP_HOLD 秒再關（手動模式實測合併是在它開著時完成的）。"""
        end, calm = time.time() + timeout, None
        while True:
            self.ris.handle_popups(expect, close_image=self.ris.image_popup_age() >= IMAGE_POPUP_HOLD)
            if self.ris.popups() or self.ris.context_menu():
                calm = None
            else:
                calm = calm or time.time()
                if time.time() - calm >= quiet:
                    return
            if time.time() >= end:
                raise StopRun(f"按了之後 {timeout} 秒，RIS 一直有視窗沒關。程式停下來，請你看一下。")
            time.sleep(POLL)

    def process(self, pat):
        """處理一位。回傳 'merged'（報告、影像都合併好了）或 'waiting'（等影像）。"""
        t0 = time.time()
        self.result = {}
        self.ris.merge_answered = None
        self.ris.no_image = False
        self.ris.image_info = ""
        if self.ris.read_form() is not None:
            raise SkipPatient("「多筆登錄作業」開著（有人在用）")
        flt = self.ris.read_filters()
        if norm(flt["room"]) != norm(self.home_room) or flt["status"]:
            self.ris.query_list(self.home_room, "")
        rows = self._rows_now(pat)
        rooms = sorted({clean(r.get("檢查室")) for r in rows})
        self._docs = [clean(r.get("報告醫師")) for r in rows]
        if len(set(self._docs) - {""}) > 1:
            cnt = collections.Counter(d or "（空白）" for d in self._docs)
            self.log("  報告醫師：" + "、".join(f"{d}×{n}" for d, n in cnt.items()))
            raise ManualPatient("報告醫師不一樣，自動模式不改醫師（請用手動模式，或在 RIS 改成一樣）")
        # 報告
        if not same(r.get("合併序號") for r in rows):
            self._menu(pat, "合併報告")
            self._settle({"report_merge", "merge_question"}, quiet=1.5)
            self.ris.query_list(self.home_room, "")
            rows = self._rows_now(pat, after=True)
            if not same(r.get("合併序號") for r in rows):
                raise StopRun("按了「合併報告」，重新查詢後報告合併單號還是不一樣。\n"
                              "程式停下來，請你看一下這位。")
            self.log(f"  合併報告 → 報告合併單號都一樣（{clean(rows[0].get('合併序號'))}）")
            self.result["report"] = "✔ 合併成功"
        else:
            self.result["report"] = "✔ 已合併"
        # 影像
        if same(r.get("影像單號") for r in rows):
            self.log("  影像單號本來就一樣，不用合併")
            self.result["image"] = "✔ 已合併"
            merged = True
        else:
            self._menu(pat, "影像合併")
            self._settle({"image_merge"}, quiet=IMAGE_INFO_AFTER)
            self.ris.query_list(self.home_room, "")
            rows = self._rows_now(pat, after=True)
            merged = same(r.get("影像單號") for r in rows)
            info = self.ris.image_info
            if merged:
                self.log("  影像合併 → 影像單號都一樣" + (f"（影像資訊：{oneline(info, 160)}）" if info else ""))
                self.result["image"] = "✔ 合併成功"
            elif self.ris.no_image:
                self.log("  影像還沒到（RIS 說無影像），之後再試")
                self.result["image"] = "✗ 無影像"
            else:
                self.log("  影像合併沒有成功，影像可能還沒到" + (f"（影像資訊：{oneline(info, 120)}）" if info else ""))
                self.result["image"] = "✗ 沒合併成功"
        now_rooms = sorted({clean(r.get("檢查室")) for r in rows})
        if now_rooms != rooms:
            raise StopRun(f"做完之後檢查室從「{'、'.join(rooms)}」變成「{'、'.join(now_rooms)}」。"
                          "程式停下來，請你看一下這位。")
        self.log(f"  V {'報告、影像都合併好了' if merged else '報告合併好了，等影像'}"
                 f"（{time.time() - t0:.1f} 秒）；檢查室沒變")
        return "merged" if merged else "waiting"


class App(tk.Tk):
    MODE = "手動模式"
    COLS = ["病歷號", "姓名", "狀態", "CT", "醫令", "報告醫師", "報告合併", "影像合併", "程式會做"]

    def __init__(self):
        super().__init__()
        self.title(f"CT 多筆合併小工具（{self.MODE}）")
        self.geometry("1280x740")
        self._init_state()

        # 第一排：檢查室、讀取、監看
        top = ttk.Frame(self, padding=(6, 6, 6, 2))
        top.pack(fill="x")
        ttk.Label(top, text="檢查室").pack(side="left")
        self.room_var = tk.StringVar()
        self.cmb_room = ttk.Combobox(top, textvariable=self.room_var, state="readonly", width=24)
        self.cmb_room.pack(side="left", padx=(4, 8))
        self.btn_load = ttk.Button(top, text="讀取清單（只看）", command=self.on_load)
        self.btn_load.pack(side="left")
        self.btn_go = ttk.Button(top, text="開始監看", command=self.on_start)
        self.btn_go.pack(side="left", padx=6)
        self.btn_stop = ttk.Button(top, text="停止", command=self.on_stop, state="disabled")
        self.btn_stop.pack(side="left")
        ttk.Button(top, text="不用合併的醫令…", command=self.on_no_merge).pack(side="left", padx=(16, 0))
        # 第二排：設定
        opt = ttk.Frame(self, padding=(6, 0, 6, 4))
        opt.pack(fill="x")
        self.ask_each = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text="每位病人先問我", variable=self.ask_each).pack(side="left")
        self.cross_room = tk.BooleanVar(value=False)
        self.chk_cross = ttk.Checkbutton(opt, text="跨檢查室使用（開始檢查後拖回原檢查室）",
                                         variable=self.cross_room)
        self.chk_cross.pack(side="left", padx=12)
        ttk.Label(opt, text="閒置").pack(side="left")
        self.idle_var = tk.StringVar(value=str(IDLE_SEC))
        ttk.Spinbox(opt, from_=5, to=300, width=4, textvariable=self.idle_var).pack(side="left")
        ttk.Label(opt, text="秒才動　影像每").pack(side="left")
        self.retry_var = tk.StringVar(value=str(IMAGE_RETRY_MIN))
        ttk.Spinbox(opt, from_=1, to=60, width=3, textvariable=self.retry_var).pack(side="left")
        ttk.Label(opt, text="分鐘試一次（同時重新查詢清單）").pack(side="left")

        self._build_body("RIS 開在「登錄作業清單」，選好檢查室後按「讀取清單（只看）」看看程式會做什麼")
        self.after(300, self._load_rooms)

    def _init_state(self):
        self.q = queue.Queue()
        self.ris = Ris(log=lambda s: self.q.put(("log", s)))
        self.patients = []
        self.memory = {}             # {key: {"last": 時間, "manual": bool, "note": 說明}}
        self.busy = False
        self.watching = False
        self.stop_req = threading.Event()

    def _build_body(self, hint):
        """說明列、中間的病人表格、下面的記錄；關視窗、開始收背景訊息。"""
        self.lbl = ttk.Label(self, text=hint, padding=(8, 0))
        self.lbl.pack(fill="x")

        mid = ttk.Frame(self, padding=(6, 4))
        mid.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(mid, columns=self.COLS, show="headings", height=12)
        widths = {"病歷號": 85, "姓名": 65, "狀態": 70, "CT": 32, "醫令": 250, "報告醫師": 110,
                  "報告合併": 130, "影像合併": 190, "程式會做": 240}
        for c in self.COLS:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=widths[c], anchor="w")
        self.tree.tag_configure("skip", foreground="#888888")
        self.tree.tag_configure("todo", foreground="#0b57d0")
        self.tree.tag_configure("fail", foreground="#c62828")
        self.tree.tag_configure("done", foreground="#2e7d32")     # 報告、影像都合併好了
        self.tree.tag_configure("wait", foreground="#b26a00")     # 報告好了，影像還在等
        sb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")

        bot = ttk.Frame(self, padding=6)
        bot.pack(fill="both", expand=True)
        self.txt = tk.Text(bot, height=14, wrap="word")
        sb2 = ttk.Scrollbar(bot, orient="vertical", command=self.txt.yview)
        self.txt.configure(yscroll=sb2.set)
        self.txt.pack(side="left", fill="both", expand=True)
        sb2.pack(side="left", fill="y")

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.log(self.intro())
        self.after(120, self._drain)

    def _set_running(self, on):
        """背景在做事（讀清單、監看）時哪些按鈕能按。"""
        self.btn_load.config(state="disabled" if on else "normal")
        self.btn_go.config(state="disabled" if on else "normal")
        if not on:
            self.btn_stop.config(state="disabled")

    def _load_rooms(self):
        """從 RIS 清單的「檢查室」選單讀選項（背景讀，不點任何東西）。"""
        def job():
            try:
                self.ris.connect()
                rooms, cur = self.ris.list_rooms()
                self.q.put(("rooms", (rooms, cur)))
            except (StopRun, NotReady) as e:
                self.q.put(("log", f"讀不到 RIS 的檢查室選項：{e}"))
            except Exception as e:
                self.q.put(("log", f"讀不到 RIS 的檢查室選項：{type(e).__name__}: {e}"))
        threading.Thread(target=job, daemon=True).start()

    # ---------- 小工具 ----------
    def log(self, s):
        self.txt.insert("end", f"{datetime.now():%H:%M:%S}  {s}\n")
        self.txt.see("end")
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S}  {s}\n")

    def _drain(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "log":
                    self.log(payload)
                elif kind == "status":
                    self.lbl.config(text=payload)
                elif kind == "rooms":
                    rooms, cur = payload
                    self.cmb_room["values"] = rooms
                    if not self.room_var.get() and cur in rooms:
                        self.room_var.set(cur)
                elif kind == "patients":
                    self.patients = payload
                    self._fill()
                elif kind == "ask":
                    title, body, holder, ev, three = payload
                    self._front()
                    if three:
                        holder.append(messagebox.askyesnocancel(title, body, parent=self))
                    else:
                        holder.append(messagebox.askokcancel(title, body, parent=self))
                    self.attributes("-topmost", False)
                    ev.set()
                elif kind == "popup":
                    title, body, ok = payload
                    self._front()
                    (messagebox.showinfo if ok else messagebox.showwarning)(title, body, parent=self)
                    self.attributes("-topmost", False)
                elif kind == "call":
                    payload()                   # 背景執行緒要改畫面：交給這裡做
                elif kind == "done":
                    self.busy = self.watching = False
                    self._set_running(False)
        except queue.Empty:
            pass
        self.after(120, self._drain)

    def _front(self):
        self.deiconify()
        self.lift()
        self.attributes("-topmost", True)
        self.focus_force()

    def _run(self, fn):
        self.busy = True
        self._set_running(True)

        def wrap():
            try:
                fn()
            except Exception as e:
                self.q.put(("log", f"程式錯誤：{type(e).__name__}: {e}"))
                self.q.put(("log", traceback.format_exc()))
            finally:
                self.q.put(("done", None))
        threading.Thread(target=wrap, daemon=True).start()

    def _ask(self, title, body, three=False):
        """背景執行緒問使用者，等回答。"""
        holder, ev = [], threading.Event()
        self.q.put(("ask", (title, body, holder, ev, three)))
        ev.wait()
        return holder[0]

    def _int(self, var, default, lo, hi):
        try:
            return max(lo, min(hi, int(var.get())))
        except ValueError:
            return default

    # ---------- 表格 ----------
    # ---------- 手動／自動模式不一樣的地方 ----------
    def intro(self):
        return ("提醒：開始檢查會把檢查換到 RIS 目前登入的檢查室。在別間（例如 portable）幫忙時，"
                "請勾「跨檢查室使用」，程式會再把它們拖回選的檢查室。檢查完成不會按。")

    def plan(self, pats):
        return pats

    def summary(self, pats):
        cnt = collections.Counter(p["kind"] for p in pats)
        return (f"要開始檢查 {cnt['start']} 位、等影像 {cnt['merge']} 位、已合併等完成 {cnt['waiting']} 位、"
                f"有非 CT 不碰 {cnt['manual']} 位、不用合併的醫令 {cnt['nomerge']} 位")

    def what_text(self, p):
        return "開始檢查＋合併" if p["kind"] == "start" else "試合併影像"

    def make_runner(self, log, room, cross):
        return Runner(self.ris, log, lambda t, b: self._ask(t, b), cross_room=cross, home_room=room)

    def confirm_text(self, room, cross):
        return (f"檢查室：{room}" + ("　（跨檢查室使用）" if cross else "") + "\n\n"
                "程式會一直看「登錄作業清單」，電腦閒置的時候自動：\n"
                "  ‧ 已報到、有多筆 CT 的病人：開始檢查、合併報告、試著合併影像\n"
                "  ‧ 檢查中、影像還沒合併的：每隔幾分鐘再試一次影像合併\n"
                + (f"  ‧ 開始檢查後被換到這台登入的檢查室的，拖回「{room}」\n" if cross else "")
                + "檢查完成不會按，留給技術員。\n\n"
                "這些按下去就直接生效。請確認：\n"
                + ("" if cross else f"  ‧ 這台 RIS 登入的是「{room}」\n")
                + "  ‧ RIS 停在「登錄作業清單」\n\n"
                "確定要開始嗎？")

    def _plan_text(self, p):
        m = self.memory.get(p["key"], {})
        if m.get("manual"):
            return m.get("note") or "你選擇自己處理，不碰"
        if m.get("note"):
            return m["note"]
        return p["why"]

    def _fill(self):
        """中間的表格。程式處理過的病人，報告合併／影像合併欄顯示這次的結果（✔ 成功、✗ 沒成功）。"""
        self.tree.delete(*self.tree.get_children())
        order = {"start": 0, "merge": 1, "waiting": 2, "manual": 3, "nomerge": 4, "other": 5, "single": 6}
        for i, p in sorted(enumerate(self.patients), key=lambda x: (order[x[1]["kind"]], x[0])):
            n = len(p["accs"])
            m = self.memory.get(p["key"], {})
            rep = m.get("report") or (("已合併" if p["combined"] else "未合併") if n > 1 else "")
            img = m.get("image") or (("已合併" if p["imaged"] else "未合併") if n > 1 else "")
            vals = [p["patno"], p["name"], p["status"], n, "、".join(o[:26] for o in p["orders"]),
                    "、".join(p["doctors"]) or "（空白）", rep, img, self._plan_text(p)]
            if m.get("stopped"):
                tag = "fail"
            elif m.get("image", "").startswith("✔") or (p["kind"] == "waiting" and n > 1):
                tag = "done"
            elif m.get("image", "").startswith("✗"):
                tag = "wait"
            else:
                tag = "todo" if p["kind"] in ("start", "merge") else "skip"
            self.tree.insert("", "end", iid=str(i), values=vals, tags=(tag,))

    def _room(self):
        """選好的檢查室；沒選就提醒。"""
        room = self.room_var.get().strip()
        if not room:
            messagebox.showinfo("檢查室", "請先在左上角選檢查室。\n"
                                "（選單是空的話：RIS 切到「登錄作業清單」再重開小工具）", parent=self)
            if not self.cmb_room["values"]:
                self._load_rooms()
        return room

    def on_no_merge(self):
        """用記事本打開「不用合併的醫令」清單（存檔後下一次讀清單就生效，不用重開小工具）。"""
        n = rule_count(no_merge_orders())
        try:
            os.startfile(NO_MERGE_FILE)
        except OSError as e:
            messagebox.showerror("打不開", f"{NO_MERGE_FILE}\n{e}", parent=self)
            return
        self.log(f"打開不用合併的醫令清單（目前 {n} 個）：{NO_MERGE_FILE}　改完存檔，下一次讀清單就生效")

    # ---------- 讀取（只看） ----------
    def on_load(self):
        """把 RIS 清單設成選的檢查室、狀態全部，按查詢，讀回來。不碰任何病人。"""
        if self.busy:
            return
        room = self._room()
        if not room:
            return

        def job():
            try:
                self.ris.connect()
                if self.ris.read_form() is not None:
                    raise StopRun("「多筆登錄作業」現在開著，先關掉再讀清單。")
                self.ris.activate()
                self.ris.query_list(room, "")
                rows = self.ris.read_list()
            except (StopRun, NotReady) as e:
                self.q.put(("status", str(e)))
                self.q.put(("log", str(e)))
                return
            pats = self.plan(group_patients(rows))
            self.q.put(("patients", pats))
            self.q.put(("log", f"清單：檢查室「{room}」、狀態全部（已按查詢）"))
            msg = f"清單 {len(rows)} 筆、{len(pats)} 位病人：" + self.summary(pats)
            self.q.put(("status", msg))
            self.q.put(("log", msg))
        self._run(job)

    # ---------- 監看 ----------
    def on_start(self):
        if self.busy:
            return
        room = self._room()
        if not room:
            return
        idle = self._int(self.idle_var, IDLE_SEC, 5, 300)
        retry = self._int(self.retry_var, IMAGE_RETRY_MIN, 1, 60) * 60
        cross = self.cross_room.get()
        if not messagebox.askyesno("開始監看", self.confirm_text(room, cross)):
            self.log("已取消。")
            return
        ask_each = self.ask_each.get()

        def job():
            log = lambda s: self.q.put(("log", s))
            status = lambda s: self.q.put(("status", s))
            runner = self.make_runner(log, room, cross)
            log("=" * 60)
            log(f"開始監看（{self.MODE}）：檢查室「{room}」{'（跨檢查室使用）' if cross else ''}"
                f"（閒置 {idle} 秒才動、影像每 {retry // 60} 分鐘試一次、"
                f"{'每位先問' if ask_each else '不先問'}）")
            stopped = None
            last_skip = {}
            current = None          # 正在處理的病人（停下來時表格上標紅）
            queried_once = False
            # 清單的狀態要是全部（空白），不然開始檢查後的病人會從清單消失，回不來合併影像。
            # 第一次在下面的迴圈裡查（要等電腦閒置、RIS 在前面才能按查詢）
            last_query = 0.0
            try:
                self.ris.connect()
                self.ris.activate()
                self.ris.reset_idle()
                while not self.stop_req.is_set():
                    # 有人在用就不動
                    idle_now = self.ris.human_idle_sec()
                    if idle_now < idle:
                        status(f"監看中：有人在用電腦，再 {idle - idle_now:.0f} 秒沒人動才會動手")
                        time.sleep(1)
                        continue
                    if not self.ris.ris_in_front():
                        fg = _u32.GetForegroundWindow()
                        if fg and pid_of(fg) == os.getpid():
                            self.ris.activate()     # 小工具自己在前面：叫 RIS 回來
                        else:
                            status("監看中：RIS 不在最前面，等它回到前面")
                            time.sleep(LOOP_SEC)
                        continue
                    if self.ris.read_form() is not None:
                        status("監看中：「多筆登錄作業」開著，等它關掉")
                        time.sleep(LOOP_SEC)
                        continue
                    try:
                        # 每隔「影像每幾分鐘」重新查詢一次清單（使用者要求），也把篩選設回來
                        if time.time() - last_query >= retry:
                            status(f"監看中（{datetime.now():%H:%M:%S}）：重新查詢清單")
                            self.ris.query_list(room, "")
                            last_query = time.time()
                            if not queried_once:
                                log(f"清單：檢查室「{room}」、狀態全部（每 {retry // 60} 分鐘重新查詢一次）")
                                queried_once = True
                        rows = self.ris.read_list()
                    except NotReady as e:
                        status(f"監看中：{e}")
                        time.sleep(LOOP_SEC)
                        continue
                    pats = self.plan(group_patients(rows))
                    self.q.put(("patients", pats))
                    p = pick_next(pats, self.memory, time.time(), retry)
                    if not p:
                        cnt = collections.Counter(x["kind"] for x in pats)
                        status(f"監看中（{datetime.now():%H:%M:%S}）：沒有要做的。"
                               f"等影像 {cnt['merge']} 位、已合併等完成 {cnt['waiting']} 位")
                        time.sleep(LOOP_SEC)
                        continue
                    what = self.what_text(p)
                    if ask_each:
                        ans = self._ask("下一位",
                                        f"{p['name']}  {len(p['accs'])} 筆 CT（{p['status']}）：\n"
                                        + "\n".join(f"  {o}" for o in p["orders"])
                                        + f"\n\n要做：{what}\n\n"
                                        "是 = 做　否 = 這位不要碰　取消 = 停止監看", three=True)
                        if ans is None:
                            stopped = "你按了取消，停止監看。"
                            break
                        if ans is False:
                            self.memory[p["key"]] = {"manual": True}
                            log(f"照會 {'、'.join(p['accs'])}：你選擇這位不要碰")
                            continue
                        self.ris.activate()     # 回答完直接做（process 會重讀清單確認）
                    status(f"處理中：{what}")
                    log(f"[{what}] 照會 {'、'.join(p['accs'])}"
                        f"（{'、'.join(o[:24] for o in p['orders'])}）")
                    t_start = time.time()
                    current = p
                    try:
                        runner.process(p)
                        # 等影像的重試：從這次「開始」算一分鐘（使用者要每分鐘點開一次）；
                        # 剛開始檢查的：從做完算，給清單時間更新
                        m = self.memory.setdefault(p["key"], {})
                        now_s = f"{datetime.now():%H:%M}"
                        m["last"] = t_start if p["kind"] == "merge" else time.time()
                        if runner.result.get("report"):
                            m.setdefault("report", f"{runner.result['report']} {now_s}")
                        img = runner.result.get("image", "")
                        if img.startswith("✔"):
                            m["image"] = f"{img} {now_s}"
                            m["note"] = "報告、影像都合併好了；檢查完成留給技術員"
                        else:
                            m["tries"] = m.get("tries", 0) + 1
                            m["image"] = f"{img or '✗ 沒合併'}（第 {m['tries']} 次 {now_s}）"
                            m["note"] = f"等影像，每 {retry // 60} 分鐘再試"
                        self.q.put(("patients", pats))      # 表格馬上更新結果
                        current = None
                    except SkipPatient as e:
                        current = None
                        if last_skip.get(p["key"]) != str(e):
                            log(f"  - 這輪先不做：{e}")
                            last_skip[p["key"]] = str(e)
                        self.memory.setdefault(p["key"], {})["last"] = time.time() - REOPEN_COOLDOWN + LOOP_SEC
                    except ManualPatient as e:
                        current = None
                        self.memory[p["key"]] = {"manual": True, "note": f"不碰：{oneline(str(e), 60)}"}
                        log(f"  - {e}，這次監看不會再碰這位")
            except StopRun as e:
                stopped = str(e)
            except Exception as e:
                stopped = f"程式錯誤：{type(e).__name__}: {e}"
                log(traceback.format_exc())
            if stopped and current is not None:
                # 停在哪一位，表格上標紅
                m = self.memory.setdefault(current["key"], {})
                m["stopped"] = True
                m["note"] = "✗ 停在這位：" + oneline(stopped, 60)
                self.q.put(("patients", self.patients))
            if self.stop_req.is_set() and not stopped:
                log("已停止監看。")
            if stopped:
                log("  X " + oneline(stopped, 400))
                log("監看已停止。")
                self.q.put(("popup", ("監看停止了", stopped, False)))
            log("=" * 60)
            self.q.put(("status", "已停止監看"))

        self.stop_req.clear()
        self.watching = True
        self._run(job)
        self.btn_stop.config(state="normal")

    def on_stop(self):
        if self.watching:
            self.stop_req.set()
            self.log("收到「停止」：目前這一位做完就停。")

    def on_close(self):
        if self.watching and not messagebox.askyesno(
                "還在監看", "程式還在監看，現在關掉可能會讓某一位停在一半。\n確定要關嗎？",
                default="no"):
            return
        self.destroy()


class AutoApp(App):
    """自動模式：介面跟手動模式一樣，但每位病人是在清單上 Ctrl 選好 → 右鍵 → 合併報告／影像合併，
    不點開、不開始檢查，所以檢查室不會變（不需要跨檢查室）。"""
    MODE = "自動模式"

    def __init__(self):
        super().__init__()
        self.cross_room.set(False)
        self.chk_cross.pack_forget()

    def intro(self):
        return ("自動模式：在清單上選好同一位的 CT → 右鍵 →「合併報告」「影像合併」。"
                "不點開病人、不按開始檢查，檢查室不會變；開始檢查、檢查完成都留給技術員。")

    def plan(self, pats):
        return auto_plan(pats)

    def summary(self, pats):
        cnt = collections.Counter(p["kind"] for p in pats)
        return (f"要合併報告 {cnt['start']} 位、只差影像 {cnt['merge']} 位、都合併好了 {cnt['waiting']} 位、"
                f"有非 CT 不碰 {cnt['manual']} 位、不用合併的醫令 {cnt['nomerge']} 位")

    def what_text(self, p):
        return "右鍵合併報告" + ("＋影像" if not p["imaged"] else "") if p["kind"] == "start" else "右鍵試合併影像"

    def make_runner(self, log, room, cross):
        return AutoRunner(self.ris, log, lambda t, b: self._ask(t, b), home_room=room)

    def confirm_text(self, room, cross):
        return (f"檢查室：{room}（自動模式）\n\n"
                "程式會一直看「登錄作業清單」，電腦閒置的時候自動：\n"
                "  ‧ 有多筆 CT、報告還沒合併的：選好 → 右鍵 → 合併報告\n"
                "  ‧ 影像還沒合併的：選好 → 右鍵 → 影像合併（沒影像就每隔幾分鐘再試）\n"
                "不點開病人、不按開始檢查，檢查室不會變。\n"
                "開始檢查、檢查完成都不會按，留給技術員。\n\n"
                "這些按下去就直接生效。請確認 RIS 停在「登錄作業清單」。\n\n"
                "確定要開始嗎？")


class Chooser(tk.Tk):
    """啟動時選模式。"""

    def __init__(self):
        super().__init__()
        self.title("CT 多筆合併小工具")
        self.resizable(False, False)
        self.choice = None
        ttk.Label(self, text="要用哪一種模式？", padding=(16, 14, 16, 4),
                  font=("", 12, "bold")).pack()
        f = ttk.Frame(self, padding=16)
        f.pack()
        for col, (key, title, desc) in enumerate((
                ("manual", "手動模式", "點開病人 → 開始檢查 → 合併報告、影像\n可以跨檢查室（開始檢查後拖回原檢查室）"),
                ("auto", "自動模式", "清單上選好 → 右鍵 → 合併報告、影像合併\n不開始檢查，檢查室不會變"))):
            box = ttk.Frame(f, padding=8)
            box.grid(row=0, column=col, sticky="n")
            ttk.Button(box, text=title, width=18, command=lambda k=key: self._pick(k)).pack(ipady=10)
            ttk.Label(box, text=desc, justify="center", foreground="#555555", padding=(0, 6)).pack()
        self.protocol("WM_DELETE_WINDOW", self.destroy)

    def _pick(self, key):
        self.choice = key
        self.destroy()


def launch(mode=None):
    """啟動：沒指定模式就先跳出選擇視窗。"""
    if mode not in ("manual", "auto"):
        c = Chooser()
        c.mainloop()
        mode = c.choice
    if mode == "manual":
        App().mainloop()
    elif mode == "auto":
        AutoApp().mainloop()


if __name__ == "__main__":
    launch(sys.argv[1] if len(sys.argv) > 1 else None)
