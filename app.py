# -*- coding: utf-8 -*-
"""
DenpyoHub_Miya
制作伝票一覧作成デスクトップアプリケーション
日報ファイル(CSV)、購入品(XLSM/XLSX)、外注加工(XLSM/XLSX)、受注表(XLSM/XLSX)から
管理Noごとの作業時間・金額・購入品費・外注加工費・各小計を算出し、制作伝票の一覧を作成します。
"""

import os
import io
import sys
import csv
import math
import stat
import json
import time
import queue
import shutil
import tempfile
import threading
import zipfile
import datetime
import importlib
import contextlib
import unicodedata
import xml.etree.ElementTree as ET
from collections import defaultdict, Counter, OrderedDict
from concurrent.futures import ThreadPoolExecutor, as_completed
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Windows 11 高DPIスケーリング対応
if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

APP_NAME = "DenpyoHub_Miya"
APP_TITLE = "DenpyoHub_Miya - 制作伝票一覧集計システム"
CONFIG_FILENAME = "denpyohub_config.json"
LEGACY_CONFIG_FILENAME = "stockmatch_config.json"
SEIZO_CACHE_FILENAME = "denpyohub_seizo_cache.json"
SEIZO_CACHE_VERSION = 2
SEIZO_TARGET_ROWS = frozenset({2, 5, 6, 9, 11, 12, 14, 17, 18, 19, 21, 28, 29, 30, 31, 32, 33, 34, 35, 36, 38})

# デフォルトの作業コード範囲設定
DEFAULT_CODE_RANGES = {
    "作業時間": "20-49",
    "移動時間": "90-99",
    "修理改造時間": "50-69",
    "設計関係": "1-19",
    "その他作業": "70-89",
}

# 表示単位の定義（デフォルトは金額(円)）
VALID_TIME_UNITS = ("金額(円)", "分(そのまま)", "時間(分÷60)", "工数(8h=1)")
KOUSU_MINUTES = 480.0  # 1工数 = 8時間 = 480分
DEFAULT_TIME_UNIT = "金額(円)"

# デフォルトの作業単価レート（円/分）
DEFAULT_WORK_RATES = {
    "作業時間": 62.5,
    "移動時間": 50.0,
    "修理改造時間": 70.0,
    "設計関係": 62.5,
    "その他作業": 62.5,
}

# 作業区分の順序（デフォルト）
WORK_CATEGORIES = [
    "作業時間",
    "移動時間",
    "修理改造時間",
    "設計関係",
    "その他作業",
]

# 制作伝票（個別Excel）および自動算出で追加する項目列
SEIZO_COLUMNS = [
    "調整用",
    "材料費１(型鋼)",
    "材料費２(鉄板)",
    "材料費３(コラム)",
    "その他材料費(31行)",
    "その他材料費(32行)",
    "その他材料費(33行)",
    "切断材料",
    "材料費計",
    "設計費",
    "小計1",
    "消耗品費(6%)",
    "重機、運搬費",
    "車両、交通費",
    "出張費",
    "管理費",
    "小計2",
    "小計1+小計2",
    "諸経費(8%)",
    "合計",
    "合計/数量",
    "調整額",
    "総請求金額",
    "総請求金額/数量",
]
# 作業費小計（単価レート連動・円換算）等から動的算出する常に円表示の列
DERIVED_YEN_COLUMNS = (
    "小計1",
    "消耗品費(6%)",
    "管理費",
    "小計2",
    "小計1+小計2",
    "諸経費(8%)",
    "合計",
    "合計/数量",
    "調整額",
    "総請求金額/数量",
)
SEIZO_COST_COLUMNS = tuple(c for c in SEIZO_COLUMNS if c not in DERIVED_YEN_COLUMNS)
COST_YEN_COLUMNS = {"購入品", "外注加工", "購入品小計"} | set(SEIZO_COST_COLUMNS)
WORK_ONLY_COLUMNS = frozenset(WORK_CATEGORIES) | frozenset({"作業小計"})
YEN_ONLY_COLUMNS = frozenset(COST_YEN_COLUMNS) | frozenset(DERIVED_YEN_COLUMNS)
RATE_DEPENDENT_COLUMNS = set(WORK_CATEGORIES) | {"作業小計"} | set(DERIVED_YEN_COLUMNS)

# 主要Excelファイル読み込み時に解析対象とする必要列セット（1行目見出しは検証用に全列取得）
KOUNYU_TARGET_COLS = frozenset({"A", "B", "C", "D", "E", "F", "G", "H", "I", "K", "L", "M", "N", "Q", "T", "U", "W"})
JUCHU_TARGET_COLS = frozenset({
    "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O",
    "P", "Q", "R", "S", "T", "U", "V", "W", "Y", "AB", "AC", "AD",
})

# 高速文字変換テーブル（ループ内での毎回生成を回避）
_KANRI_TRANS = str.maketrans(
    "０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ",
    "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz",
)
_RANGE_TRANS = str.maketrans("０１２３４５６７８９～〜－ー，、", "0123456789~~--,,")

# 算出列の順序（デフォルト）
CALC_COLUMNS = [
    "作業時間",
    "移動時間",
    "修理改造時間",
    "設計関係",
    "その他作業",
    "作業小計",
    "購入品",
    "外注加工",
    "購入品小計",
] + SEIZO_COLUMNS

# 基本情報列（デフォルト）
BASE_COLUMNS = [
    "管理No",
    "受注日",
    "受注先",
    "出荷先",
    "プロジェクト名",
    "品名、ライン名等",
    "製品詳細",
    "数量",
    "単位",
    "製造又は手配先",
    "売り上げ区分",
    "出荷日",
    "請求完了日",
]

# 小計1・小計2を構成する内訳列（デフォルトでは非表示にし、ボタンで展開・折りたたみ可能）
SUBTOTAL1_DETAIL_COLUMNS = (
    "作業時間",
    "移動時間",
    "修理改造時間",
    "設計関係",
    "その他作業",
    "作業小計",
    "購入品",
    "外注加工",
    "購入品小計",
    "調整用",
    "材料費１(型鋼)",
    "材料費２(鉄板)",
    "材料費３(コラム)",
    "その他材料費(31行)",
    "その他材料費(32行)",
    "その他材料費(33行)",
    "切断材料",
    "材料費計",
    "設計費",
)
SUBTOTAL2_DETAIL_COLUMNS = (
    "消耗品費(6%)",
    "重機、運搬費",
    "車両、交通費",
    "出張費",
    "管理費",
)
DEFAULT_HIDDEN_COLUMNS = set(SUBTOTAL1_DETAIL_COLUMNS) | set(SUBTOTAL2_DETAIL_COLUMNS) | {"備考"}

# 利用可能な全列（デフォルト表示順は BASE_COLUMNS + CALC_COLUMNS、オプションで「備考」も追加可能）
DEFAULT_COLUMN_ORDER = BASE_COLUMNS + CALC_COLUMNS
DEFAULT_VISIBLE_COLUMNS = [c for c in DEFAULT_COLUMN_ORDER if c not in DEFAULT_HIDDEN_COLUMNS]
ALL_AVAILABLE_COLUMNS = BASE_COLUMNS + CALC_COLUMNS + ["備考"]

NUMERIC_COLUMNS = set(CALC_COLUMNS) | {"数量"}

DEFAULT_COL_WIDTHS = {
    "管理No": 96,
    "受注日": 92,
    "受注先": 90,
    "出荷先": 100,
    "プロジェクト名": 95,
    "品名、ライン名等": 105,
    "製品詳細": 210,
    "数量": 58,
    "単位": 52,
    "製造又は手配先": 92,
    "売り上げ区分": 125,
    "出荷日": 92,
    "請求完了日": 92,
    "備考": 180,
    "作業時間": 106,
    "移動時間": 106,
    "修理改造時間": 128,
    "設計関係": 106,
    "その他作業": 130,
    "作業小計": 122,
    "購入品": 114,
    "外注加工": 114,
    "購入品小計": 126,
    "調整用": 106,
    "材料費１(型鋼)": 132,
    "材料費２(鉄板)": 132,
    "材料費３(コラム)": 140,
    "その他材料費(31行)": 146,
    "その他材料費(32行)": 146,
    "その他材料費(33行)": 146,
    "切断材料": 114,
    "材料費計": 114,
    "設計費": 110,
    "小計1": 124,
    "消耗品費(6%)": 138,
    "重機、運搬費": 138,
    "車両、交通費": 138,
    "出張費": 110,
    "管理費": 114,
    "小計2": 124,
    "小計1+小計2": 132,
    "諸経費(8%)": 128,
    "合計": 124,
    "合計/数量": 124,
    "調整額": 118,
    "総請求金額": 130,
    "総請求金額/数量": 146,
}

# 制作伝票項目の共通見出し（常に円表示）
SEIZO_DISPLAY_LABELS = {
    "調整用": "調整用(円)",
    "材料費１(型鋼)": "材料費１(型鋼)",
    "材料費２(鉄板)": "材料費２(鉄板)",
    "材料費３(コラム)": "材料費３(コラム)",
    "その他材料費(31行)": "その他材料費(31行)",
    "その他材料費(32行)": "その他材料費(32行)",
    "その他材料費(33行)": "その他材料費(33行)",
    "切断材料": "切断材料(円)",
    "材料費計": "材料費計(円)",
    "設計費": "設計費(円)",
    "小計1": "小計1(円)",
    "消耗品費(6%)": "消耗品費(6%)(円)",
    "重機、運搬費": "重機、運搬費(円)",
    "車両、交通費": "車両、交通費(円)",
    "出張費": "出張費(円)",
    "管理費": "管理費(円)",
    "小計2": "小計2(円)",
    "小計1+小計2": "小計1+小計2(円)",
    "諸経費(8%)": "諸経費(8%)(円)",
    "合計": "合計(円)",
    "合計/数量": "合計/数量(円)",
    "調整額": "調整額(円)",
    "総請求金額": "総請求金額(円)",
    "総請求金額/数量": "総請求金額/数量(円)",
}

# 表示単位ごとの見出し表記マッピング（内部キーは変えずに画面見出しと出力見出しを分かりやすく変換）
DISPLAY_COL_LABELS = {
    "金額(円)": {
        "作業時間": "作業費(円)",
        "移動時間": "移動費(円)",
        "修理改造時間": "修理改造費(円)",
        "設計関係": "設計費(関係)(円)",
        "その他作業": "その他作業費(円)",
        "作業小計": "作業費小計(円)",
        "購入品": "購入品費(円)",
        "外注加工": "外注加工費(円)",
        "購入品小計": "購入品小計(円)",
        **SEIZO_DISPLAY_LABELS,
    },
    "分(そのまま)": {
        "作業時間": "作業時間(分)",
        "移動時間": "移動時間(分)",
        "修理改造時間": "修理改造時間(分)",
        "設計関係": "設計関係(分)",
        "その他作業": "その他作業(分)",
        "作業小計": "作業小計(分)",
        "購入品": "購入品費(円)",
        "外注加工": "外注加工費(円)",
        "購入品小計": "購入品小計(円)",
        **SEIZO_DISPLAY_LABELS,
    },
    "時間(分÷60)": {
        "作業時間": "作業時間(h)",
        "移動時間": "移動時間(h)",
        "修理改造時間": "修理改造時間(h)",
        "設計関係": "設計関係(h)",
        "その他作業": "その他作業(h)",
        "作業小計": "作業小計(h)",
        "購入品": "購入品費(円)",
        "外注加工": "外注加工費(円)",
        "購入品小計": "購入品小計(円)",
        **SEIZO_DISPLAY_LABELS,
    },
    "工数(8h=1)": {
        "作業時間": "作業工数(8h)",
        "移動時間": "移動工数(8h)",
        "修理改造時間": "修理改造工数(8h)",
        "設計関係": "設計工数(8h)",
        "その他作業": "その他工数(8h)",
        "作業小計": "作業工数小計(8h)",
        "購入品": "購入品費(円)",
        "外注加工": "外注加工費(円)",
        "購入品小計": "購入品小計(円)",
        **SEIZO_DISPLAY_LABELS,
    },
}


def get_display_column_label(col_name: str, time_unit: str) -> str:
    """表示単位に応じた見出し名（金額時は「作業費(円)」等）を返す"""
    return DISPLAY_COL_LABELS.get(time_unit, {}).get(col_name, col_name)

# ライトモード固定カラーパレット（システムがダークモードでも常にライトモードで表示）
COLORS = {
    "bg_main": "#F8FAFC",
    "bg_card": "#FFFFFF",
    "bg_header": "#EFF6FF",
    "bg_sub": "#F1F5F9",
    "fg_main": "#0F172A",
    "fg_sub": "#475569",
    "fg_muted": "#64748B",
    "accent": "#2563EB",
    "accent_hover": "#1D4ED8",
    "accent_light": "#DBEAFE",
    "success": "#16A34A",
    "success_hover": "#15803D",
    "warning": "#D97706",
    "warning_bg": "#FEF3C7",
    "danger": "#DC2626",
    "border": "#CBD5E1",
    "row_even": "#FFFFFF",
    "row_odd": "#F8FAFC",
    "subtotal_bg": "#FEF9C3",
    "total_bg": "#E0F2FE",
}

NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def get_base_dir() -> str:
    """実行ファイルまたはスクリプトの基準ディレクトリを取得する"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_resource_path(relative_path: str) -> str:
    """PyInstallerバンドルリソースまたはローカルファイルのパスを取得する"""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        bundled = os.path.join(sys._MEIPASS, relative_path)
        if os.path.exists(bundled):
            return bundled
    return os.path.join(get_base_dir(), relative_path)


def normalize_kanri_no(val) -> str:
    """管理Noを正規化する（空白除去、浮動小数点表記.0の除去、全角→半角）"""
    if val is None:
        return ""
    s = val.strip() if isinstance(val, str) else str(val).strip()
    if not s:
        return ""
    s = s.translate(_KANRI_TRANS)
    if s.endswith(".0"):
        prefix = s[:-2]
        if prefix.lstrip("-").isdigit():
            s = prefix
    return s


def normalize_search_text(val) -> str:
    """
    検索・フィルター照合用に文字列を正規化する。
    NFKC正規化により、半角カナ（ﾎｯﾊﾟｰ等）と全角カナ（ホッパー等）、
    全角英数記号（（株）サカタ等）と半角英数記号（(株)サカタ等）を同一視して検索しやすくする。
    """
    if val is None:
        return ""
    s = val.strip() if isinstance(val, str) else str(val).strip()
    if not s:
        return ""
    return unicodedata.normalize("NFKC", s).lower()


def to_float(val, default: float = 0.0) -> float:
    """安全に有限のfloatへ変換する（NaN/Inf・エラー文字列を除外）"""
    if val is None:
        return default
    if isinstance(val, (int, float)):
        f = float(val)
        return f if math.isfinite(f) else default
    s = (val if isinstance(val, str) else str(val)).strip().replace(",", "")
    if not s or s.startswith("#"):
        return default
    try:
        f = float(s)
        return f if math.isfinite(f) else default
    except ValueError:
        return default


def format_excel_date(val) -> str:
    """Excelシリアル値または日付文字列をYYYY/MM/DD形式に整形する"""
    if val is None:
        return ""
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime("%Y/%m/%d")
    s = str(val).strip()
    if not s or s == "0":
        return ""
    try:
        num = float(s)
        if 30000 <= num <= 65000:
            dt = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(num))
            return dt.strftime("%Y/%m/%d")
    except ValueError:
        pass
    return s


def format_file_mtime(filepath: str) -> str:
    """ファイルの更新日時を取得してフォーマットする"""
    if not filepath or not os.path.exists(filepath):
        return "ファイルなし"
    try:
        mtime = os.path.getmtime(filepath)
        dt = datetime.datetime.fromtimestamp(mtime)
        return dt.strftime("%Y/%m/%d %H:%M:%S")
    except Exception:
        return "取得不可"


def summarize_file_list_mtime(file_paths: list, extra_info: str = "") -> str:
    """複数ファイルの更新日時サマリー文字列を生成する"""
    if not file_paths:
        return "未選択"
    existing = [p for p in file_paths if p and os.path.exists(p)]
    missing_cnt = len(file_paths) - len(existing)
    if not existing:
        return f"ファイルが見つかりません ({len(file_paths)}件)"
    mtimes = []
    for p in existing:
        try:
            mtimes.append(os.path.getmtime(p))
        except OSError:
            pass
    if not mtimes:
        return f"更新日時取得不可 ({len(file_paths)}件)"
    latest_ts = max(mtimes)
    latest_str = datetime.datetime.fromtimestamp(latest_ts).strftime("%Y/%m/%d %H:%M:%S")
    if len(file_paths) == 1:
        base = f"更新日: {latest_str}"
    else:
        base = f"最新更新: {latest_str} (全{len(file_paths)}件"
        if missing_cnt > 0:
            base += f" / 不在{missing_cnt}件"
        base += ")"
    if extra_info:
        base += f" {extra_info}"
    return base


def summarize_file_paths_text(file_paths: list) -> str:
    """選択済みファイルリストの表示用要約文字列を作成する"""
    if not file_paths:
        return ""
    if len(file_paths) == 1:
        return file_paths[0]
    names = [os.path.basename(p) for p in file_paths]
    joined = ", ".join(names)
    if len(joined) > 95:
        joined = joined[:92] + "..."
    return f"[{len(file_paths)}件選択] {joined}"


def parse_code_range_spec(spec_str: str):
    """
    '20～49' や '20-49, 100-105' などの文字列から、対象コード判定用のタプルリスト [(start, end), ...] を返す
    """
    if not spec_str:
        return []
    s = str(spec_str).strip().translate(_RANGE_TRANS)
    ranges = []
    for part in s.split(","):
        part = part.strip()
        if not part:
            continue
        sep = None
        for token in ["~", "-"]:
            if token in part:
                sep = token
                break
        if sep:
            sub = part.split(sep, 1)
            try:
                start = int(float(sub[0].strip()))
                end = int(float(sub[1].strip()))
                if start > end:
                    start, end = end, start
                ranges.append((start, end))
            except ValueError:
                continue
        else:
            try:
                v = int(float(part))
                ranges.append((v, v))
            except ValueError:
                continue
    return ranges


def match_code_category(code_int: int, compiled_ranges: dict) -> str:
    """作業コードがどの作業区分に該当するか判定する"""
    for cat in WORK_CATEGORIES:
        for start, end in compiled_ranges.get(cat, []):
            if start <= code_int <= end:
                return cat
    return ""


class FormatValidationError(ValueError):
    """サンプルファイルとフォーマットが異なる場合に送出する構造化エラー"""

    def __init__(self, category: str, filepath: str, reason: str, expected: str = "", detected: str = ""):
        self.category = category
        self.filepath = filepath
        self.reason = reason
        self.expected = expected
        self.detected = detected
        fname = os.path.basename(filepath) if filepath else "(不明)"
        msg = f"【フォーマット不一致: {category}】{fname} - {reason}"
        if expected:
            msg += f" (期待形式: {expected})"
        if detected:
            msg += f" [検出内容: {detected}]"
        super().__init__(msg)

    def to_dict(self) -> dict:
        return {
            "category": self.category,
            "file": os.path.basename(self.filepath) if self.filepath else "",
            "path": self.filepath or "",
            "kind": "フォーマット不一致",
            "reason": self.reason,
            "expected": self.expected,
            "detected": self.detected,
        }


def _sleep_cancellable(seconds: float, cancel_check=None):
    """キャンセル判定を挟みながら短時間待機する（他ユーザー保存中の一時ロック解除待ち用）"""
    end_t = time.perf_counter() + max(0.0, float(seconds))
    while True:
        if cancel_check and cancel_check():
            raise InterruptedError("処理がキャンセルされました")
        rem = end_t - time.perf_counter()
        if rem <= 0:
            break
        time.sleep(min(0.05, rem))


def _calc_kounyu_row_amount(r: dict) -> float:
    """購入品・外注加工の1行金額を算出する共通コア（Q列優先、空欄時は 単価M×必要数N の整数値）"""
    q_val = r.get("Q")
    if q_val is not None and str(q_val).strip() != "":
        return to_float(q_val, 0.0)
    return float(int(to_float(r.get("M"), 0.0) * to_float(r.get("N"), 0.0)))


def _to_win32_extended_path(filepath: str) -> str:
    """260文字超の深いフォルダ（Google Drive / NAS UNC等）でもCreateFileWで開けるよう拡張長パス形式へ変換する"""
    abs_p = os.path.abspath(os.path.normpath(filepath))
    if abs_p.startswith("\\\\?\\") or len(abs_p) < 248:
        return abs_p
    if abs_p.startswith("\\\\"):
        return "\\\\?\\UNC\\" + abs_p[2:]
    return "\\\\?\\" + abs_p


_WIN32_CREATE_FILE = None
_WIN32_CLOSE_HANDLE = None
_WIN32_INVALID_HANDLE_VALUE = None
if sys.platform == "win32":
    try:
        import ctypes as _ct_init
        from ctypes import wintypes as _wt_init

        _k32_init = _ct_init.WinDLL("kernel32", use_last_error=True)
        _WIN32_CREATE_FILE = _k32_init.CreateFileW
        _WIN32_CREATE_FILE.argtypes = [
            _wt_init.LPCWSTR,
            _wt_init.DWORD,
            _wt_init.DWORD,
            _ct_init.c_void_p,
            _wt_init.DWORD,
            _wt_init.DWORD,
            _wt_init.HANDLE,
        ]
        _WIN32_CREATE_FILE.restype = _ct_init.c_void_p
        _WIN32_CLOSE_HANDLE = _k32_init.CloseHandle
        _WIN32_CLOSE_HANDLE.argtypes = [_ct_init.c_void_p]
        _WIN32_CLOSE_HANDLE.restype = _wt_init.BOOL
        _WIN32_INVALID_HANDLE_VALUE = _ct_init.c_void_p(-1).value
    except Exception:
        _WIN32_CREATE_FILE = None
        _WIN32_CLOSE_HANDLE = None
        _WIN32_INVALID_HANDLE_VALUE = None


def _open_win32_shared_readonly_file(filepath: str):
    """
    Windows 11 で他のユーザーやExcelが開いているファイルでも支障なく読み取りでき、
    かつこちらの読み込みが他ユーザーの保存（名前変更・削除）を阻害しないよう
    FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE で読み取り専用オープンする。
    """
    if _WIN32_CREATE_FILE is None or _WIN32_CLOSE_HANDLE is None:
        return open(filepath, "rb")

    import ctypes
    import msvcrt

    GENERIC_READ = 0x80000000
    FILE_SHARE_ALL = 0x00000001 | 0x00000002 | 0x00000004  # READ | WRITE | DELETE
    OPEN_EXISTING = 3
    FILE_FLAGS = 0x00000080 | 0x08000000  # FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN

    win_path = _to_win32_extended_path(filepath)
    handle = _WIN32_CREATE_FILE(win_path, GENERIC_READ, FILE_SHARE_ALL, None, OPEN_EXISTING, FILE_FLAGS, None)
    if handle is None or handle == _WIN32_INVALID_HANDLE_VALUE:
        err = ctypes.get_last_error()
        if err in (2, 3):
            raise FileNotFoundError(f"ファイルが見つかりません: {filepath}")
        if err in (5, 32, 33):
            exc = PermissionError(13, f"他のプロセスがファイルを使用中です (WinError {err})", filepath)
            exc.winerror = err
            raise exc
        raise ctypes.WinError(err)

    try:
        fd = msvcrt.open_osfhandle(handle, os.O_RDONLY | os.O_BINARY)
    except Exception:
        try:
            _WIN32_CLOSE_HANDLE(handle)
        except Exception:
            pass
        raise
    try:
        return os.fdopen(fd, "rb")
    except Exception:
        try:
            os.close(fd)
        except Exception:
            pass
        raise


@contextlib.contextmanager
def open_shared_readonly(filepath: str, cancel_check=None, retries: int = 4):
    """
    他のユーザー・Excelが開いているファイルでも支障なく読み取れる共有読み取りコンテキストマネージャ。
    - Windows環境では FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE でオープンし、
      他ユーザーの閲覧・上書き保存を一切ブロックしない。
    - 他ユーザーがちょうど保存中で一時的な排他ロック(WinError 32/33)が発生した場合は、
      短い待機（指数バックオフ・キャンセル対応）で自動リトライして透過的に読み込む。
    """
    if cancel_check and cancel_check():
        raise InterruptedError("処理がキャンセルされました")
    if not filepath or not os.path.exists(filepath):
        raise FileNotFoundError(f"ファイルが見つかりません: {filepath}")
    if os.path.isdir(filepath):
        raise IsADirectoryError(f"ファイルではなくフォルダが指定されています: {filepath}")

    backoffs = (0.15, 0.35, 0.70, 1.20)
    last_err = None
    f_obj = None
    for attempt in range(max(1, retries + 1)):
        if cancel_check and cancel_check():
            raise InterruptedError("処理がキャンセルされました")
        try:
            if sys.platform == "win32":
                f_obj = _open_win32_shared_readonly_file(filepath)
            else:
                f_obj = open(filepath, "rb")
            break
        except (PermissionError, OSError) as exc:
            win_err = getattr(exc, "winerror", None)
            is_retryable = isinstance(exc, PermissionError) or (win_err in (5, 32, 33))
            if not is_retryable or attempt >= retries:
                last_err = exc
                break
            last_err = exc
            wait_s = backoffs[min(attempt, len(backoffs) - 1)]
            _sleep_cancellable(wait_s, cancel_check=cancel_check)

    if f_obj is None:
        if last_err is not None:
            raise last_err
        raise PermissionError(f"ファイルを開けません: {filepath}")

    try:
        yield f_obj
    finally:
        try:
            f_obj.close()
        except Exception:
            pass


def open_readonly_bytes(src_path: str, cancel_check=None) -> bytes:
    """
    参照元ファイルを一切変更・ロックせず、共有読み取り専用でメモリへ高速に読み込む。
    他ユーザーがExcelで開いている状態や保存直後の一時ロックにも対応する。
    """
    backoffs = (0.20, 0.50, 1.00)
    last_err = None
    for attempt in range(len(backoffs) + 1):
        if cancel_check and cancel_check():
            raise InterruptedError("処理がキャンセルされました")
        try:
            chunks = []
            with open_shared_readonly(src_path, cancel_check=cancel_check) as f_src:
                while True:
                    if cancel_check and cancel_check():
                        raise InterruptedError("処理がキャンセルされました")
                    buf = f_src.read(1024 * 1024)
                    if not buf:
                        break
                    chunks.append(buf)
            return b"".join(chunks)
        except (InterruptedError, FileNotFoundError, IsADirectoryError):
            raise
        except (PermissionError, OSError) as exc:
            win_err = getattr(exc, "winerror", None)
            is_retryable = isinstance(exc, PermissionError) or (win_err in (5, 32, 33))
            if not is_retryable or attempt >= len(backoffs):
                raise
            last_err = exc
            _sleep_cancellable(backoffs[attempt], cancel_check=cancel_check)
    if last_err is not None:
        raise last_err
    return b""


def _load_shared_strings_from_zip(z: zipfile.ZipFile, namelist_set: set, cancel_check=None) -> list:
    """ZIP内の xl/sharedStrings.xml をメモリ上で高速に解析して文字列リストを返す"""
    shared_strings = []
    if "xl/sharedStrings.xml" not in namelist_set:
        return shared_strings
    raw_ss = z.read("xl/sharedStrings.xml")
    if len(raw_ss) <= 8 * 1024 * 1024:
        ss_root = ET.fromstring(raw_ss)
        t_tag = f"{NS_MAIN}t"
        r_tag = f"{NS_MAIN}r"
        for idx, elem in enumerate(ss_root.findall(f"{NS_MAIN}si")):
            if cancel_check and (idx & 2047) == 0 and cancel_check():
                raise InterruptedError("処理がキャンセルされました")
            t_direct = elem.find(t_tag)
            if t_direct is not None:
                shared_strings.append(t_direct.text or "")
            else:
                parts = []
                for r_el in elem.findall(r_tag):
                    t_el = r_el.find(t_tag)
                    if t_el is not None and t_el.text:
                        parts.append(t_el.text)
                shared_strings.append("".join(parts))
        return shared_strings

    with io.BytesIO(raw_ss) as sf:
        si_count = 0
        ss_root = None
        for event, elem in ET.iterparse(sf, events=("start", "end")):
            if event == "start" and ss_root is None:
                ss_root = elem
            elif event == "end" and elem.tag == f"{NS_MAIN}si":
                if cancel_check and (si_count & 1023) == 0 and cancel_check():
                    raise InterruptedError("処理がキャンセルされました")
                parts = []
                t_direct = elem.find(f"{NS_MAIN}t")
                if t_direct is not None and t_direct.text:
                    parts.append(t_direct.text)
                for r_el in elem.findall(f"{NS_MAIN}r"):
                    t_el = r_el.find(f"{NS_MAIN}t")
                    if t_el is not None and t_el.text:
                        parts.append(t_el.text)
                shared_strings.append("".join(parts))
                elem.clear()
                si_count += 1
                if ss_root is not None and (si_count & 1023) == 0:
                    ss_root.clear()
    return shared_strings


def _read_excel_sheet_rows_from_zip(
    filepath: str,
    z: zipfile.ZipFile,
    preferred_sheets=None,
    cancel_check=None,
    target_cols=None,
):
    """開かれた ZipFile から対象シートの行を高速ストリーミング抽出する共通コア"""
    namelist_set = set(z.namelist())
    if "xl/workbook.xml" not in namelist_set or "xl/_rels/workbook.xml.rels" not in namelist_set:
        raise FormatValidationError(
            "Excelファイル",
            filepath,
            "Excelブック構造（xl/workbook.xml）が見つかりません。",
            "標準のExcelブック (.xlsm / .xlsx)",
            "workbook.xml 不在のZIPファイル",
        )

    wb_root = ET.fromstring(z.read("xl/workbook.xml"))
    rels_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rels = {r.attrib["Id"]: r.attrib["Target"] for r in rels_root if "Id" in r.attrib and "Target" in r.attrib}

    sheet_target = None
    matched_sheet_name = ""
    first_target = None
    first_sheet_name = ""
    all_sheet_names = []
    for sheet in wb_root.findall(f".//{NS_MAIN}sheet"):
        name = sheet.attrib.get("name", "")
        all_sheet_names.append(name)
        rid = sheet.attrib.get(f"{NS_REL}id", "")
        raw_target = rels.get(rid, "").lstrip("/")
        tgt = raw_target if raw_target.startswith("xl/") else f"xl/{raw_target}"
        if tgt not in namelist_set and raw_target in namelist_set:
            tgt = raw_target
        if first_target is None and tgt in namelist_set:
            first_target = tgt
            first_sheet_name = name
        if preferred_sheets and name in preferred_sheets and tgt in namelist_set:
            sheet_target = tgt
            matched_sheet_name = name
            break

    if sheet_target is None:
        sheet_target = first_target
        matched_sheet_name = first_sheet_name
    if not sheet_target:
        return [], "", all_sheet_names

    shared_strings = _load_shared_strings_from_zip(z, namelist_set, cancel_check=cancel_check)

    rows = []
    c_tag = f"{NS_MAIN}c"
    v_tag = f"{NS_MAIN}v"
    is_tag = f"{NS_MAIN}is"
    t_tag = f"{NS_MAIN}t"
    row_tag = f"{NS_MAIN}row"
    sd_tag = f"{NS_MAIN}sheetData"
    ss_len = len(shared_strings)

    with z.open(sheet_target) as wf:
        sheet_data_el = None
        row_count = 0
        for event, elem in ET.iterparse(wf, events=("start", "end")):
            if event == "start" and elem.tag == sd_tag:
                sheet_data_el = elem
            elif event == "end" and elem.tag == row_tag:
                if cancel_check and (row_count & 511) == 0 and cancel_check():
                    raise InterruptedError("処理がキャンセルされました")
                # 1行目（見出し行）はフォーマット検証と詳細診断のため全列を保持し、2行目以降は必要列のみ高速抽出
                col_filter = target_cols if (target_cols is not None and rows) else None
                row_vals = {}
                for c_el in elem:
                    if c_el.tag != c_tag:
                        continue
                    attrib = c_el.attrib
                    ref = attrib.get("r", "")
                    if not ref:
                        continue
                    ref_len = len(ref)
                    if ref_len >= 2 and ref[1] <= "9":
                        col = ref[0]
                    elif ref_len >= 3 and ref[2] <= "9":
                        col = ref[:2]
                    else:
                        col = ref.rstrip("0123456789")
                        if not col:
                            continue
                    if col_filter is not None and col not in col_filter:
                        continue
                    t_attr = attrib.get("t", "")
                    if t_attr == "e":
                        continue
                    val = None
                    if t_attr == "s":
                        v_el = c_el.find(v_tag)
                        if v_el is not None and v_el.text:
                            try:
                                idx = int(v_el.text)
                                if 0 <= idx < ss_len:
                                    val = shared_strings[idx]
                            except ValueError:
                                pass
                    elif t_attr == "inlineStr":
                        is_el = c_el.find(is_tag)
                        if is_el is not None:
                            val = "".join(x.text or "" for x in is_el.iter(t_tag))
                    else:
                        v_el = c_el.find(v_tag)
                        if v_el is not None:
                            val = v_el.text
                    if val is not None:
                        s_val = val.strip() if isinstance(val, str) else str(val).strip()
                        if s_val:
                            row_vals[col] = s_val
                if row_vals:
                    rows.append(row_vals)
                elem.clear()
                row_count += 1
                if sheet_data_el is not None and (row_count & 511) == 0:
                    sheet_data_el.clear()

    return rows, matched_sheet_name, all_sheet_names


def read_excel_sheet_rows(
    filepath: str,
    preferred_sheets=None,
    cancel_check=None,
    return_meta: bool = False,
    target_cols=None,
):
    """
    XLSM/XLSXファイルから指定シート（無ければ先頭シート）の行を辞書リストとして高速に読み込む。
    各行は {'A': val, 'B': val, ...} の辞書。
    読み取り専用ハンドル("rb")によるZIP部分読取でVBAバイナリや他シートをスキップし、
    元ファイルを一切変更・ロックせずに省メモリかつ高速に読み込む。
    """
    if cancel_check and cancel_check():
        raise InterruptedError("処理がキャンセルされました")
    if not filepath or not os.path.exists(filepath):
        raise FileNotFoundError(f"ファイルが見つかりません: {filepath}")
    if os.path.isdir(filepath):
        raise IsADirectoryError(f"ファイルではなくフォルダが指定されています: {filepath}")

    try:
        file_size = os.path.getsize(filepath)
    except OSError:
        file_size = -1
    if file_size == 0:
        raise FormatValidationError(
            "Excelファイル",
            filepath,
            "ファイルサイズが0バイト（空ファイル）です。",
            "有効なExcelファイル (.xlsm / .xlsx)",
            "空ファイル (0 bytes)",
        )

    try:
        with open_shared_readonly(filepath, cancel_check=cancel_check) as f_ro:
            try:
                with zipfile.ZipFile(f_ro, "r") as z:
                    rows, matched_sheet_name, all_sheet_names = _read_excel_sheet_rows_from_zip(
                        filepath,
                        z,
                        preferred_sheets=preferred_sheets,
                        cancel_check=cancel_check,
                        target_cols=target_cols,
                    )
                    if return_meta:
                        return rows, matched_sheet_name, all_sheet_names
                    return rows
            except zipfile.BadZipFile:
                try:
                    f_ro.seek(0)
                    head = f_ro.read(8)
                except Exception:
                    head = b""
                # 他ユーザーがちょうど保存途中でZIP構造が一時不整合になった場合はスナップショット読込で再試行
                if head.startswith(b"PK\x03\x04"):
                    raise RuntimeError("concurrent_zip_retry")
                if head.startswith(b"\xd0\xcf\x11\xe0"):
                    det = "旧Excel形式 (.xls) またはパスワード保護されたOfficeファイル"
                else:
                    det = "Excel (.xlsx/.xlsm) ではないファイル形式（CSVまたはテキスト等）"
                raise FormatValidationError(
                    "Excelファイル",
                    filepath,
                    "有効な .xlsx / .xlsm 形式として読み込めません。",
                    "Excelブック (.xlsm / .xlsx)",
                    det,
                )
    except (FormatValidationError, InterruptedError, FileNotFoundError, IsADirectoryError):
        raise
    except Exception:
        # 特殊ファイルシステムや他ユーザー保存タイミングと重なった場合は待機付き共有スナップショット読込へフォールバック
        _sleep_cancellable(0.15, cancel_check=cancel_check)
        raw_bytes = open_readonly_bytes(filepath, cancel_check=cancel_check)
        if not raw_bytes:
            raise FormatValidationError(
                "Excelファイル",
                filepath,
                "ファイルサイズが0バイト（空ファイル）です。",
                "有効なExcelファイル (.xlsm / .xlsx)",
                "空ファイル (0 bytes)",
            )
        try:
            with zipfile.ZipFile(io.BytesIO(raw_bytes), "r") as z:
                rows, matched_sheet_name, all_sheet_names = _read_excel_sheet_rows_from_zip(
                    filepath,
                    z,
                    preferred_sheets=preferred_sheets,
                    cancel_check=cancel_check,
                    target_cols=target_cols,
                )
                if return_meta:
                    return rows, matched_sheet_name, all_sheet_names
                return rows
        except zipfile.BadZipFile:
            if raw_bytes.startswith(b"\xd0\xcf\x11\xe0"):
                det = "旧Excel形式 (.xls) またはパスワード保護されたOfficeファイル"
            else:
                det = "Excel (.xlsx/.xlsm) ではないファイル形式（CSVまたはテキスト等）"
            raise FormatValidationError(
                "Excelファイル",
                filepath,
                "有効な .xlsx / .xlsm 形式として読み込めません。",
                "Excelブック (.xlsm / .xlsx)",
                det,
            )


def validate_table_excel_rows(filepath: str, category_key: str, rows: list, matched_sheet: str = "", all_sheets: list = None):
    """
    購入品・外注加工・受注表の各Excelデータがサンプルファイルのフォーマットと一致するか検証する。
    一致しない場合は詳細な FormatValidationError を送出する。
    """
    all_sheets = all_sheets or []
    cat_names = {
        "kounyu": "購入品ファイル",
        "gaichu": "外注加工ファイル",
        "juchu": "受注表ファイル",
    }
    cat_label = cat_names.get(category_key, "Excelファイル")

    if not rows:
        raise FormatValidationError(
            cat_label,
            filepath,
            "シート内にデータ行（見出し行）が存在しません。",
            "サンプルファイルと同じ見出し行を持つExcelシート",
            f"対象シート: '{matched_sheet or 'なし'}' (全シート: {', '.join(all_sheets) or 'なし'}) / データ0行",
        )

    header = rows[0]
    if category_key in ("kounyu", "gaichu"):
        sample_name = "購入品.xlsm" if category_key == "kounyu" else "外注加工.xlsm"
        col_a = normalize_kanri_no(header.get("A", "")).replace(".", "")
        col_q = str(header.get("Q", "")).strip()
        col_m = str(header.get("M", "")).strip()
        col_n = str(header.get("N", "")).strip()
        has_valid_a = (col_a == "管理No")
        has_valid_q = ("金額" in col_q or col_m == "単価" or col_n == "必要数")
        if not (has_valid_a and has_valid_q):
            det_cols = ", ".join(f"{k}列='{header.get(k, '')}'" for k in ("A", "B", "H", "M", "N", "Q") if k in header)
            if not det_cols:
                det_cols = ", ".join(f"{k}列='{v}'" for k, v in list(header.items())[:6]) or "(見出しなし)"
            raise FormatValidationError(
                cat_label,
                filepath,
                f"サンプル（{sample_name}）と列フォーマットが異なります。A列「管理No」およびQ列「金額 単価X必要数」（M列:単価、N列:必要数）が必要です。",
                f"シート「テーブル」 / 1行目見出し: A列=管理No, B列=出荷先, H列=品名, M列=単価, N列=必要数, Q列=金額 単価X必要数",
                f"シート='{matched_sheet}' / 1行目: {det_cols}",
            )
    elif category_key == "juchu":
        col_b = normalize_kanri_no(header.get("B", "")).replace(".", "")
        col_a = str(header.get("A", "")).strip()
        col_c = str(header.get("C", "")).strip()
        col_e = str(header.get("E", "")).strip()
        has_valid_b = (col_b == "管理No")
        has_valid_meta = ("受注日" in col_a or "受注先" in col_c or "出荷先" in col_e)
        if not (has_valid_b and has_valid_meta):
            det_cols = ", ".join(f"{k}列='{header.get(k, '')}'" for k in ("A", "B", "C", "E", "G", "H", "I") if k in header)
            if not det_cols:
                det_cols = ", ".join(f"{k}列='{v}'" for k, v in list(header.items())[:6]) or "(見出しなし)"
            raise FormatValidationError(
                cat_label,
                filepath,
                "サンプル（受注票頚城.xlsm）と列フォーマットが異なります。A列「受注日」、B列「管理No」、C列「受注先」、E列「出荷先」が必要です。",
                "シート「受注表」 / 1行目見出し: A列=受注日, B列=管理No, C列=受注先, E列=出荷先, G列=プロジェクト名, H列=品名、ライン名等, I列=製品詳細",
                f"シート='{matched_sheet}' / 1行目: {det_cols}",
            )


def col_index_to_letter(idx: int) -> str:
    """0始まりの列インデックスからExcel列名（0->'A', 26->'AA'）を返す"""
    res = ""
    n = idx + 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        res = chr(65 + rem) + res
    return res


def scan_seizo_excel_entries(paths: list, cancel_check=None) -> list:
    """
    制作伝票用のフォルダ・ファイル指定リストから、指定フォルダ直下にある
    .xls / .xlsx / .xlsm ファイルのメタ情報リスト
    [(norm_path, size_bytes, mtime_ns, mtime_sec), ...] を 1 パスで高速に取得する。
    DirEntry.stat() を利用するため、ネットワーク／クラウドドライブ上でも余分な stat 呼び出しを省ける。
    一時ロックファイル（~$...）は除外し、重複を排除する。
    """
    result = []
    seen = set()
    exts = (".xls", ".xlsx", ".xlsm")

    def add_entry(fpath: str, st):
        norm = os.path.normpath(fpath)
        key = os.path.normcase(os.path.abspath(norm))
        if key not in seen:
            seen.add(key)
            size_b = int(getattr(st, "st_size", 0))
            mtime_ns = int(getattr(st, "st_mtime_ns", int(st.st_mtime * 1_000_000_000)))
            result.append((norm, size_b, mtime_ns, float(st.st_mtime)))

    for p in (paths or []):
        if cancel_check and cancel_check():
            raise InterruptedError("処理がキャンセルされました")
        if not p:
            continue
        if os.path.isdir(p):
            try:
                raw_entries = []
                with os.scandir(p) as it:
                    for idx_e, entry in enumerate(it):
                        if cancel_check and (idx_e & 255) == 0 and cancel_check():
                            raise InterruptedError("処理がキャンセルされました")
                        name = entry.name
                        if name.startswith("~$") or not name.lower().endswith(exts):
                            continue
                        try:
                            if entry.is_file():
                                raw_entries.append((name.lower(), entry.path, entry.stat()))
                        except Exception:
                            pass
                raw_entries.sort(key=lambda t: t[0])
                for _, e_path, e_st in raw_entries:
                    add_entry(e_path, e_st)
            except InterruptedError:
                raise
            except Exception:
                pass
        elif os.path.isfile(p):
            name = os.path.basename(p)
            if not name.startswith("~$") and name.lower().endswith(exts):
                try:
                    add_entry(p, os.stat(p))
                except Exception:
                    pass
    return result


def expand_seizo_excel_files(paths: list) -> list:
    """
    制作伝票用のフォルダ・ファイル指定リストから、指定フォルダ直下にある
    .xls / .xlsx / .xlsm ファイル一覧を展開する（サブフォルダは含まない）。
    一時ロックファイル（~$...）は除外し、重複を排除する。
    """
    return [e[0] for e in scan_seizo_excel_entries(paths)]


def expand_seizo_protected_dirs(paths: list) -> set:
    """制作伝票フォルダの正規化絶対パス集合（書き込み禁止対象）を返す"""
    protected_dirs = set()
    for p in (paths or []):
        if not p:
            continue
        if os.path.isdir(p):
            protected_dirs.add(os.path.normcase(os.path.abspath(os.path.normpath(p))))
    return protected_dirs


def summarize_seizo_folder_mtime(folder_paths: list, entries: list = None) -> str:
    """制作伝票フォルダ直下（サブフォルダ除く・複数フォルダ対応）の対象Excelファイル数と最新更新日時を返す"""
    if not folder_paths:
        return "未選択"
    missing_folders = [p for p in folder_paths if p and not os.path.exists(p)]
    scanned = entries if entries is not None else scan_seizo_excel_entries(folder_paths)
    folder_cnt = len([p for p in folder_paths if p])
    if not scanned:
        if missing_folders:
            return "フォルダが見つかりません"
        return f"フォルダ直下に対象Excelなし ({folder_cnt}フォルダ)" if folder_cnt > 1 else "フォルダ直下に対象Excelなし"
    latest_ts = max(e[3] for e in scanned)
    latest_str = datetime.datetime.fromtimestamp(latest_ts).strftime("%Y/%m/%d %H:%M:%S")
    if folder_cnt > 1:
        return f"最新更新: {latest_str} ({folder_cnt}フォルダ / Excel {len(scanned):,}ファイル)"
    return f"最新更新: {latest_str} (Excel {len(scanned):,}ファイル)"


class MainFileParseCache:
    """
    日報・購入品・外注加工・受注表のインメモリ解析キャッシュ（スレッドセーフ・上限付きLRU）。
    (種別, 正規化絶対パス, ファイルサイズ, mtime_ns) が完全一致する未変更ファイルを再利用し、
    6万行超の購入品ファイル等を含む2回目以降の「集計・一覧更新」を劇的に高速化する。
    """

    def __init__(self, max_entries: int = 32):
        self.max_entries = max(4, int(max_entries))
        self._lock = threading.Lock()
        self._store = OrderedDict()

    @staticmethod
    def _make_key(category: str, filepath: str, size_b: int, mtime_ns: int) -> tuple:
        norm = os.path.normcase(os.path.abspath(os.path.normpath(filepath)))
        return (category, norm, int(size_b), int(mtime_ns))

    def get(self, category: str, filepath: str, size_b: int, mtime_ns: int):
        key = self._make_key(category, filepath, size_b, mtime_ns)
        with self._lock:
            val = self._store.get(key)
            if val is not None:
                self._store.move_to_end(key)
            return val

    def put(self, category: str, filepath: str, size_b: int, mtime_ns: int, payload):
        key = self._make_key(category, filepath, size_b, mtime_ns)
        with self._lock:
            self._store[key] = payload
            self._store.move_to_end(key)
            while len(self._store) > self.max_entries:
                self._store.popitem(last=False)


class SeizoCacheStore:
    """
    制作伝票パース結果のインメモリ＋JSONファイル永続キャッシュ。
    (正規化絶対パス, ファイルサイズ, mtime_ns, キャッシュバージョン) が完全一致する
    未変更ファイルのみを再利用することで、数千ファイルの制作伝票でも2回目以降を1秒未満に高速化する。
    """

    def __init__(self, cache_path: str):
        self.cache_path = cache_path
        self._lock = threading.Lock()
        self._entries = {}  # norm_key -> {"size": int, "mtime_ns": int, "rec": dict}
        self._loaded = False
        self._dirty = False

    @staticmethod
    def _norm_key(filepath: str) -> str:
        return os.path.normcase(os.path.abspath(os.path.normpath(filepath)))

    @staticmethod
    def _restore_record(rec_raw: dict, filepath: str) -> dict:
        rec = dict(rec_raw)
        specs = rec.get("_specs")
        if isinstance(specs, dict):
            rec["_specs"] = {int(k): v for k, v in specs.items()}
        units = rec.get("_units")
        if isinstance(units, dict):
            rec["_units"] = {int(k): tuple(v) if isinstance(v, list) else v for k, v in units.items()}
        rec["_src_path"] = filepath
        return rec

    @staticmethod
    def _serialize_record(rec: dict) -> dict:
        out = dict(rec)
        specs = out.get("_specs")
        if isinstance(specs, dict):
            out["_specs"] = {str(k): v for k, v in specs.items()}
        units = out.get("_units")
        if isinstance(units, dict):
            out["_units"] = {str(k): list(v) if isinstance(v, tuple) else v for k, v in units.items()}
        return out

    def ensure_loaded(self):
        with self._lock:
            if self._loaded:
                return
            self._loaded = True
            if not self.cache_path or not os.path.exists(self.cache_path):
                return
            try:
                with open(self.cache_path, "r", encoding="utf-8") as f:
                    payload = json.load(f)
                if isinstance(payload, dict) and payload.get("version") == SEIZO_CACHE_VERSION:
                    items = payload.get("entries")
                    if isinstance(items, dict):
                        self._entries = items
            except Exception:
                self._entries = {}

    def get(self, filepath: str, size_b: int, mtime_ns: int):
        self.ensure_loaded()
        key = self._norm_key(filepath)
        with self._lock:
            item = self._entries.get(key)
            if not item:
                return None
            if item.get("size") == size_b and item.get("mtime_ns") == mtime_ns and isinstance(item.get("rec"), dict):
                return self._restore_record(item["rec"], filepath)
            return None

    def put(self, filepath: str, size_b: int, mtime_ns: int, rec: dict):
        self.ensure_loaded()
        key = self._norm_key(filepath)
        ser = self._serialize_record(rec)
        with self._lock:
            self._entries[key] = {
                "size": int(size_b),
                "mtime_ns": int(mtime_ns),
                "rec": ser,
            }
            self._dirty = True

    def save_if_dirty(self, protected_paths: set = None, protected_dirs: set = None):
        with self._lock:
            if not self._dirty or not self.cache_path:
                return
            if is_same_or_protected_path(self.cache_path, protected_paths or set(), protected_dirs or set()):
                return
            snapshot = {
                "version": SEIZO_CACHE_VERSION,
                "entries": dict(self._entries),
            }
            self._dirty = False

        tmp_path = None
        try:
            parent = os.path.dirname(os.path.abspath(self.cache_path))
            if parent and not os.path.exists(parent):
                os.makedirs(parent, exist_ok=True)
            tmp_path = f"{self.cache_path}.tmp.{os.getpid()}.{threading.get_ident()}"
            with open(tmp_path, "w", encoding="utf-8") as f:
                json.dump(snapshot, f, ensure_ascii=False, separators=(",", ":"))
            try:
                os.replace(tmp_path, self.cache_path)
            except OSError:
                time.sleep(0.06)
                os.replace(tmp_path, self.cache_path)
            tmp_path = None
        except Exception:
            if tmp_path and os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass


def _load_shared_strings_subset_from_zip(z: zipfile.ZipFile, namelist_set: set, needed_indices: set, cancel_check=None) -> dict:
    """
    ZIP内の xl/sharedStrings.xml から、needed_indices に含まれる文字列インデックスだけを
    ストリーム解析で遅延抽出し、最大インデックス到達時点で即座に打ち切る超高速ローダー。
    """
    if not needed_indices or "xl/sharedStrings.xml" not in namelist_set:
        return {}
    max_needed = max(needed_indices)
    result = {}
    si_tag = f"{NS_MAIN}si"
    t_tag = f"{NS_MAIN}t"
    r_tag = f"{NS_MAIN}r"
    idx = 0

    with z.open("xl/sharedStrings.xml") as sf:
        for event, elem in ET.iterparse(sf, events=("end",)):
            if elem.tag == si_tag:
                if cancel_check and (idx & 1023) == 0 and cancel_check():
                    raise InterruptedError("処理がキャンセルされました")
                if idx in needed_indices:
                    t_direct = elem.find(t_tag)
                    if t_direct is not None:
                        result[idx] = t_direct.text or ""
                    else:
                        parts = []
                        for r_el in elem.findall(r_tag):
                            t_el = r_el.find(t_tag)
                            if t_el is not None and t_el.text:
                                parts.append(t_el.text)
                        result[idx] = "".join(parts)
                    if len(result) >= len(needed_indices):
                        elem.clear()
                        break
                elem.clear()
                idx += 1
                if idx > max_needed:
                    break
    return result


def _read_seizo_cells_xml_from_zip(filepath: str, z: zipfile.ZipFile, cancel_check=None) -> tuple:
    """
    開かれた ZipFile から「製作伝票」シートの対象行（SEIZO_TARGET_ROWS）のみを解析し、
    参照された sharedStrings インデックスだけを遅延ロードして {(row_num, col): val} を返す。
    """
    namelist_set = set(z.namelist())
    if "xl/workbook.xml" not in namelist_set or "xl/_rels/workbook.xml.rels" not in namelist_set:
        raise FormatValidationError(
            "制作伝票ファイル",
            filepath,
            "Excelブック構造（xl/workbook.xml）が見つかりません。",
            "制作伝票Excel (.xls / .xlsx / .xlsm)",
            "workbook.xml 不在",
        )

    wb_root = ET.fromstring(z.read("xl/workbook.xml"))
    rels_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rels = {r.attrib["Id"]: r.attrib["Target"] for r in rels_root if "Id" in r.attrib and "Target" in r.attrib}

    sheet_target = None
    matched_sheet_name = ""
    all_sheet_names = []
    preferred_names = ("製作伝票", "制作伝票")
    for sheet in wb_root.findall(f".//{NS_MAIN}sheet"):
        name = sheet.attrib.get("name", "")
        all_sheet_names.append(name)
        rid = sheet.attrib.get(f"{NS_REL}id", "")
        raw_target = rels.get(rid, "").lstrip("/")
        tgt = raw_target if raw_target.startswith("xl/") else f"xl/{raw_target}"
        if tgt not in namelist_set and raw_target in namelist_set:
            tgt = raw_target
        if name in preferred_names and tgt in namelist_set:
            sheet_target = tgt
            matched_sheet_name = name
            break

    if sheet_target is None:
        for sheet in wb_root.findall(f".//{NS_MAIN}sheet"):
            name = sheet.attrib.get("name", "")
            if "製作伝票" in name or "制作伝票" in name:
                rid = sheet.attrib.get(f"{NS_REL}id", "")
                raw_target = rels.get(rid, "").lstrip("/")
                tgt = raw_target if raw_target.startswith("xl/") else f"xl/{raw_target}"
                if tgt not in namelist_set and raw_target in namelist_set:
                    tgt = raw_target
                if tgt in namelist_set:
                    sheet_target = tgt
                    matched_sheet_name = name
                    break

    if not sheet_target:
        raise FormatValidationError(
            "制作伝票ファイル",
            filepath,
            "ブック内に「製作伝票」（または「制作伝票」）シートが存在しません。",
            "シート「製作伝票」（5行目H列=管理No、21行=調整用、28～35行=材料費・材料費計、36行=設計費、38行=小計１）",
            f"検出されたシート一覧: {', '.join(all_sheet_names) or '(なし)'}",
        )

    c_tag = f"{NS_MAIN}c"
    v_tag = f"{NS_MAIN}v"
    is_tag = f"{NS_MAIN}is"
    t_tag = f"{NS_MAIN}t"
    row_tag = f"{NS_MAIN}row"

    cells = {}
    pending_ss = []  # list of ((r_num, col), ss_idx)
    needed_ss_indices = set()

    with z.open(sheet_target) as wf:
        for event, elem in ET.iterparse(wf, events=("end",)):
            if elem.tag == row_tag:
                if cancel_check and cancel_check():
                    raise InterruptedError("処理がキャンセルされました")
                r_attr = elem.attrib.get("r", "")
                try:
                    r_num = int(r_attr)
                except ValueError:
                    r_num = 0
                if r_num > 38:
                    elem.clear()
                    break
                if r_num in SEIZO_TARGET_ROWS:
                    for c_el in elem.findall(c_tag):
                        ref = c_el.attrib.get("r", "")
                        col = ref.rstrip("0123456789")
                        if not col:
                            continue
                        t_attr = c_el.attrib.get("t", "")
                        if t_attr == "e":
                            continue
                        if t_attr == "s":
                            v_el = c_el.find(v_tag)
                            if v_el is not None and v_el.text:
                                try:
                                    ss_idx = int(v_el.text)
                                    needed_ss_indices.add(ss_idx)
                                    pending_ss.append(((r_num, col), ss_idx))
                                except ValueError:
                                    pass
                        elif t_attr == "inlineStr":
                            is_el = c_el.find(is_tag)
                            if is_el is not None:
                                s_val = "".join(x.text or "" for x in is_el.iter(t_tag)).strip()
                                if s_val != "":
                                    cells[(r_num, col)] = s_val
                        else:
                            v_el = c_el.find(v_tag)
                            if v_el is not None and v_el.text is not None:
                                s_val = str(v_el.text).strip()
                                if s_val != "":
                                    cells[(r_num, col)] = s_val
                elem.clear()

    if needed_ss_indices:
        ss_map = _load_shared_strings_subset_from_zip(z, namelist_set, needed_ss_indices, cancel_check=cancel_check)
        for coord, ss_idx in pending_ss:
            txt = ss_map.get(ss_idx, "")
            if txt:
                s_val = str(txt).strip()
                if s_val != "":
                    cells[coord] = s_val

    return cells, matched_sheet_name, all_sheet_names


def _read_seizo_cells_xml(filepath: str, raw_bytes: bytes = None, cancel_check=None) -> tuple:
    """
    XLSX/XLSMから「製作伝票」シートの対象セルのみを読み取る。
    raw_bytes 未指定時は読み取り専用ハンドル("rb")のままZIP部分読取を行い、
    マクロバイナリや他シートの不要なI/Oをスキップして大幅に高速化する。
    """
    if raw_bytes is None:
        try:
            with open_shared_readonly(filepath, cancel_check=cancel_check) as f_ro:
                with zipfile.ZipFile(f_ro, "r") as z:
                    return _read_seizo_cells_xml_from_zip(filepath, z, cancel_check=cancel_check)
        except (FormatValidationError, InterruptedError, FileNotFoundError, IsADirectoryError):
            raise
        except zipfile.BadZipFile:
            # 他ユーザー保存途中の可能性も考慮し、短時間待機後にスナップショット読込で再検証する
            _sleep_cancellable(0.15, cancel_check=cancel_check)
            raw_bytes = open_readonly_bytes(filepath, cancel_check=cancel_check)
        except Exception:
            _sleep_cancellable(0.15, cancel_check=cancel_check)
            raw_bytes = open_readonly_bytes(filepath, cancel_check=cancel_check)

    if not raw_bytes:
        raise FormatValidationError(
            "制作伝票ファイル",
            filepath,
            "ファイルサイズが0バイト（空ファイル）です。",
            "シート「製作伝票」を含む制作伝票Excelファイル",
            "空ファイル (0 bytes)",
        )
    try:
        with zipfile.ZipFile(io.BytesIO(raw_bytes), "r") as z:
            return _read_seizo_cells_xml_from_zip(filepath, z, cancel_check=cancel_check)
    except zipfile.BadZipFile:
        raise FormatValidationError(
            "制作伝票ファイル",
            filepath,
            "有効なExcel (.xlsx / .xlsm) ファイルとして開けません（ファイル破損または非対応形式）。",
            "制作伝票Excel (.xls / .xlsx / .xlsm)",
            "ZIP構造不正",
        )


def _has_xlrd_available() -> bool:
    try:
        importlib.import_module("xlrd")
        return True
    except ImportError:
        return False


def _read_seizo_cells_xls(filepath: str, raw_bytes: bytes) -> tuple:
    """旧形式 .xls のメモリバイト列から「製作伝票」シートの1～40行のセル {(row_num, col): val} とシート情報を読み込む"""
    cells = {}
    try:
        xlrd = importlib.import_module("xlrd")
        try:
            wb = xlrd.open_workbook(filename=filepath, file_contents=raw_bytes)
        except TypeError:
            wb = xlrd.open_workbook(filepath)
        all_sheet_names = list(wb.sheet_names())
        sheet = None
        matched_sheet_name = ""
        for s_name in all_sheet_names:
            if s_name in ("製作伝票", "制作伝票") or "製作伝票" in s_name or "制作伝票" in s_name or "伝票" in s_name:
                sheet = wb.sheet_by_name(s_name)
                matched_sheet_name = s_name
                break
        if sheet is None:
            raise FormatValidationError(
                "制作伝票ファイル",
                filepath,
                "ブック内に「製作伝票」（または「制作伝票」）シートが存在しません。",
                "シート「製作伝票」（5行目H列=管理No、21行=調整用、28～35行=材料費・材料費計、36行=設計費、38行=小計１）",
                f"検出されたシート一覧: {', '.join(all_sheet_names) or '(なし)'}",
            )
        max_r = min(40, sheet.nrows)
        max_c = min(36, sheet.ncols)
        for r_i in range(max_r):
            if (r_i + 1) not in SEIZO_TARGET_ROWS:
                continue
            for c_i in range(max_c):
                val = sheet.cell_value(r_i, c_i)
                if val is not None and str(val).strip() != "":
                    ctype = sheet.cell_type(r_i, c_i)
                    if ctype == 3:  # XL_CELL_DATE
                        dt_tup = xlrd.xldate_as_tuple(val, wb.datemode)
                        cells[(r_i + 1, col_index_to_letter(c_i))] = f"{dt_tup[0]:04d}/{dt_tup[1]:02d}/{dt_tup[2]:02d}"
                    else:
                        cells[(r_i + 1, col_index_to_letter(c_i))] = str(val).strip()
        return cells, matched_sheet_name, all_sheet_names
    except ImportError:
        pass
    except FormatValidationError:
        raise

    if sys.platform == "win32":
        tmp_dir = tempfile.mkdtemp(prefix="denpyohub_xls_")
        tmp_path = os.path.join(tmp_dir, os.path.basename(filepath))
        with open(tmp_path, "wb") as f_tmp:
            f_tmp.write(raw_bytes)
        try:
            os.chmod(tmp_path, stat.S_IREAD)
        except Exception:
            pass
        pythoncom = importlib.import_module("pythoncom")
        win32_client = importlib.import_module("win32com.client")
        pythoncom.CoInitialize()
        excel = None
        wb = None
        created_excel = False
        try:
            excel = win32_client.DispatchEx("Excel.Application")
            created_excel = True
            excel.Visible = False
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Open(tmp_path, UpdateLinks=0, ReadOnly=True)
            ws = None
            matched_sheet_name = ""
            all_sheet_names = []
            for sh in wb.Worksheets:
                all_sheet_names.append(sh.Name)
                if sh.Name in ("製作伝票", "制作伝票") or "伝票" in sh.Name:
                    ws = sh
                    matched_sheet_name = sh.Name
                    break
            if ws is None:
                raise FormatValidationError(
                    "制作伝票ファイル",
                    filepath,
                    "ブック内に「製作伝票」（または「制作伝票」）シートが存在しません。",
                    "シート「製作伝票」（5行目H列=管理No、21行=調整用、28～35行=材料費・材料費計、36行=設計費、38行=小計１）",
                    f"検出されたシート一覧: {', '.join(all_sheet_names) or '(なし)'}",
                )
            data = ws.Range("A1:AJ40").Value
            if data:
                for r_i, row_tup in enumerate(data, 1):
                    if r_i not in SEIZO_TARGET_ROWS:
                        continue
                    for c_i, val in enumerate(row_tup):
                        if val is not None and str(val).strip() != "":
                            cells[(r_i, col_index_to_letter(c_i))] = str(val).strip()
            return cells, matched_sheet_name, all_sheet_names
        finally:
            if wb is not None:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass
            if created_excel and excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()
            shutil.rmtree(tmp_dir, ignore_errors=True)

    raise RuntimeError(".xls 形式を読み込むための xlrd が利用できません。")


def _read_seizo_xls_batch_com(file_entries: list, cancel_check=None, base_folders: list = None, on_item_done=None) -> list:
    """
    xlrd が使えない環境で複数の .xls ファイルを読み込む際、
    Excel COM インスタンスを1回だけ起動して全ファイルを ReadOnly で連続処理し、劇的に高速化する。
    戻り値: [(entry_tuple, rec_or_none, err_or_none), ...]
    """
    if not file_entries or sys.platform != "win32":
        return []
    pythoncom = importlib.import_module("pythoncom")
    win32_client = importlib.import_module("win32com.client")
    tmp_dir = tempfile.mkdtemp(prefix="denpyohub_xls_batch_")
    results = []
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32_client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            excel.ScreenUpdating = False
            excel.EnableEvents = False
        except Exception:
            pass
        for idx_i, entry in enumerate(file_entries):
            if cancel_check and cancel_check():
                raise InterruptedError("処理がキャンセルされました")
            fp, _sz, _mns, m_sec = entry
            wb = None
            tmp_path = None
            try:
                raw_bytes = open_readonly_bytes(fp, cancel_check=cancel_check)
                if not raw_bytes:
                    raise FormatValidationError(
                        "制作伝票ファイル",
                        fp,
                        "ファイルサイズが0バイト（空ファイル）です。",
                        "シート「製作伝票」を含む制作伝票Excelファイル",
                        "空ファイル (0 bytes)",
                    )
                tmp_path = os.path.join(tmp_dir, f"tmp_{idx_i}_{os.path.basename(fp)}")
                with open(tmp_path, "wb") as f_tmp:
                    f_tmp.write(raw_bytes)
                try:
                    os.chmod(tmp_path, stat.S_IREAD)
                except Exception:
                    pass
                wb = excel.Workbooks.Open(tmp_path, UpdateLinks=0, ReadOnly=True)
                ws = None
                matched_sheet = ""
                all_sheets = []
                for sh in wb.Worksheets:
                    all_sheets.append(sh.Name)
                    if sh.Name in ("製作伝票", "制作伝票") or "伝票" in sh.Name:
                        ws = sh
                        matched_sheet = sh.Name
                        break
                if ws is None:
                    raise FormatValidationError(
                        "制作伝票ファイル",
                        fp,
                        "ブック内に「製作伝票」（または「制作伝票」）シートが存在しません。",
                        "シート「製作伝票」（5行目H列=管理No、21行=調整用、28～35行=材料費・材料費計、36行=設計費、38行=小計１）",
                        f"検出されたシート一覧: {', '.join(all_sheets) or '(なし)'}",
                    )
                cells = {}
                data = ws.Range("A1:AJ40").Value
                if data:
                    for r_i, row_tup in enumerate(data, 1):
                        if r_i not in SEIZO_TARGET_ROWS:
                            continue
                        for c_i, val in enumerate(row_tup):
                            if val is not None and str(val).strip() != "":
                                cells[(r_i, col_index_to_letter(c_i))] = str(val).strip()
                rec = _build_seizo_record_from_cells(fp, cells, matched_sheet, base_folders=base_folders, mtime_sec=m_sec)
                results.append((entry, rec, None))
            except InterruptedError:
                raise
            except Exception as exc:
                results.append((entry, None, exc))
            finally:
                if wb is not None:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
                if tmp_path and os.path.exists(tmp_path):
                    try:
                        os.chmod(tmp_path, stat.S_IWRITE | stat.S_IREAD)
                        os.remove(tmp_path)
                    except Exception:
                        pass
                if on_item_done:
                    on_item_done(fp)
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
        shutil.rmtree(tmp_dir, ignore_errors=True)
    return results


def _build_seizo_record_from_cells(
    filepath: str,
    cells: dict,
    matched_sheet: str,
    base_folders: list = None,
    mtime_sec: float = None,
) -> dict:
    """抽出したセル辞書をフォーマット検証し、制作伝票レコード辞書を構築する共通コア"""
    def cell(r, c):
        return cells.get((r, c), "")

    kno = normalize_kanri_no(cell(5, "H"))
    if not kno or kno in ("管理No", "管理No."):
        raise FormatValidationError(
            "制作伝票ファイル",
            filepath,
            f"シート「{matched_sheet}」の5行目H列（H5セル）に有効な管理Noが記載されていません。",
            "シート「製作伝票」の H5 セルに管理No（例: 22601013）、21行=調整用、28～35行=材料費・材料費計、36行=設計費、38行=小計１",
            f"H5セル='{cell(5, 'H')}', G5セル='{cell(5, 'G')}'",
        )

    # 制作伝票シートの構造シグネチャ検証（5行目ラベル、または21行・28～38行のいずれかに項目・値が存在すること）
    has_seizo_structure = (
        "管理No" in cell(5, "G")
        or "伝票" in cell(2, "N")
        or any(
            cell(r_no, col_c) != ""
            for r_no in (21, 28, 29, 30, 31, 32, 33, 34, 35, 36, 38)
            for col_c in ("C", "I", "M", "O", "Q")
        )
    )
    if not has_seizo_structure:
        raise FormatValidationError(
            "制作伝票ファイル",
            filepath,
            f"シート「{matched_sheet}」の21行（調整用）・28～35行（材料費）・36行（設計費）・38行（小計１）の伝票構造が確認できません。",
            "サンプル（22601013.xlsm / 22601003.xlsm）と同じ行配置の制作伝票シート",
            f"シート='{matched_sheet}', H5='{kno}', 21～38行のデータなし",
        )

    def row_yen_or_calc(row_num: int, val_col: str, unit_col: str = "I", qty_col: str = "M") -> float:
        raw = cell(row_num, val_col)
        if raw != "":
            return to_float(raw, 0.0)
        u = cell(row_num, unit_col)
        q = cell(row_num, qty_col)
        if u != "" and q != "":
            return to_float(u, 0.0) * to_float(q, 0.0)
        return 0.0

    v_21 = row_yen_or_calc(21, "O", "A", "I")
    v_28 = row_yen_or_calc(28, "Q", "I", "M")
    v_29 = row_yen_or_calc(29, "Q", "I", "M")
    v_30 = row_yen_or_calc(30, "Q", "I", "M")
    v_31 = row_yen_or_calc(31, "Q", "I", "M")
    v_32 = row_yen_or_calc(32, "Q", "I", "M")
    v_33 = row_yen_or_calc(33, "Q", "I", "M")
    v_34 = row_yen_or_calc(34, "Q", "I", "M")

    raw_35 = cell(35, "Q")
    if raw_35 != "":
        v_35 = to_float(raw_35, 0.0)
    else:
        factor = to_float(cell(35, "B"), 1.0) if cell(35, "B") != "" else 1.0
        v_35 = (v_28 + v_29 + v_30 + v_31 + v_32 + v_33 + v_34) * factor

    raw_36 = cell(36, "O") if cell(36, "O") != "" else cell(36, "Q")
    v_36 = to_float(raw_36, 0.0)
    sheet_sub1 = to_float(cell(38, "I"), 0.0)
    v_aa17 = to_float(cell(17, "AA"), 0.0)
    v_aa18 = to_float(cell(18, "AA"), 0.0)
    v_aa19 = to_float(cell(19, "AA"), 0.0)
    v_aa30 = to_float(cell(30, "AA"), 0.0)
    v_ad9 = to_float(cell(9, "AD"), 0.0)

    src_display = os.path.basename(filepath)
    if base_folders:
        norm_fp = os.path.abspath(os.path.normpath(filepath))
        for bf in base_folders:
            if bf and os.path.isdir(bf):
                try:
                    norm_bf = os.path.abspath(os.path.normpath(bf))
                    if os.path.commonpath([norm_fp, norm_bf]) == norm_bf:
                        rel = os.path.relpath(norm_fp, norm_bf)
                        if rel and not rel.startswith(".."):
                            src_display = rel
                            break
                except Exception:
                    pass

    if mtime_sec is not None:
        mtime_str = datetime.datetime.fromtimestamp(mtime_sec).strftime("%Y/%m/%d %H:%M:%S")
    else:
        mtime_str = format_file_mtime(filepath)

    return {
        "管理No": kno,
        "調整用": v_21,
        "材料費１(型鋼)": v_28,
        "材料費２(鉄板)": v_29,
        "材料費３(コラム)": v_30,
        "その他材料費(31行)": v_31,
        "その他材料費(32行)": v_32,
        "その他材料費(33行)": v_33,
        "切断材料": v_34,
        "材料費計": v_35,
        "設計費": v_36,
        "重機、運搬費": v_aa17,
        "車両、交通費": v_aa18,
        "出張費": v_aa19,
        "総請求金額": v_aa30,
        "_制作伝票数量": v_ad9,
        "_シート小計1": sheet_sub1,
        "_specs": {
            31: cell(31, "B") or "ヘルテン",
            32: cell(32, "B") or "ｴﾊﾞｰﾊｰﾄﾞ",
            33: cell(33, "B") or "SUS",
        },
        "_meta": {
            "受注日": format_excel_date(cell(12, "G")),
            "受注先": cell(11, "G"),
            "出荷先": cell(9, "F"),
            "プロジェクト名": cell(6, "G"),
            "品名、ライン名等": cell(9, "I"),
            "製品詳細": cell(9, "Q"),
            "数量": cell(9, "AD"),
            "単位": cell(9, "AF"),
            "売り上げ区分": cell(9, "C"),
            "出荷日": format_excel_date(cell(12, "O")),
            "請求完了日": format_excel_date(cell(12, "W")),
            "備考": cell(14, "G") or cell(14, "V"),
        },
        "_units": {
            21: (cell(21, "A"), cell(21, "I")),
            28: (cell(28, "I"), cell(28, "M")),
            29: (cell(29, "I"), cell(29, "M")),
            30: (cell(30, "I"), cell(30, "M")),
            31: (cell(31, "I"), cell(31, "M")),
            32: (cell(32, "I"), cell(32, "M")),
            33: (cell(33, "I"), cell(33, "M")),
            34: (cell(34, "I"), cell(34, "M")),
        },
        "_src_file": src_display,
        "_src_path": filepath,
        "_mtime": mtime_str,
    }


def read_seizo_denpyo_sheet(
    filepath: str,
    cancel_check=None,
    base_folders: list = None,
    mtime_sec: float = None,
) -> dict:
    """
    制作伝票（.xls / .xlsx / .xlsm）から管理Noおよび指定行の項目を高速抽出して返す。
    必ず読み取り専用で解析し、参照元ファイルを一切変更しない。
    サンプル（22601013.xlsm / 22601003.xlsm）とフォーマットが異なる場合は FormatValidationError を送出する。
    """
    lower = filepath.lower()
    if lower.endswith(".xls"):
        raw_bytes = open_readonly_bytes(filepath, cancel_check=cancel_check)
        if not raw_bytes:
            raise FormatValidationError(
                "制作伝票ファイル",
                filepath,
                "ファイルサイズが0バイト（空ファイル）です。",
                "シート「製作伝票」を含む制作伝票Excelファイル",
                "空ファイル (0 bytes)",
            )
        cells, matched_sheet, _ = _read_seizo_cells_xls(filepath, raw_bytes)
    else:
        if os.path.exists(filepath) and os.path.getsize(filepath) == 0:
            raise FormatValidationError(
                "制作伝票ファイル",
                filepath,
                "ファイルサイズが0バイト（空ファイル）です。",
                "シート「製作伝票」を含む制作伝票Excelファイル",
                "空ファイル (0 bytes)",
            )
        cells, matched_sheet, _ = _read_seizo_cells_xml(filepath, raw_bytes=None, cancel_check=cancel_check)

    return _build_seizo_record_from_cells(
        filepath,
        cells,
        matched_sheet,
        base_folders=base_folders,
        mtime_sec=mtime_sec,
    )


def read_daily_report_csv(filepath: str, cancel_check=None):
    """
    日報CSVファイルを読み取り専用で読み込み、(行リスト, データ内最新更新日文字列) を返す。
    UTF-8(BOM有無)およびCP932(Shift-JIS)に自動対応する。
    サンプル（daily_report.csv）とフォーマットが異なる場合は FormatValidationError を送出する。
    """
    raw = open_readonly_bytes(filepath, cancel_check=cancel_check)
    if not raw:
        raise FormatValidationError(
            "日報ファイル",
            filepath,
            "ファイルサイズが0バイト（空ファイル）です。",
            "1行目に見出し（管理No, 更新日, 作業者, ..., 作業コード, 作業名, 合計）を含むCSV",
            "空ファイル (0 bytes)",
        )
    if raw.startswith(b"PK\x03\x04") or raw.startswith(b"\xd0\xcf\x11\xe0"):
        raise FormatValidationError(
            "日報ファイル",
            filepath,
            "CSVファイルではなくExcelファイル（バイナリ）が指定されています。",
            "カンマ区切りCSVファイル（サンプル: daily_report.csv）",
            "Excelバイナリファイル (.xlsx / .xlsm / .xls)",
        )
    if b"\x00" in raw[:4096]:
        raise FormatValidationError(
            "日報ファイル",
            filepath,
            "テキストCSV形式ではないバイナリファイルが指定されています。",
            "カンマ区切りCSVファイル（UTF-8 / Shift-JIS）",
            "バイナリデータ",
        )

    text = None
    for enc in ("utf-8-sig", "cp932", "utf-8", "shift_jis"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = raw.decode("utf-8", errors="replace")

    reader = [r for r in csv.reader(text.splitlines()) if any(c.strip() for c in r)]
    if not reader:
        raise FormatValidationError(
            "日報ファイル",
            filepath,
            "CSV内に有効なデータ行が存在しません。",
            "1行目に見出し（管理No, 更新日, 作業者, ..., 作業コード, 作業名, 合計）を含むCSV",
            "有効行 0件",
        )

    header = [c.strip() for c in reader[0]]
    normalized_header = [c.rstrip(".") if c.startswith("管理No") else c for c in header]

    required_cols = ("管理No", "作業コード", "合計")
    missing_cols = [req for req in required_cols if req not in normalized_header]
    if missing_cols:
        detected_preview = ", ".join(header[:10]) if header else "(見出しなし)"
        raise FormatValidationError(
            "日報ファイル",
            filepath,
            f"サンプル（daily_report.csv）と列フォーマットが異なります。1行目に必須列（{', '.join(missing_cols)}）が見つかりません。",
            "1行目見出し: 管理No, 更新日, 作業者, 出荷先, プロジェクト名, 品名、ライン名等, 製品詳細, 数量, 単位, 型枠コード, 型枠班, 作業コード, 作業名, 合計",
            f"1行目の検出列: {detected_preview}",
        )

    col_map = {
        "管理No": -1,
        "更新日": -1,
        "作業者": -1,
        "出荷先": -1,
        "プロジェクト名": -1,
        "品名、ライン名等": -1,
        "製品詳細": -1,
        "数量": -1,
        "単位": -1,
        "型枠コード": -1,
        "型枠班": -1,
        "作業コード": -1,
        "作業名": -1,
        "合計": -1,
    }
    for idx, col_name in enumerate(normalized_header):
        if col_name in col_map and col_map[col_name] == -1:
            col_map[col_name] = idx

    idx_kno = col_map["管理No"]
    idx_upd = col_map["更新日"]
    idx_worker = col_map["作業者"]
    idx_shukka = col_map["出荷先"]
    idx_proj = col_map["プロジェクト名"]
    idx_hinmei = col_map["品名、ライン名等"]
    idx_detail = col_map["製品詳細"]
    idx_qty = col_map["数量"]
    idx_unit = col_map["単位"]
    idx_code = col_map["作業コード"]
    idx_wname = col_map["作業名"]
    idx_total = col_map["合計"]

    records = []
    latest_update = ""
    src_name = os.path.basename(filepath)

    for row_idx, row in enumerate(reader[1:]):
        if cancel_check and (row_idx & 1023) == 0 and cancel_check():
            raise InterruptedError("処理がキャンセルされました")
        r_len = len(row)
        if idx_kno >= r_len:
            continue
        kno = normalize_kanri_no(row[idx_kno])
        if not kno:
            continue
        upd = row[idx_upd].strip() if 0 <= idx_upd < r_len else ""
        if upd and upd > latest_update:
            latest_update = upd
        code_str = row[idx_code].strip() if 0 <= idx_code < r_len else ""
        total_str = row[idx_total].strip() if 0 <= idx_total < r_len else ""
        records.append({
            "管理No": kno,
            "更新日": upd,
            "作業者": row[idx_worker].strip() if 0 <= idx_worker < r_len else "",
            "出荷先": row[idx_shukka].strip() if 0 <= idx_shukka < r_len else "",
            "プロジェクト名": row[idx_proj].strip() if 0 <= idx_proj < r_len else "",
            "品名、ライン名等": row[idx_hinmei].strip() if 0 <= idx_hinmei < r_len else "",
            "製品詳細": row[idx_detail].strip() if 0 <= idx_detail < r_len else "",
            "数量": row[idx_qty].strip() if 0 <= idx_qty < r_len else "",
            "単位": row[idx_unit].strip() if 0 <= idx_unit < r_len else "",
            "作業コード": code_str,
            "作業名": row[idx_wname].strip() if 0 <= idx_wname < r_len else "",
            "合計": to_float(total_str, 0.0),
            "ファイル名": src_name,
        })

    if latest_update and "." in latest_update:
        latest_update = latest_update.split(".")[0]
    latest_update = latest_update.replace("-", "/")
    return records, latest_update


def escape_tsv_cell(val) -> str:
    """Excel等への貼り付け用TSVセル文字列をエスケープする（タブ・改行・クォートによる列ズレ防止）"""
    s = "" if val is None else str(val)
    if any(ch in s for ch in ("\t", "\n", "\r", '"')):
        return '"' + s.replace('"', '""') + '"'
    return s


def is_same_or_protected_path(target_path: str, protected_paths: set, protected_dirs: set = None) -> bool:
    """
    保存先パスが参照元ファイルまたは参照元フォルダ（制作伝票フォルダおよびサブフォルダ）に
    一致・包含されるか判定する。一致・包含される場合は変更・書き込みを厳重にブロックする。
    """
    if not target_path:
        return False
    norm_target = os.path.normcase(os.path.abspath(os.path.normpath(target_path)))
    if protected_paths:
        if norm_target in protected_paths:
            return True
        if os.path.exists(target_path):
            for p in protected_paths:
                if os.path.exists(p):
                    try:
                        if os.path.samefile(target_path, p):
                            return True
                    except Exception:
                        pass
    if protected_dirs:
        if norm_target in protected_dirs:
            return True
        for d in protected_dirs:
            if not d:
                continue
            try:
                if os.path.commonpath([norm_target, d]) == d:
                    return True
            except ValueError:
                pass
    return False


def _header_style_kind(h_text: str) -> str:
    """表示単位付き見出しからExcelヘッダースタイル種別を判定する"""
    if "小計" in h_text or "材料費計" in h_text or "合計" in h_text or "総請求金額" in h_text:
        return "subtotal"
    if h_text in WORK_CATEGORIES or any(
        k in h_text
        for k in ("作業費", "移動費", "修理改造費", "設計費(関係)", "その他作業費", "作業時間", "移動時間", "修理改造時間", "設計関係", "その他作業", "工数")
    ):
        return "work"
    if h_text in ("購入品", "外注加工") or any(
        k in h_text
        for k in ("購入品費", "外注加工費", "調整用", "材料費", "切断材料", "設計費", "消耗品費", "重機、運搬費", "車両、交通費", "出張費", "管理費", "諸経費", "調整額")
    ):
        return "cost"
    return "base"


def export_to_xlsx(
    filepath: str,
    headers: list,
    rows: list,
    numeric_col_indices: set,
    totals_row: list = None,
    protected_paths: set = None,
    col_groups: list = None,
    protected_dirs: set = None,
):
    """
    集計結果をExcel(.xlsx)に出力する。
    openpyxlがインストールされていればスタイル・オートフィルタ・合計行・列グループ化付きで出力し、
    未インストール環境でも標準ライブラリのみで正しい.xlsxを生成する。
    参照元ファイルおよび参照元フォルダ内への上書きは厳重にブロックする。
    """
    if is_same_or_protected_path(filepath, protected_paths or set(), protected_dirs or set()):
        raise PermissionError(f"参照元ファイル保護のため、上書きをブロックしました:\n{filepath}")
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "制作伝票一覧"

        header_font = Font(name="Yu Gothic UI", size=10, bold=True, color="0F172A")
        header_fill_base = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
        header_fill_work = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
        header_fill_cost = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
        header_fill_sub = PatternFill(start_color="FDE68A", end_color="FDE68A", fill_type="solid")
        total_fill = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")

        data_font = Font(name="Yu Gothic UI", size=10, color="0F172A")
        total_font = Font(name="Yu Gothic UI", size=10, bold=True, color="0F172A")
        thin_side = Side(border_style="thin", color="CBD5E1")
        border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        ws.append(headers)
        ws.row_dimensions[1].height = 24
        for col_idx, h_text in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            kind = _header_style_kind(h_text)
            if kind == "subtotal":
                cell.fill = header_fill_sub
            elif kind == "work":
                cell.fill = header_fill_work
            elif kind == "cost":
                cell.fill = header_fill_cost
            else:
                cell.fill = header_fill_base

        for r_idx, row_data in enumerate(rows, 2):
            for c_idx, val in enumerate(row_data, 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                if (c_idx - 1) in numeric_col_indices and val != "":
                    num_val = to_float(val, 0.0)
                    if abs(num_val - round(num_val)) < 1e-6:
                        cell.value = int(round(num_val))
                        cell.number_format = "#,##0"
                    else:
                        cell.value = round(num_val, 2)
                        cell.number_format = "#,##0.00"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.value = str(val) if val is not None else ""
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                cell.font = data_font
                cell.border = border

        last_data_row = len(rows) + 1
        if len(headers) >= 1 and last_data_row >= 1:
            ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{last_data_row}"

        if totals_row:
            tot_r = last_data_row + 1
            for c_idx, val in enumerate(totals_row, 1):
                cell = ws.cell(row=tot_r, column=c_idx)
                if (c_idx - 1) in numeric_col_indices and val != "":
                    num_val = to_float(val, 0.0)
                    if abs(num_val - round(num_val)) < 1e-6:
                        cell.value = int(round(num_val))
                        cell.number_format = "#,##0"
                    else:
                        cell.value = round(num_val, 2)
                        cell.number_format = "#,##0.00"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.value = str(val) if val is not None else ""
                    cell.alignment = Alignment(horizontal="center" if c_idx == 1 else "left", vertical="center")
                cell.font = total_font
                cell.fill = total_fill
                cell.border = border

        ws.freeze_panes = "B2"
        for col_idx in range(1, len(headers) + 1):
            col_letter = get_column_letter(col_idx)
            max_len = max(
                [len(str(headers[col_idx - 1])) * 2 + 2]
                + [len(str(r[col_idx - 1])) for r in rows[:300]]
            )
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 11), 42)

        # 列グループ化（小計1・小計2の内訳列をデフォルトで折りたたみ、右隣の小計列上に「+」を表示）
        if col_groups:
            ws.sheet_properties.outlinePr.summaryRight = True
            for start_c, end_c in col_groups:
                if 1 <= start_c <= end_c <= len(headers):
                    ws.column_dimensions.group(
                        get_column_letter(start_c),
                        get_column_letter(end_c),
                        outline_level=1,
                        hidden=True,
                    )

        # 印刷レイアウト設定（A3横デフォルト・全列1ページ収め・見出し行繰り返し・ヘッダー/フッター・ページレイアウト表示）
        # 列グループの表示/非表示（折りたたみ/展開）に自動適応して印刷される
        total_last_row = last_data_row + (1 if totals_row else 0)
        ws.page_setup.paperSize = ws.PAPERSIZE_A3
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_margins.top = 0.55
        ws.page_margins.bottom = 0.55
        ws.page_margins.left = 0.32
        ws.page_margins.right = 0.32
        ws.page_margins.header = 0.24
        ws.page_margins.footer = 0.24
        ws.print_options.horizontalCentered = True
        if len(headers) >= 1 and total_last_row >= 1:
            ws.print_area = f"A1:{get_column_letter(len(headers))}{total_last_row}"
            ws.print_title_rows = "1:1"
        ws.oddHeader.left.text = "制作伝票一覧（DenpyoHub_Miya）"
        ws.oddHeader.left.size = 11
        ws.oddHeader.left.font = "Yu Gothic UI,Bold"
        ws.oddHeader.right.text = "出力日時: &D &T"
        ws.oddHeader.right.size = 9
        ws.oddFooter.right.text = "&P / &N ページ"
        ws.oddFooter.right.size = 9
        try:
            ws.views.sheetView[0].view = "pageLayout"
            ws.views.sheetView[0].zoomScale = 85
        except Exception:
            pass

        wb.save(filepath)
        return
    except ImportError:
        pass

    # openpyxl未導入環境向けの標準ライブラリZip+XMLによるXLSX出力
    def col_name(idx):
        res = ""
        n = idx + 1
        while n > 0:
            n, rem = divmod(n - 1, 26)
            res = chr(65 + rem) + res
        return res

    from xml.sax.saxutils import escape

    all_rows = [headers] + rows
    if totals_row:
        all_rows.append(totals_row)

    grouped_col_set = set()
    for start_c, end_c in (col_groups or []):
        for c_num in range(start_c, end_c + 1):
            if 1 <= c_num <= len(headers):
                grouped_col_set.add(c_num)

    cols_xml = ""
    if headers:
        col_elems = []
        for c_num in range(1, len(headers) + 1):
            if c_num in grouped_col_set:
                col_elems.append(f'<col min="{c_num}" max="{c_num}" width="14" customWidth="1" outlineLevel="1" hidden="1"/>')
            else:
                col_elems.append(f'<col min="{c_num}" max="{c_num}" width="15" customWidth="1"/>')
        cols_xml = "<cols>" + "".join(col_elems) + "</cols>"

    sheet_lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<sheetPr><outlinePr summaryRight="1"/><pageSetUpPr fitToPage="1"/></sheetPr>',
        '<sheetViews><sheetView tabSelected="1" view="pageLayout" zoomScale="85" workbookViewId="0"/></sheetViews>',
        cols_xml,
        "<sheetData>",
    ]
    for r_i, r_data in enumerate(all_rows, 1):
        sheet_lines.append(f'<row r="{r_i}">')
        for c_i, val in enumerate(r_data):
            ref = f"{col_name(c_i)}{r_i}"
            if r_i > 1 and c_i in numeric_col_indices and val != "":
                num_val = to_float(val, 0.0)
                sheet_lines.append(f'<c r="{ref}"><v>{num_val}</v></c>')
            else:
                text_escaped = escape(str(val if val is not None else ""))
                sheet_lines.append(f'<c r="{ref}" t="inlineStr"><is><t>{text_escaped}</t></is></c>')
        sheet_lines.append("</row>")
    sheet_lines.append("</sheetData>")
    sheet_lines.append('<printOptions horizontalCentered="1"/>')
    sheet_lines.append('<pageMargins left="0.32" right="0.32" top="0.55" bottom="0.55" header="0.24" footer="0.24"/>')
    sheet_lines.append('<pageSetup paperSize="8" orientation="landscape" fitToWidth="1" fitToHeight="0"/>')
    sheet_lines.append(
        "<headerFooter>"
        '<oddHeader>&amp;L&amp;"Yu Gothic UI,Bold"&amp;11制作伝票一覧（DenpyoHub_Miya）&amp;R&amp;"Yu Gothic UI"&amp;9出力日時: &amp;D &amp;T</oddHeader>'
        '<oddFooter>&amp;R&amp;"Yu Gothic UI"&amp;9&amp;P / &amp;N ページ</oddFooter>'
        "</headerFooter>"
    )
    sheet_lines.append("</worksheet>")

    last_col_letter = col_name(max(0, len(headers) - 1))
    last_row_num = max(1, len(all_rows))
    defined_names_xml = (
        "<definedNames>"
        f'<definedName name="_xlnm.Print_Titles" localSheetId="0">\'制作伝票一覧\'!$1:$1</definedName>'
        f'<definedName name="_xlnm.Print_Area" localSheetId="0">\'制作伝票一覧\'!$A$1:${last_col_letter}${last_row_num}</definedName>'
        "</definedNames>"
    )

    with zipfile.ZipFile(filepath, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            "</Types>",
        )
        zf.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>",
        )
        zf.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            "</Relationships>",
        )
        zf.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="制作伝票一覧" sheetId="1" r:id="rId1"/></sheets>'
            f"{defined_names_xml}"
            "</workbook>",
        )
        zf.writestr("xl/worksheets/sheet1.xml", "".join(sheet_lines))


def open_in_excel_new_workbook(
    headers: list,
    rows: list,
    numeric_col_indices: set,
    totals_row: list = None,
    protected_paths: set = None,
    col_groups: list = None,
    protected_dirs: set = None,
):
    """
    集計結果を保存ファイルではなくExcelの新規ブック（未保存ブック）として開く。
    win32com.client（Excel COM）が利用可能な場合は新規ブックに直接高速書き込みし、
    利用不可の場合は一時ファイルを開いてExcelで表示する。
    """
    if sys.platform == "win32":
        try:
            pythoncom = importlib.import_module("pythoncom")
            win32_client = importlib.import_module("win32com.client")

            pythoncom.CoInitialize()
            excel = None
            wb = None
            wb_completed = False
            try:
                excel = win32_client.Dispatch("Excel.Application")
                excel.Visible = True
                try:
                    excel.ScreenUpdating = False
                except Exception:
                    pass

                wb = excel.Workbooks.Add()
                ws = wb.ActiveSheet
                try:
                    ws.Name = "制作伝票一覧"
                except Exception:
                    pass

                all_rows = [headers]
                for r in rows:
                    converted = []
                    for c_i, val in enumerate(r):
                        if c_i in numeric_col_indices and val != "":
                            num_val = to_float(val, 0.0)
                            if abs(num_val - round(num_val)) < 1e-6:
                                converted.append(int(round(num_val)))
                            else:
                                converted.append(round(num_val, 2))
                        else:
                            converted.append("" if val is None else str(val))
                    all_rows.append(converted)

                if totals_row:
                    tot_conv = []
                    for c_i, val in enumerate(totals_row):
                        if c_i in numeric_col_indices and val != "":
                            num_val = to_float(val, 0.0)
                            if abs(num_val - round(num_val)) < 1e-6:
                                tot_conv.append(int(round(num_val)))
                            else:
                                tot_conv.append(round(num_val, 2))
                        else:
                            tot_conv.append("" if val is None else str(val))
                    all_rows.append(tot_conv)

                n_rows = len(all_rows)
                n_cols = len(headers)

                # 文字列列（管理No等）が指数表示にならないようテキスト書式を設定
                if n_rows >= 2:
                    for c_idx in range(1, n_cols + 1):
                        if (c_idx - 1) not in numeric_col_indices:
                            ws.Range(ws.Cells(2, c_idx), ws.Cells(n_rows, c_idx)).NumberFormatLocal = "@"
                        else:
                            # 小数点があるかチェック
                            has_dec = any(
                                isinstance(all_rows[r_i][c_idx - 1], float) and abs(all_rows[r_i][c_idx - 1] - round(all_rows[r_i][c_idx - 1])) >= 1e-6
                                for r_i in range(1, n_rows)
                            )
                            ws.Range(ws.Cells(2, c_idx), ws.Cells(n_rows, c_idx)).NumberFormatLocal = "#,##0.00" if has_dec else "#,##0"

                # 2次元配列で高速書き込み（大量行時は4,000行単位でチャンク書込しCOMメモリ上限を回避）
                chunk_rows = 4000
                if n_rows <= chunk_rows:
                    ws.Range(ws.Cells(1, 1), ws.Cells(n_rows, n_cols)).Value = tuple(tuple(r) for r in all_rows)
                else:
                    for start_r in range(0, n_rows, chunk_rows):
                        end_r = min(start_r + chunk_rows, n_rows)
                        ws.Range(ws.Cells(start_r + 1, 1), ws.Cells(end_r, n_cols)).Value = tuple(
                            tuple(r) for r in all_rows[start_r:end_r]
                        )

                # ヘッダースタイル（BGR値: Excel COMのColorはBGR整数）
                def rgb_to_bgr(r, g, b):
                    return r + (g << 8) + (b << 16)

                hdr_rng = ws.Range(ws.Cells(1, 1), ws.Cells(1, n_cols))
                hdr_rng.Font.Bold = True
                hdr_rng.HorizontalAlignment = -4108  # xlCenter
                for c_idx, h_text in enumerate(headers, 1):
                    cell = ws.Cells(1, c_idx)
                    kind = _header_style_kind(h_text)
                    if kind == "subtotal":
                        cell.Interior.Color = rgb_to_bgr(253, 230, 138)
                    elif kind == "work":
                        cell.Interior.Color = rgb_to_bgr(220, 252, 231)
                    elif kind == "cost":
                        cell.Interior.Color = rgb_to_bgr(254, 243, 199)
                    else:
                        cell.Interior.Color = rgb_to_bgr(219, 234, 254)

                # 合計行スタイル
                if totals_row:
                    tot_rng = ws.Range(ws.Cells(n_rows, 1), ws.Cells(n_rows, n_cols))
                    tot_rng.Font.Bold = True
                    tot_rng.Interior.Color = rgb_to_bgr(224, 242, 254)

                # 罫線・フォント・行高さ・オートフィルタ・列幅自動調整
                full_rng = ws.Range(ws.Cells(1, 1), ws.Cells(n_rows, n_cols))
                try:
                    full_rng.Font.Name = "Yu Gothic UI"
                    full_rng.Font.Size = 9.5
                    full_rng.VerticalAlignment = -4108  # xlCenter
                    ws.Rows(1).RowHeight = 26
                    hdr_rng.WrapText = True
                    if n_rows >= 2:
                        ws.Range(ws.Rows(2), ws.Rows(n_rows)).RowHeight = 19
                    if totals_row:
                        ws.Rows(n_rows).RowHeight = 22
                except Exception:
                    pass
                full_rng.Borders.LineStyle = 1  # xlContinuous
                full_rng.Borders.Color = rgb_to_bgr(203, 213, 225)
                if totals_row:
                    try:
                        # 合計行の上罫線を濃い線の二重線/中太線にして視認性を高める
                        tot_top_border = ws.Range(ws.Cells(n_rows, 1), ws.Cells(n_rows, n_cols)).Borders(8)  # xlEdgeTop
                        tot_top_border.LineStyle = -4119  # xlDouble
                        tot_top_border.Color = rgb_to_bgr(37, 99, 235)
                    except Exception:
                        pass
                last_data_r = n_rows - 1 if totals_row else n_rows
                if last_data_r >= 1:
                    ws.Range(ws.Cells(1, 1), ws.Cells(last_data_r, n_cols)).AutoFilter()

                try:
                    ws.Columns.AutoFit()
                    for c_idx in range(1, n_cols + 1):
                        cur_w = ws.Columns(c_idx).ColumnWidth
                        ws.Columns(c_idx).ColumnWidth = min(max(cur_w + 2, 10), 42)
                except Exception:
                    pass

                # 列グループ化（小計1・小計2の内訳列をグループ化し、デフォルトで折りたたむ）
                # 折りたたまれた非表示列は印刷対象から自動で除外され、展開すると印刷にも自動反映される
                if col_groups:
                    try:
                        ws.Outline.SummaryColumn = 1  # xlSummaryOnRight（右隣の小計列上に「+」ボタンを配置）
                        for start_c, end_c in col_groups:
                            if 1 <= start_c <= end_c <= n_cols:
                                ws.Range(ws.Columns(start_c), ws.Columns(end_c)).Group()
                        ws.Outline.ShowLevels(ColumnLevels=1)
                    except Exception:
                        pass

                # 印刷ページ設定（A3対応プリンタの場合はA3横をデフォルト、非対応時はA4横へ自動フォールバック）
                try:
                    try:
                        excel.PrintCommunication = False
                    except Exception:
                        pass
                    ps = ws.PageSetup
                    ps.Orientation = 2     # xlLandscape (横向き)
                    ps.Zoom = False
                    ps.FitToPagesWide = 1  # 横方向は全列を1ページに収める
                    ps.FitToPagesTall = False
                    ps.TopMargin = excel.InchesToPoints(0.55)
                    ps.BottomMargin = excel.InchesToPoints(0.55)
                    ps.LeftMargin = excel.InchesToPoints(0.32)
                    ps.RightMargin = excel.InchesToPoints(0.32)
                    ps.HeaderMargin = excel.InchesToPoints(0.24)
                    ps.FooterMargin = excel.InchesToPoints(0.24)
                    ps.CenterHorizontally = True
                    ps.CenterVertically = False
                    ps.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(n_rows, n_cols)).Address
                    ps.PrintTitleRows = "$1:$1"
                    ps.LeftHeader = '&"Yu Gothic UI,Bold"&11制作伝票一覧（DenpyoHub_Miya）'
                    ps.RightHeader = '&"Yu Gothic UI"&9出力日時: &D &T'
                    ps.RightFooter = '&"Yu Gothic UI"&9&P / &N ページ'
                except Exception:
                    pass
                finally:
                    try:
                        excel.PrintCommunication = True
                    except Exception:
                        pass

                # プリンタ通信を有効化した状態でA3(xlPaperA3=8)を実適用し、非対応プリンタならA4(xlPaperA4=9)へ安全にフォールバック
                try:
                    ps = ws.PageSetup
                    try:
                        ps.PaperSize = 8  # xlPaperA3
                        if getattr(ps, "PaperSize", 8) != 8:
                            ps.PaperSize = 9  # xlPaperA4
                    except Exception:
                        try:
                            ps.PaperSize = 9  # xlPaperA4
                        except Exception:
                            pass
                except Exception:
                    pass

                # 印刷レイアウト表示（ページレイアウトビュー）で開く
                try:
                    excel.ActiveWindow.View = 3  # xlPageLayoutView
                    excel.ActiveWindow.Zoom = 85
                except Exception:
                    try:
                        excel.ActiveWindow.SplitRow = 1
                        excel.ActiveWindow.FreezePanes = True
                    except Exception:
                        pass
                wb_completed = True
                return
            finally:
                if wb is not None and not wb_completed:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
                if excel is not None:
                    try:
                        excel.ScreenUpdating = True
                    except Exception:
                        pass
                pythoncom.CoUninitialize()
        except ImportError:
            pass
        except Exception:
            pass

    # COMが利用できない場合のフォールバック（一時ファイルを作成して開く）
    tmp_dir = tempfile.mkdtemp(prefix="denpyohub_excel_")
    tmp_path = os.path.join(tmp_dir, f"制作伝票一覧_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    export_to_xlsx(
        tmp_path,
        headers,
        rows,
        numeric_col_indices,
        totals_row=totals_row,
        protected_paths=protected_paths,
        col_groups=col_groups,
        protected_dirs=protected_dirs,
    )
    if sys.platform == "win32":
        os.startfile(tmp_path)
    return tmp_path


def _sanitize_sheet_name(name: str, used_names: set) -> str:
    """Excelシート名として有効かつ一意な文字列（31文字以内）に整形する"""
    s = str(name or "明細").strip()
    for ch in ("[", "]", ":", "*", "?", "/", "\\"):
        s = s.replace(ch, "_")
    s = s.strip("'") or "明細"
    base = s[:31]
    cand = base
    suffix_i = 2
    while cand.lower() in used_names:
        suf = f"_{suffix_i}"
        cand = base[: 31 - len(suf)] + suf
        suffix_i += 1
    used_names.add(cand.lower())
    return cand


def export_multi_sheet_xlsx(
    filepath: str,
    sheets_spec: list,
    protected_paths: set = None,
    protected_dirs: set = None,
):
    """
    内訳明細（単一タブまたは全5タブ）を印刷レイアウト設定済みの複数シートExcel(.xlsx)に出力する。
    openpyxl利用時は装飾・折り返し・合計行・A4横1ページ幅収め・見出し行繰り返し・ページレイアウト表示を設定し、
    未導入環境でも標準ライブラリのみで正しい複数シート.xlsxを生成する。
    """
    if is_same_or_protected_path(filepath, protected_paths or set(), protected_dirs or set()):
        raise PermissionError(f"参照元ファイル保護のため、上書きをブロックしました:\n{filepath}")
    if not sheets_spec:
        return

    used_names = set()
    normalized_specs = []
    for sp in sheets_spec:
        sp_copy = dict(sp)
        sp_copy["safe_name"] = _sanitize_sheet_name(sp.get("name", "明細"), used_names)
        normalized_specs.append(sp_copy)

    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        wb = openpyxl.Workbook()
        title_font = Font(name="Yu Gothic UI", size=12, bold=True, color="0F172A")
        sub_font = Font(name="Yu Gothic UI", size=9, color="334155")
        hdr_font = Font(name="Yu Gothic UI", size=9.5, bold=True, color="0F172A")
        data_font = Font(name="Yu Gothic UI", size=9.5, color="0F172A")
        key_font = Font(name="Yu Gothic UI", size=9.5, bold=True, color="1E3A8A")
        total_font = Font(name="Yu Gothic UI", size=10, bold=True, color="0F172A")

        title_fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
        sub_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        key_fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
        total_fill = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
        hdr_fills = {
            "daily": PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid"),
            "kounyu": PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid"),
            "gaichu": PatternFill(start_color="FFEDD5", end_color="FFEDD5", fill_type="solid"),
            "juchu": PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid"),
            "seizo": PatternFill(start_color="E0E7FF", end_color="E0E7FF", fill_type="solid"),
        }

        thin_side = Side(border_style="thin", color="CBD5E1")
        dbl_top = Side(border_style="double", color="2563EB")
        border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        total_border = Border(left=thin_side, right=thin_side, top=dbl_top, bottom=thin_side)

        for idx_s, sp in enumerate(normalized_specs):
            ws = wb.active if idx_s == 0 else wb.create_sheet()
            ws.title = sp["safe_name"]

            headers = sp.get("headers") or []
            rows = sp.get("rows") or []
            n_cols = max(1, len(headers))
            num_cols = set(sp.get("numeric_col_indices") or ())
            dec_cols = set(sp.get("decimal_col_indices") or ())
            wrap_cols = set(sp.get("wrap_col_indices") or ())
            kv_cols = set(sp.get("key_col_indices") or ())
            totals_row = sp.get("totals_row")
            kind = sp.get("kind", "juchu")
            h_fill = hdr_fills.get(kind, hdr_fills["juchu"])

            # 1行目: タイトルバナー / 2行目: 案件情報サブタイトル
            ws.cell(row=1, column=1, value=sp.get("title", sp["safe_name"]))
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
            ws.row_dimensions[1].height = 26
            for c_i in range(1, n_cols + 1):
                c1 = ws.cell(row=1, column=c_i)
                c1.font = title_font
                c1.fill = title_fill
                c1.alignment = Alignment(horizontal="left", vertical="center")

            ws.cell(row=2, column=1, value=sp.get("subtitle", ""))
            ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
            ws.row_dimensions[2].height = 20
            for c_i in range(1, n_cols + 1):
                c2 = ws.cell(row=2, column=c_i)
                c2.font = sub_font
                c2.fill = sub_fill
                c2.alignment = Alignment(horizontal="left", vertical="center")

            # 3行目: 列見出し
            ws.row_dimensions[3].height = 24
            for c_idx, h_text in enumerate(headers, 1):
                cell = ws.cell(row=3, column=c_idx, value=h_text)
                cell.font = hdr_font
                cell.fill = h_fill
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            # 4行目以降: 明細データ
            cur_r = 4
            if rows:
                for row_data in rows:
                    ws.row_dimensions[cur_r].height = 20
                    for c_idx, val in enumerate(row_data, 1):
                        c0 = c_idx - 1
                        cell = ws.cell(row=cur_r, column=c_idx)
                        if c0 in num_cols and val != "" and val is not None:
                            num_v = to_float(val, 0.0)
                            if c0 in dec_cols:
                                cell.value = round(num_v, 2)
                                cell.number_format = "#,##0.00"
                            elif abs(num_v - round(num_v)) < 1e-6:
                                cell.value = int(round(num_v))
                                cell.number_format = "#,##0"
                            else:
                                cell.value = round(num_v, 2)
                                cell.number_format = "#,##0.00"
                            cell.alignment = Alignment(horizontal="right", vertical="center")
                            cell.font = data_font
                        else:
                            cell.value = "" if val is None else str(val)
                            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=(c0 in wrap_cols))
                            if c0 in kv_cols:
                                cell.font = key_font
                                cell.fill = key_fill
                            else:
                                cell.font = data_font
                        cell.border = border
                    cur_r += 1
            else:
                ws.cell(row=cur_r, column=1, value="（対象データはありません）")
                ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=n_cols)
                ws.row_dimensions[cur_r].height = 22
                for c_i in range(1, n_cols + 1):
                    ce = ws.cell(row=cur_r, column=c_i)
                    ce.font = sub_font
                    ce.border = border
                    ce.alignment = Alignment(horizontal="center", vertical="center")
                cur_r += 1

            # 合計行
            if totals_row:
                ws.row_dimensions[cur_r].height = 22
                for c_idx, val in enumerate(totals_row, 1):
                    c0 = c_idx - 1
                    cell = ws.cell(row=cur_r, column=c_idx)
                    if c0 in num_cols and val != "" and val is not None:
                        num_v = to_float(val, 0.0)
                        if c0 in dec_cols:
                            cell.value = round(num_v, 2)
                            cell.number_format = "#,##0.00"
                        elif abs(num_v - round(num_v)) < 1e-6:
                            cell.value = int(round(num_v))
                            cell.number_format = "#,##0"
                        else:
                            cell.value = round(num_v, 2)
                            cell.number_format = "#,##0.00"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.value = "" if val is None else str(val)
                        cell.alignment = Alignment(horizontal="center" if c_idx == 1 else "left", vertical="center")
                    cell.font = total_font
                    cell.fill = total_fill
                    cell.border = total_border
                cur_r += 1

            last_row = cur_r - 1
            for c_idx in range(1, n_cols + 1):
                col_letter = get_column_letter(c_idx)
                hdr_w = len(str(headers[c_idx - 1])) * 2 + 3 if c_idx - 1 < len(headers) else 12
                sample_w = [
                    sum(2 if ord(ch) > 127 else 1 for ch in str(r[c_idx - 1]))
                    for r in rows[:200]
                    if c_idx - 1 < len(r)
                ]
                max_len = max([hdr_w] + sample_w) if sample_w else hdr_w
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 42)

            # 印刷設定（A4横向き・横1ページ収め・見出し3行繰り返し・ページレイアウト表示）
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.page_margins.top = 0.55
            ws.page_margins.bottom = 0.55
            ws.page_margins.left = 0.35
            ws.page_margins.right = 0.35
            ws.page_margins.header = 0.24
            ws.page_margins.footer = 0.24
            ws.print_options.horizontalCentered = True
            ws.print_area = f"A1:{get_column_letter(n_cols)}{last_row}"
            ws.print_title_rows = "1:3"
            ws.oddHeader.left.text = sp.get("print_header", sp["safe_name"])
            ws.oddHeader.left.size = 10
            ws.oddHeader.left.font = "Yu Gothic UI,Bold"
            ws.oddHeader.right.text = "出力日時: &D &T"
            ws.oddHeader.right.size = 9
            ws.oddFooter.right.text = "&P / &N ページ"
            ws.oddFooter.right.size = 9
            try:
                ws.views.sheetView[0].view = "pageLayout"
                ws.views.sheetView[0].zoomScale = 85
            except Exception:
                pass

        wb.save(filepath)
        return
    except ImportError:
        pass

    # openpyxl未導入環境向けの標準ライブラリZip+XMLによる複数シートXLSX出力
    from xml.sax.saxutils import escape

    ct_overrides = [
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    ]
    wb_rels = []
    wb_sheets = []
    defined_names = []
    sheet_xml_map = {}

    for idx_s, sp in enumerate(normalized_specs, 1):
        sname = sp["safe_name"]
        headers = sp.get("headers") or []
        rows = sp.get("rows") or []
        n_cols = max(1, len(headers))
        num_cols = set(sp.get("numeric_col_indices") or ())
        totals_row = sp.get("totals_row")

        all_rows = [
            [sp.get("title", sname)] + [""] * (n_cols - 1),
            [sp.get("subtitle", "")] + [""] * (n_cols - 1),
            list(headers),
        ]
        if rows:
            all_rows.extend(rows)
        else:
            all_rows.append(["（対象データはありません）"] + [""] * (n_cols - 1))
        if totals_row:
            all_rows.append(totals_row)

        cols_xml = "<cols>" + "".join(
            f'<col min="{c_num}" max="{c_num}" width="16" customWidth="1"/>' for c_num in range(1, n_cols + 1)
        ) + "</cols>"
        lines = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
            '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>',
            f'<sheetViews><sheetView tabSelected="{1 if idx_s == 1 else 0}" view="pageLayout" zoomScale="85" workbookViewId="0"/></sheetViews>',
            cols_xml,
            "<sheetData>",
        ]
        has_data = bool(rows)
        for r_i, r_data in enumerate(all_rows, 1):
            lines.append(f'<row r="{r_i}">')
            for c_i, val in enumerate(r_data):
                ref = f"{col_index_to_letter(c_i)}{r_i}"
                is_num_cell = (r_i >= 4) and (has_data or r_i != 4) and (c_i in num_cols) and (val != "") and (val is not None)
                if is_num_cell:
                    lines.append(f'<c r="{ref}"><v>{to_float(val, 0.0)}</v></c>')
                else:
                    lines.append(f'<c r="{ref}" t="inlineStr"><is><t>{escape(str(val if val is not None else ""))}</t></is></c>')
            lines.append("</row>")
        lines.append("</sheetData>")
        last_col_l = col_index_to_letter(n_cols - 1)
        if n_cols > 1:
            merges = [f'<mergeCell ref="A1:{last_col_l}1"/>', f'<mergeCell ref="A2:{last_col_l}2"/>']
            if not has_data:
                merges.append(f'<mergeCell ref="A4:{last_col_l}4"/>')
            lines.append(f'<mergeCells count="{len(merges)}">' + "".join(merges) + "</mergeCells>")
        lines.append('<printOptions horizontalCentered="1"/>')
        lines.append('<pageMargins left="0.35" right="0.35" top="0.55" bottom="0.55" header="0.24" footer="0.24"/>')
        lines.append('<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>')
        hdr_escaped = escape(sp.get("print_header", sname))
        lines.append(
            "<headerFooter>"
            f'<oddHeader>&amp;L&amp;"Yu Gothic UI,Bold"&amp;10{hdr_escaped}&amp;R&amp;"Yu Gothic UI"&amp;9出力日時: &amp;D &amp;T</oddHeader>'
            '<oddFooter>&amp;R&amp;"Yu Gothic UI"&amp;9&amp;P / &amp;N ページ</oddFooter>'
            "</headerFooter>"
        )
        lines.append("</worksheet>")

        sheet_path = f"xl/worksheets/sheet{idx_s}.xml"
        sheet_xml_map[sheet_path] = "".join(lines)
        ct_overrides.append(
            f'<Override PartName="/{sheet_path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
        wb_rels.append(
            f'<Relationship Id="rId{idx_s}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{idx_s}.xml"/>'
        )
        sname_esc = escape(sname)
        sname_quoted = sname_esc.replace("'", "''")
        wb_sheets.append(f'<sheet name="{sname_esc}" sheetId="{idx_s}" r:id="rId{idx_s}"/>')
        local_id = idx_s - 1
        defined_names.append(
            f'<definedName name="_xlnm.Print_Titles" localSheetId="{local_id}">\'{sname_quoted}\'!$1:$3</definedName>'
        )
        defined_names.append(
            f'<definedName name="_xlnm.Print_Area" localSheetId="{local_id}">\'{sname_quoted}\'!$A$1:${last_col_l}${len(all_rows)}</definedName>'
        )

    with zipfile.ZipFile(filepath, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            + "".join(ct_overrides)
            + "</Types>",
        )
        zf.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>",
        )
        zf.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(wb_rels)
            + "</Relationships>",
        )
        zf.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            "<sheets>" + "".join(wb_sheets) + "</sheets>"
            "<definedNames>" + "".join(defined_names) + "</definedNames>"
            "</workbook>",
        )
        for spath, sxml in sheet_xml_map.items():
            zf.writestr(spath, sxml)


def open_detail_in_excel_new_workbook(
    kno: str,
    sheets_spec: list,
    active_sheet_name: str = None,
    protected_paths: set = None,
    protected_dirs: set = None,
):
    """
    選択した管理Noの内訳明細（単一タブまたは全5タブ）を印刷レイアウト設定済みの
    Excel新規ブック（未保存ブック）として開く。
    win32com.client（Excel COM）利用時は直接高速書き込み＋印刷レイアウト表示で開き、
    利用不可時は複数シート.xlsx一時ファイル経由で開く。
    """
    if not sheets_spec:
        return None

    used_names = set()
    normalized_specs = []
    for sp in sheets_spec:
        sp_copy = dict(sp)
        sp_copy["safe_name"] = _sanitize_sheet_name(sp.get("name", "明細"), used_names)
        normalized_specs.append(sp_copy)

    if sys.platform == "win32":
        try:
            pythoncom = importlib.import_module("pythoncom")
            win32_client = importlib.import_module("win32com.client")

            def rgb_to_bgr(r, g, b):
                return r + (g << 8) + (b << 16)

            hdr_colors = {
                "daily": rgb_to_bgr(220, 252, 231),
                "kounyu": rgb_to_bgr(254, 243, 199),
                "gaichu": rgb_to_bgr(255, 237, 213),
                "juchu": rgb_to_bgr(219, 234, 254),
                "seizo": rgb_to_bgr(224, 231, 255),
            }

            pythoncom.CoInitialize()
            excel = None
            wb = None
            wb_completed = False
            orig_alerts = True
            try:
                excel = win32_client.Dispatch("Excel.Application")
                excel.Visible = True
                try:
                    orig_alerts = bool(excel.DisplayAlerts)
                    excel.DisplayAlerts = False
                except Exception:
                    pass
                try:
                    excel.ScreenUpdating = False
                except Exception:
                    pass

                wb = excel.Workbooks.Add()
                # 既存シートを優先再利用し、不足分のみ追加、余剰分のみDisplayAlerts=Falseで安全に削除
                created_sheets = []
                need_cnt = len(normalized_specs)
                existing_cnt = int(getattr(wb.Worksheets, "Count", 1))
                for idx_s, sp in enumerate(normalized_specs):
                    if idx_s < existing_cnt:
                        ws = wb.Worksheets(idx_s + 1)
                    else:
                        ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
                    try:
                        ws.Name = sp["safe_name"]
                    except Exception:
                        pass
                    created_sheets.append((sp, ws))
                try:
                    while wb.Worksheets.Count > need_cnt:
                        wb.Worksheets(wb.Worksheets.Count).Delete()
                except Exception:
                    pass

                target_ws = created_sheets[0][1]
                for sp, ws in created_sheets:
                    if active_sheet_name and sp.get("name") == active_sheet_name:
                        target_ws = ws

                    headers = sp.get("headers") or []
                    rows = sp.get("rows") or []
                    n_cols = max(1, len(headers))
                    num_cols = set(sp.get("numeric_col_indices") or ())
                    dec_cols = set(sp.get("decimal_col_indices") or ())
                    wrap_cols = set(sp.get("wrap_col_indices") or ())
                    kv_cols = set(sp.get("key_col_indices") or ())
                    totals_row = sp.get("totals_row")
                    kind = sp.get("kind", "juchu")

                    # テーブルデータ（3行目=見出し、4行目以降=データ、末尾=合計行）を2次元配列で高速書き込み
                    table_rows = [list(headers)]
                    has_data = bool(rows)
                    if has_data:
                        for r in rows:
                            conv = []
                            for c_i in range(n_cols):
                                val = r[c_i] if c_i < len(r) else ""
                                if c_i in num_cols and val != "" and val is not None:
                                    nv = to_float(val, 0.0)
                                    if c_i in dec_cols:
                                        conv.append(round(nv, 2))
                                    elif abs(nv - round(nv)) < 1e-6:
                                        conv.append(int(round(nv)))
                                    else:
                                        conv.append(round(nv, 2))
                                else:
                                    conv.append("" if val is None else str(val))
                            table_rows.append(conv)
                    else:
                        table_rows.append(["（対象データはありません）"] + [""] * (n_cols - 1))

                    if totals_row:
                        t_conv = []
                        for c_i in range(n_cols):
                            val = totals_row[c_i] if c_i < len(totals_row) else ""
                            if c_i in num_cols and val != "" and val is not None:
                                nv = to_float(val, 0.0)
                                if c_i in dec_cols:
                                    t_conv.append(round(nv, 2))
                                elif abs(nv - round(nv)) < 1e-6:
                                    t_conv.append(int(round(nv)))
                                else:
                                    t_conv.append(round(nv, 2))
                            else:
                                t_conv.append("" if val is None else str(val))
                        table_rows.append(t_conv)

                    start_tbl_r = 3
                    last_r = start_tbl_r + len(table_rows) - 1

                    # テキスト列・数値列の書式を書き込み前に設定
                    if last_r >= 4:
                        for c_idx in range(1, n_cols + 1):
                            c0 = c_idx - 1
                            col_rng = ws.Range(ws.Cells(4, c_idx), ws.Cells(last_r, c_idx))
                            if c0 not in num_cols:
                                col_rng.NumberFormatLocal = "@"
                            elif c0 in dec_cols:
                                col_rng.NumberFormatLocal = "#,##0.00"
                            else:
                                has_dec = any(
                                    isinstance(table_rows[r_i][c0], float) and abs(table_rows[r_i][c0] - round(table_rows[r_i][c0])) >= 1e-6
                                    for r_i in range(1, len(table_rows))
                                )
                                col_rng.NumberFormatLocal = "#,##0.00" if has_dec else "#,##0"

                    # 1～2行目タイトル・サブタイトル
                    ws.Cells(1, 1).Value = sp.get("title", sp["safe_name"])
                    ws.Cells(2, 1).Value = sp.get("subtitle", "")
                    if n_cols > 1:
                        try:
                            ws.Range(ws.Cells(1, 1), ws.Cells(1, n_cols)).Merge()
                            ws.Range(ws.Cells(2, 1), ws.Cells(2, n_cols)).Merge()
                        except Exception:
                            pass

                    # 3行目以降を一括高速書き込み（大量行時は4,000行単位でチャンク書込）
                    chunk_rows = 4000
                    n_tbl_rows = len(table_rows)
                    if n_tbl_rows <= chunk_rows:
                        ws.Range(ws.Cells(start_tbl_r, 1), ws.Cells(last_r, n_cols)).Value = tuple(tuple(r) for r in table_rows)
                    else:
                        for st_i in range(0, n_tbl_rows, chunk_rows):
                            ed_i = min(st_i + chunk_rows, n_tbl_rows)
                            ws.Range(
                                ws.Cells(start_tbl_r + st_i, 1),
                                ws.Cells(start_tbl_r + ed_i - 1, n_cols),
                            ).Value = tuple(tuple(r) for r in table_rows[st_i:ed_i])

                    # スタイル設定
                    try:
                        all_rng = ws.Range(ws.Cells(1, 1), ws.Cells(last_r, n_cols))
                        all_rng.Font.Name = "Yu Gothic UI"
                        all_rng.Font.Size = 9.5
                        all_rng.VerticalAlignment = -4108  # xlCenter

                        t1_rng = ws.Range(ws.Cells(1, 1), ws.Cells(1, n_cols))
                        t1_rng.Font.Size = 12
                        t1_rng.Font.Bold = True
                        t1_rng.Interior.Color = rgb_to_bgr(219, 234, 254)
                        ws.Rows(1).RowHeight = 26

                        sub_txt = sp.get("subtitle", "")
                        t2_rng = ws.Range(ws.Cells(2, 1), ws.Cells(2, n_cols))
                        t2_rng.Font.Size = 9
                        t2_rng.Font.Color = rgb_to_bgr(51, 65, 85)
                        t2_rng.Interior.Color = rgb_to_bgr(241, 245, 249)
                        t2_rng.WrapText = True
                        ws.Rows(2).RowHeight = 30 if len(sub_txt) > 85 else 20

                        hdr_rng = ws.Range(ws.Cells(3, 1), ws.Cells(3, n_cols))
                        hdr_rng.Font.Bold = True
                        hdr_rng.HorizontalAlignment = -4108
                        hdr_rng.WrapText = True
                        hdr_rng.Interior.Color = hdr_colors.get(kind, hdr_colors["juchu"])
                        ws.Rows(3).RowHeight = 24

                        tbl_rng = ws.Range(ws.Cells(3, 1), ws.Cells(last_r, n_cols))
                        tbl_rng.Borders.LineStyle = 1  # xlContinuous
                        tbl_rng.Borders.Color = rgb_to_bgr(203, 213, 225)

                        if has_data and kv_cols:
                            data_end_r = last_r - 1 if totals_row else last_r
                            for c0 in kv_cols:
                                kr = ws.Range(ws.Cells(4, c0 + 1), ws.Cells(data_end_r, c0 + 1))
                                kr.Font.Bold = True
                                kr.Interior.Color = rgb_to_bgr(239, 246, 255)
                        if not has_data and n_cols > 1:
                            try:
                                emp_rng = ws.Range(ws.Cells(4, 1), ws.Cells(4, n_cols))
                                emp_rng.Merge()
                                emp_rng.HorizontalAlignment = -4108
                            except Exception:
                                pass

                        if totals_row:
                            tot_rng = ws.Range(ws.Cells(last_r, 1), ws.Cells(last_r, n_cols))
                            tot_rng.Font.Bold = True
                            tot_rng.Interior.Color = rgb_to_bgr(224, 242, 254)
                            tot_top = tot_rng.Borders(8)  # xlEdgeTop
                            tot_top.LineStyle = -4119     # xlDouble
                            tot_top.Color = rgb_to_bgr(37, 99, 235)
                            ws.Rows(last_r).RowHeight = 22

                        for c0 in wrap_cols:
                            if last_r >= 4:
                                ws.Range(ws.Cells(4, c0 + 1), ws.Cells(last_r, c0 + 1)).WrapText = True

                        # タイトル結合行(1～2行)を除いたテーブル範囲で列幅自動調整
                        ws.Range(ws.Cells(3, 1), ws.Cells(last_r, n_cols)).Columns.AutoFit()
                        for c_idx in range(1, n_cols + 1):
                            cur_w = ws.Columns(c_idx).ColumnWidth
                            ws.Columns(c_idx).ColumnWidth = min(max(cur_w + 2.5, 9.5), 42)
                        if last_r >= 4:
                            data_last_r = last_r - 1 if totals_row else last_r
                            if data_last_r >= 4:
                                ws.Range(ws.Rows(4), ws.Rows(data_last_r)).RowHeight = 20
                                if wrap_cols:
                                    ws.Range(ws.Rows(4), ws.Rows(data_last_r)).AutoFit()
                                    for r_i in range(4, data_last_r + 1):
                                        if ws.Rows(r_i).RowHeight < 19:
                                            ws.Rows(r_i).RowHeight = 19
                    except Exception:
                        pass

                    # 印刷設定（A4横向き・横1ページ収め・タイトル/見出し3行毎ページ繰り返し・ヘッダー/フッター）
                    try:
                        try:
                            excel.PrintCommunication = False
                        except Exception:
                            pass
                        ps = ws.PageSetup
                        ps.Orientation = 2     # xlLandscape
                        ps.Zoom = False
                        ps.FitToPagesWide = 1
                        ps.FitToPagesTall = False
                        ps.TopMargin = excel.InchesToPoints(0.55)
                        ps.BottomMargin = excel.InchesToPoints(0.55)
                        ps.LeftMargin = excel.InchesToPoints(0.35)
                        ps.RightMargin = excel.InchesToPoints(0.35)
                        ps.HeaderMargin = excel.InchesToPoints(0.24)
                        ps.FooterMargin = excel.InchesToPoints(0.24)
                        ps.CenterHorizontally = True
                        ps.CenterVertically = False
                        ps.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(last_r, n_cols)).Address
                        ps.PrintTitleRows = "$1:$3"
                        ps.LeftHeader = f'&"Yu Gothic UI,Bold"&10{sp.get("print_header", sp["safe_name"])}'
                        ps.RightHeader = '&"Yu Gothic UI"&9出力日時: &D &T'
                        ps.RightFooter = '&"Yu Gothic UI"&9&P / &N ページ'
                    except Exception:
                        pass
                    finally:
                        try:
                            excel.PrintCommunication = True
                        except Exception:
                            pass

                    try:
                        ws.PageSetup.PaperSize = 9  # xlPaperA4
                    except Exception:
                        pass

                # 各シートを印刷レイアウト表示（ページレイアウトビュー）にし、対象タブをアクティブ化
                for _sp, ws in created_sheets:
                    try:
                        ws.Activate()
                        excel.ActiveWindow.View = 3  # xlPageLayoutView
                        excel.ActiveWindow.Zoom = 85
                    except Exception:
                        pass
                try:
                    target_ws.Activate()
                except Exception:
                    pass
                wb_completed = True
                return None
            finally:
                if wb is not None and not wb_completed:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
                if excel is not None:
                    try:
                        excel.DisplayAlerts = orig_alerts
                    except Exception:
                        pass
                    try:
                        excel.ScreenUpdating = True
                    except Exception:
                        pass
                pythoncom.CoUninitialize()
        except ImportError:
            pass
        except Exception:
            pass

    # COMが利用できない場合のフォールバック（一時ファイルを作成して開く）
    tmp_dir = tempfile.mkdtemp(prefix="denpyohub_detail_excel_")
    safe_kno = "".join(ch for ch in str(kno or "detail") if ch.isalnum() or ch in ("-", "_")) or "detail"
    tmp_path = os.path.join(tmp_dir, f"内訳明細_{safe_kno}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    export_multi_sheet_xlsx(
        tmp_path,
        normalized_specs,
        protected_paths=protected_paths,
        protected_dirs=protected_dirs,
    )
    if sys.platform == "win32":
        os.startfile(tmp_path)
    return tmp_path


class FileListDialog(tk.Toplevel):
    """参照ファイルの複数選択・追加・削除・個別更新日確認ダイアログ"""

    def __init__(self, parent, title_label: str, current_paths: list, filetypes: list, on_apply):
        super().__init__(parent)
        self.title(f"{title_label} のファイル一覧管理")
        self.geometry("760x430")
        self.minsize(620, 340)
        self.configure(bg=COLORS["bg_card"])
        self.transient(parent)
        self.grab_set()

        self.filetypes = filetypes
        self.on_apply = on_apply
        self.paths = list(current_paths)

        frame = ttk.Frame(self, style="Card.TFrame", padding=14)
        frame.pack(fill=tk.BOTH, expand=True)

        btn_bar = ttk.Frame(frame, style="Card.TFrame")
        btn_bar.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))

        ttk.Button(btn_bar, text="ファイルを追加(複数選択可)", style="Primary.TButton", command=self._add_files).pack(side=tk.LEFT)
        ttk.Button(btn_bar, text="選択行を削除", command=self._remove_selected).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_bar, text="すべてクリア", command=self._clear_all).pack(side=tk.LEFT)

        ttk.Button(btn_bar, text="キャンセル", command=self.destroy).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(btn_bar, text="OK（反映）", style="Success.TButton", command=self._ok).pack(side=tk.RIGHT)

        ttk.Label(
            frame,
            text=f"{title_label}：複数ファイルを追加・管理できます（集計時に全ファイルが統合・合算されます）",
            style="Title.TLabel",
        ).pack(anchor=tk.W, pady=(0, 8))

        tv_frame = ttk.Frame(frame, style="Card.TFrame")
        tv_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        cols = ("ファイル名", "更新日時", "状態", "フルパス")
        self.tv = ttk.Treeview(tv_frame, columns=cols, show="headings", selectmode="extended", height=6)
        vsb = ttk.Scrollbar(tv_frame, orient=tk.VERTICAL, command=self.tv.yview)
        hsb = ttk.Scrollbar(tv_frame, orient=tk.HORIZONTAL, command=self.tv.xview)
        self.tv.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tv.grid(row=0, column=0, sticky=tk.NSEW)
        vsb.grid(row=0, column=1, sticky=tk.NS)
        hsb.grid(row=1, column=0, sticky=tk.EW)
        tv_frame.rowconfigure(0, weight=1)
        tv_frame.columnconfigure(0, weight=1)

        widths = [180, 155, 70, 310]
        for c, w in zip(cols, widths):
            self.tv.heading(c, text=c)
            self.tv.column(c, width=w, minwidth=50, anchor=tk.W)

        self._refresh_list()
        self.bind("<Escape>", lambda e: self.destroy())

    def _refresh_list(self):
        self.tv.delete(*self.tv.get_children())
        for i, p in enumerate(self.paths):
            exists = os.path.exists(p)
            mtime = format_file_mtime(p) if exists else "-"
            status = "OK" if exists else "不在"
            self.tv.insert("", tk.END, iid=str(i), values=(os.path.basename(p), mtime, status, p))

    def _add_files(self):
        init_dir = os.path.dirname(self.paths[-1]) if self.paths and os.path.exists(os.path.dirname(self.paths[-1])) else get_base_dir()
        selected = filedialog.askopenfilenames(
            parent=self,
            title="追加するファイルを選択してください（Ctrl/Shiftで複数選択可）",
            initialdir=init_dir,
            filetypes=self.filetypes,
        )
        if not selected:
            return
        existing_norm = {os.path.normcase(os.path.abspath(p)) for p in self.paths}
        for p in selected:
            norm = os.path.normpath(p)
            key = os.path.normcase(os.path.abspath(norm))
            if key not in existing_norm:
                self.paths.append(norm)
                existing_norm.add(key)
        self._refresh_list()

    def _remove_selected(self):
        sel = self.tv.selection()
        if not sel:
            return
        indices = sorted({int(i) for i in sel}, reverse=True)
        for idx in indices:
            if 0 <= idx < len(self.paths):
                self.paths.pop(idx)
        self._refresh_list()

    def _clear_all(self):
        self.paths.clear()
        self._refresh_list()

    def _ok(self):
        self.on_apply(self.paths)
        self.destroy()


class SeizoFolderDialog(tk.Toplevel):
    """制作伝票フォルダの複数選択・管理および検出済みExcel一覧ダイアログ（各フォルダ直下のみ参照）"""

    def __init__(self, parent, current_folders: list, on_apply, initial_tab: int = 0):
        super().__init__(parent)
        self.title("制作伝票フォルダ 複数選択・管理（各フォルダ直下のExcelのみ参照）")
        self.geometry("920x660")
        self.minsize(740, 520)
        self.configure(bg=COLORS["bg_card"])
        self.transient(parent)
        self.grab_set()

        self.on_apply = on_apply
        self.selected_folders = []
        self._selected_key_set = set()
        for p in (current_folders or []):
            self._add_folder_internal(p)

        # 初期表示する基準フォルダを決定（既存選択があればその親フォルダ、なければアプリ基準フォルダ）
        init_base = get_base_dir()
        for p in self.selected_folders:
            if os.path.isdir(p):
                parent_dir = os.path.dirname(os.path.abspath(p))
                if parent_dir and os.path.isdir(parent_dir):
                    init_base = parent_dir
                else:
                    init_base = os.path.abspath(p)
                break
        self.base_dir_var = tk.StringVar(value=init_base)
        self.filter_var = tk.StringVar(value="")
        self.summary_var = tk.StringVar(value="")
        self._browser_rows = []  # list of (full_path, disp_name, excel_count, is_self)

        frame = ttk.Frame(self, style="Card.TFrame", padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        # 最下部フッターを最優先で固定配置（見切れ防止）
        ftr = ttk.Frame(frame, style="Flat.TFrame")
        ftr.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))
        ttk.Label(ftr, textvariable=self.summary_var, style="Mtime.TLabel").pack(side=tk.LEFT)
        ttk.Button(ftr, text="保存して適用", style="Primary.TButton", command=self._apply).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(ftr, text="キャンセル", command=self._close).pack(side=tk.RIGHT)

        # 上部：親フォルダから複数フォルダをチェック選択するブラウザエリア
        browser_box = ttk.LabelFrame(
            frame,
            text="フォルダ一覧から複数チェック選択（クリックでON/OFF・ダブルクリックで中へ移動・各フォルダ直下のExcelのみ対象）",
            padding=(8, 6),
        )
        browser_box.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 8))

        nav_row = ttk.Frame(browser_box, style="Flat.TFrame")
        nav_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(nav_row, text="基準フォルダ:", style="Card.TLabel").pack(side=tk.LEFT, padx=(0, 4))
        ent_base = ttk.Entry(nav_row, textvariable=self.base_dir_var, state="readonly")
        ent_base.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Button(nav_row, text="基準フォルダを変更...", command=self._choose_base_dir).pack(side=tk.LEFT, padx=2)
        ttk.Button(nav_row, text="↑ 上の階層へ", command=self._go_parent_dir).pack(side=tk.LEFT, padx=2)

        act_row = ttk.Frame(browser_box, style="Flat.TFrame")
        act_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(act_row, text="フォルダ名絞り込み:", style="Card.TLabel").pack(side=tk.LEFT, padx=(0, 4))
        ent_flt = ttk.Entry(act_row, textvariable=self.filter_var, width=24)
        ent_flt.pack(side=tk.LEFT, padx=(0, 6))
        ent_flt.bind("<KeyRelease>", lambda e: self._refresh_browser_tree())
        ttk.Button(act_row, text="クリア", width=6, command=self._clear_filter).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(act_row, text="Excelありフォルダを全チェック", command=self._check_all_with_excel).pack(side=tk.LEFT, padx=2)
        ttk.Button(act_row, text="表示中を全チェック", command=self._check_all_visible).pack(side=tk.LEFT, padx=2)
        ttk.Button(act_row, text="表示中のチェック解除", command=self._uncheck_all_visible).pack(side=tk.LEFT, padx=2)

        b_tv_frame = ttk.Frame(browser_box, style="Card.TFrame")
        b_tv_frame.pack(fill=tk.BOTH, expand=True)
        self.browser_tv = ttk.Treeview(
            b_tv_frame,
            columns=("選択", "フォルダ名", "直下Excel数", "フルパス"),
            show="headings",
            height=8,
            selectmode="browse",
        )
        b_vsb = ttk.Scrollbar(b_tv_frame, orient=tk.VERTICAL, command=self.browser_tv.yview)
        b_hsb = ttk.Scrollbar(b_tv_frame, orient=tk.HORIZONTAL, command=self.browser_tv.xview)
        self.browser_tv.configure(yscrollcommand=b_vsb.set, xscrollcommand=b_hsb.set)
        self.browser_tv.grid(row=0, column=0, sticky=tk.NSEW)
        b_vsb.grid(row=0, column=1, sticky=tk.NS)
        b_hsb.grid(row=1, column=0, sticky=tk.EW)
        b_tv_frame.rowconfigure(0, weight=1)
        b_tv_frame.columnconfigure(0, weight=1)

        self.browser_tv.heading("選択", text="選択")
        self.browser_tv.heading("フォルダ名", text="フォルダ名")
        self.browser_tv.heading("直下Excel数", text="直下Excel数")
        self.browser_tv.heading("フルパス", text="フルパス")
        self.browser_tv.column("選択", width=52, anchor=tk.CENTER, stretch=False)
        self.browser_tv.column("フォルダ名", width=250, anchor=tk.W)
        self.browser_tv.column("直下Excel数", width=96, anchor=tk.E, stretch=False)
        self.browser_tv.column("フルパス", width=430, anchor=tk.W)

        self.browser_tv.bind("<ButtonRelease-1>", self._on_browser_click)
        self.browser_tv.bind("<Double-1>", self._on_browser_double_click)
        self.browser_tv.bind("<space>", self._on_browser_space)

        # 下部：選択済みフォルダ一覧 & 検出済みExcelファイル一覧のタブ
        self.nb = ttk.Notebook(frame)
        self.nb.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        tab_folders = ttk.Frame(self.nb, style="Card.TFrame", padding=6)
        tab_files = ttk.Frame(self.nb, style="Card.TFrame", padding=6)
        self.nb.add(tab_folders, text="選択中の制作伝票フォルダ (0件)")
        self.nb.add(tab_files, text="検出されるExcelファイル一覧 (0件)")

        # タブ1：選択済みフォルダ一覧
        f_btn_bar = ttk.Frame(tab_folders, style="Flat.TFrame")
        f_btn_bar.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(f_btn_bar, text="別フォルダを1つ追加...", command=self._add_one_folder_dialog).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(f_btn_bar, text="選択フォルダを削除", command=self._remove_selected_folders).pack(side=tk.LEFT, padx=4)
        ttk.Button(f_btn_bar, text="すべてクリア", command=self._clear_all_folders).pack(side=tk.LEFT, padx=4)

        sel_tv_frame = ttk.Frame(tab_folders, style="Card.TFrame")
        sel_tv_frame.pack(fill=tk.BOTH, expand=True)
        self.sel_tv = ttk.Treeview(
            sel_tv_frame,
            columns=("フォルダ名", "直下Excel数", "最新更新日時", "フルパス"),
            show="headings",
            height=6,
            selectmode="extended",
        )
        s_vsb = ttk.Scrollbar(sel_tv_frame, orient=tk.VERTICAL, command=self.sel_tv.yview)
        s_hsb = ttk.Scrollbar(sel_tv_frame, orient=tk.HORIZONTAL, command=self.sel_tv.xview)
        self.sel_tv.configure(yscrollcommand=s_vsb.set, xscrollcommand=s_hsb.set)
        self.sel_tv.grid(row=0, column=0, sticky=tk.NSEW)
        s_vsb.grid(row=0, column=1, sticky=tk.NS)
        s_hsb.grid(row=1, column=0, sticky=tk.EW)
        sel_tv_frame.rowconfigure(0, weight=1)
        sel_tv_frame.columnconfigure(0, weight=1)

        self.sel_tv.heading("フォルダ名", text="フォルダ名")
        self.sel_tv.heading("直下Excel数", text="直下Excel数")
        self.sel_tv.heading("最新更新日時", text="最新更新日時")
        self.sel_tv.heading("フルパス", text="フルパス")
        self.sel_tv.column("フォルダ名", width=180, anchor=tk.W)
        self.sel_tv.column("直下Excel数", width=92, anchor=tk.E, stretch=False)
        self.sel_tv.column("最新更新日時", width=150, anchor=tk.CENTER, stretch=False)
        self.sel_tv.column("フルパス", width=410, anchor=tk.W)

        # タブ2：検出されるExcelファイル一覧
        ef_tv_frame = ttk.Frame(tab_files, style="Card.TFrame")
        ef_tv_frame.pack(fill=tk.BOTH, expand=True)
        self.files_tv = ttk.Treeview(
            ef_tv_frame,
            columns=("ファイル名", "所属フォルダ", "更新日時", "フルパス"),
            show="headings",
            height=7,
        )
        e_vsb = ttk.Scrollbar(ef_tv_frame, orient=tk.VERTICAL, command=self.files_tv.yview)
        e_hsb = ttk.Scrollbar(ef_tv_frame, orient=tk.HORIZONTAL, command=self.files_tv.xview)
        self.files_tv.configure(yscrollcommand=e_vsb.set, xscrollcommand=e_hsb.set)
        self.files_tv.grid(row=0, column=0, sticky=tk.NSEW)
        e_vsb.grid(row=0, column=1, sticky=tk.NS)
        e_hsb.grid(row=1, column=0, sticky=tk.EW)
        ef_tv_frame.rowconfigure(0, weight=1)
        ef_tv_frame.columnconfigure(0, weight=1)

        self.files_tv.heading("ファイル名", text="ファイル名")
        self.files_tv.heading("所属フォルダ", text="所属フォルダ")
        self.files_tv.heading("更新日時", text="更新日時")
        self.files_tv.heading("フルパス", text="フルパス")
        self.files_tv.column("ファイル名", width=180, anchor=tk.W)
        self.files_tv.column("所属フォルダ", width=170, anchor=tk.W)
        self.files_tv.column("更新日時", width=150, anchor=tk.CENTER, stretch=False)
        self.files_tv.column("フルパス", width=340, anchor=tk.W)

        self._folder_entries_cache = {}
        self._files_render_token = 0
        self._files_chunk_job = None
        self._scan_token = 0
        self._scan_after_id = None
        self._scan_queue = queue.Queue()
        self._dialog_closed = False

        self.protocol("WM_DELETE_WINDOW", self._close)
        self._scan_base_dir_async()
        self._refresh_selected_views()
        if initial_tab in (0, 1):
            self.nb.select(initial_tab)
        self.bind("<Escape>", lambda e: self._close())

    def _close(self):
        self._dialog_closed = True
        self._scan_token += 1
        self._files_render_token += 1
        for attr in ("_files_chunk_job", "_scan_after_id"):
            job = getattr(self, attr, None)
            if job is not None:
                try:
                    self.after_cancel(job)
                except Exception:
                    pass
                setattr(self, attr, None)
        try:
            self.destroy()
        except Exception:
            pass

    @staticmethod
    def _norm_key(path: str) -> str:
        return os.path.normcase(os.path.abspath(os.path.normpath(path)))

    def _add_folder_internal(self, path: str) -> bool:
        if not path:
            return False
        norm = os.path.normpath(path)
        key = self._norm_key(norm)
        if key in self._selected_key_set:
            return False
        self._selected_key_set.add(key)
        self.selected_folders.append(norm)
        return True

    def _remove_folder_internal(self, path: str) -> bool:
        key = self._norm_key(path)
        if key not in self._selected_key_set:
            return False
        self._selected_key_set.remove(key)
        self.selected_folders = [p for p in self.selected_folders if self._norm_key(p) != key]
        return True

    @staticmethod
    def _count_direct_excels(folder_path: str) -> int:
        if not folder_path or not os.path.isdir(folder_path):
            return 0
        exts = (".xls", ".xlsx", ".xlsm")
        cnt = 0
        try:
            for entry in os.scandir(folder_path):
                try:
                    if entry.is_file() and not entry.name.startswith("~$") and entry.name.lower().endswith(exts):
                        cnt += 1
                except Exception:
                    pass
        except Exception:
            pass
        return cnt

    def _choose_base_dir(self):
        cur = self.base_dir_var.get()
        init_dir = cur if cur and os.path.isdir(cur) else get_base_dir()
        folder = filedialog.askdirectory(
            parent=self,
            title="制作伝票フォルダが入っている基準（親）フォルダを選択してください",
            initialdir=init_dir,
        )
        if folder:
            self.base_dir_var.set(os.path.normpath(folder))
            self._scan_base_dir_async()

    def _go_parent_dir(self):
        cur = self.base_dir_var.get()
        if not cur:
            return
        parent_dir = os.path.dirname(os.path.abspath(cur))
        if parent_dir and parent_dir != os.path.abspath(cur) and os.path.isdir(parent_dir):
            self.base_dir_var.set(os.path.normpath(parent_dir))
            self._scan_base_dir_async()

    def _clear_filter(self):
        self.filter_var.set("")
        self._refresh_browser_tree()

    def _collect_base_dir_rows(self, base: str, token: int = None) -> list:
        rows = []
        if base and os.path.isdir(base):
            base_norm = os.path.normpath(base)
            rows.append((base_norm, f"[このフォルダ自体]  {os.path.basename(base_norm) or base_norm}", self._count_direct_excels(base_norm), True))
            sub_entries = []
            try:
                with os.scandir(base_norm) as it:
                    for entry in it:
                        if token is not None and (self._dialog_closed or token != self._scan_token):
                            return []
                        try:
                            if entry.is_dir() and not entry.name.startswith("."):
                                sub_entries.append((entry.name.lower(), entry.name, os.path.normpath(entry.path)))
                        except Exception:
                            pass
            except Exception:
                pass
            sub_entries.sort(key=lambda t: t[0])
            for _, disp_name, p_norm in sub_entries:
                if token is not None and (self._dialog_closed or token != self._scan_token):
                    return []
                rows.append((p_norm, disp_name, self._count_direct_excels(p_norm), False))
        return rows

    def _scan_base_dir(self):
        """基準フォルダ直下のサブフォルダ一覧とExcelファイル数を同期スキャンして表示を更新する"""
        self._scan_token += 1
        base = self.base_dir_var.get()
        self._browser_rows = self._collect_base_dir_rows(base)
        self._refresh_browser_tree()

    def _scan_base_dir_async(self):
        """ネットワーク／クラウドドライブ上の大量サブフォルダでもUIを固めずにスレッド安全に非同期スキャンする"""
        self._scan_token += 1
        token = self._scan_token
        if self._scan_after_id is not None:
            try:
                self.after_cancel(self._scan_after_id)
            except Exception:
                pass
            self._scan_after_id = None
        base = self.base_dir_var.get()
        self._browser_rows = []
        try:
            self.browser_tv.delete(*self.browser_tv.get_children())
            self.browser_tv.insert("", tk.END, iid="_loading", values=("…", "フォルダを読み込み中...", "-", base))
        except Exception:
            pass

        def worker():
            rows = self._collect_base_dir_rows(base, token=token)
            if self._dialog_closed or token != self._scan_token:
                return
            self._scan_queue.put((token, rows))

        threading.Thread(target=worker, daemon=True).start()
        if not self._dialog_closed:
            self._scan_after_id = self.after(15, self._poll_scan_queue)

    def _poll_scan_queue(self):
        self._scan_after_id = None
        if self._dialog_closed:
            return
        applied_current = False
        try:
            while True:
                token, rows = self._scan_queue.get_nowait()
                if token == self._scan_token:
                    self._apply_scanned_rows(rows, token)
                    applied_current = True
        except queue.Empty:
            pass
        if not applied_current and not self._dialog_closed:
            try:
                if self.winfo_exists():
                    self._scan_after_id = self.after(20, self._poll_scan_queue)
            except Exception:
                pass

    def _apply_scanned_rows(self, rows: list, token: int):
        if self._dialog_closed or token != self._scan_token:
            return
        try:
            if not self.browser_tv.winfo_exists():
                return
        except Exception:
            return
        self._browser_rows = rows
        self._refresh_browser_tree()

    def _get_filtered_browser_rows(self) -> list:
        kw = normalize_search_text(self.filter_var.get())
        if not kw:
            return list(self._browser_rows)
        tokens = [t for t in kw.split() if t]
        return [
            row for row in self._browser_rows
            if all(t in normalize_search_text(row[1]) for t in tokens)
        ]

    def _refresh_browser_tree(self):
        self.browser_tv.delete(*self.browser_tv.get_children())
        for idx, (fpath, disp, cnt, _is_self) in enumerate(self._get_filtered_browser_rows()):
            mark = "☑" if self._norm_key(fpath) in self._selected_key_set else "☐"
            self.browser_tv.insert("", tk.END, iid=str(idx), values=(mark, disp, f"{cnt:,}件", fpath))

    def _toggle_browser_iid(self, iid: str):
        if not iid or iid == "_loading":
            return
        vals = self.browser_tv.item(iid, "values")
        if not vals or len(vals) < 4:
            return
        fpath = vals[3]
        if self._norm_key(fpath) in self._selected_key_set:
            self._remove_folder_internal(fpath)
            new_mark = "☐"
        else:
            self._add_folder_internal(fpath)
            new_mark = "☑"
        self.browser_tv.item(iid, values=(new_mark, vals[1], vals[2], fpath))
        self._refresh_selected_views()

    def _on_browser_click(self, event):
        region = self.browser_tv.identify_region(event.x, event.y)
        if region not in ("cell", "tree"):
            return
        iid = self.browser_tv.identify_row(event.y)
        if iid and iid != "_loading":
            self._toggle_browser_iid(iid)

    def _on_browser_double_click(self, event):
        iid = self.browser_tv.identify_row(event.y)
        if not iid or iid == "_loading":
            return
        vals = self.browser_tv.item(iid, "values")
        if not vals or len(vals) < 4:
            return
        fpath = vals[3]
        # ダブルクリックでそのフォルダの中へ移動（直前の単クリックによるチェック切替を戻してから移動）
        if os.path.isdir(fpath) and self._norm_key(fpath) != self._norm_key(self.base_dir_var.get()):
            self._toggle_browser_iid(iid)
            self.base_dir_var.set(os.path.normpath(fpath))
            self.filter_var.set("")
            self._scan_base_dir_async()

    def _on_browser_space(self, event):
        sel = self.browser_tv.selection()
        if sel:
            self._toggle_browser_iid(sel[0])
        return "break"

    def _check_all_with_excel(self):
        for fpath, _disp, cnt, _is_self in self._get_filtered_browser_rows():
            if cnt > 0:
                self._add_folder_internal(fpath)
        self._refresh_browser_tree()
        self._refresh_selected_views()

    def _check_all_visible(self):
        for fpath, _disp, _cnt, _is_self in self._get_filtered_browser_rows():
            self._add_folder_internal(fpath)
        self._refresh_browser_tree()
        self._refresh_selected_views()

    def _uncheck_all_visible(self):
        for fpath, _disp, _cnt, _is_self in self._get_filtered_browser_rows():
            self._remove_folder_internal(fpath)
        self._refresh_browser_tree()
        self._refresh_selected_views()

    def _add_one_folder_dialog(self):
        cur = self.base_dir_var.get()
        init_dir = cur if cur and os.path.isdir(cur) else get_base_dir()
        folder = filedialog.askdirectory(
            parent=self,
            title="追加する制作伝票フォルダを選択してください（直下のExcelのみ参照）",
            initialdir=init_dir,
        )
        if folder:
            self._add_folder_internal(folder)
            self._refresh_browser_tree()
            self._refresh_selected_views()

    def _remove_selected_folders(self):
        sel = self.sel_tv.selection()
        if not sel:
            return
        for iid in sel:
            vals = self.sel_tv.item(iid, "values")
            if vals and len(vals) >= 4:
                self._remove_folder_internal(vals[3])
        self._refresh_browser_tree()
        self._refresh_selected_views()

    def _clear_all_folders(self):
        self.selected_folders = []
        self._selected_key_set = set()
        self._refresh_browser_tree()
        self._refresh_selected_views()

    def _get_cached_folder_entries(self, folder_path: str) -> list:
        key = self._norm_key(folder_path)
        if key not in self._folder_entries_cache:
            self._folder_entries_cache[key] = scan_seizo_excel_entries([folder_path])
        return self._folder_entries_cache[key]

    def _refresh_selected_views(self):
        if self._files_chunk_job is not None:
            try:
                self.after_cancel(self._files_chunk_job)
            except Exception:
                pass
            self._files_chunk_job = None

        # タブ1：選択済みフォルダ一覧（1パスキャッシュ結果を使い二重statを排除）
        self.sel_tv.delete(*self.sel_tv.get_children())
        all_entries = []
        seen_keys = set()
        for fp in self.selected_folders:
            entries = self._get_cached_folder_entries(fp)
            if entries:
                latest_ts = max(e[3] for e in entries)
                latest_str = datetime.datetime.fromtimestamp(latest_ts).strftime("%Y/%m/%d %H:%M:%S")
                for e in entries:
                    ek = self._norm_key(e[0])
                    if ek not in seen_keys:
                        seen_keys.add(ek)
                        all_entries.append(e)
            else:
                latest_str = "対象Excelなし" if os.path.isdir(fp) else "フォルダなし"
            self.sel_tv.insert("", tk.END, values=(os.path.basename(fp) or fp, f"{len(entries):,}件", latest_str, fp))

        # タブ2：全選択フォルダ直下の検出済みExcelファイル一覧（大量ファイルでもフリーズしないチャンク描画）
        self._files_render_token += 1
        token = self._files_render_token
        self.files_tv.delete(*self.files_tv.get_children())
        self._render_files_chunk(all_entries, 0, 200, token)

        self.nb.tab(0, text=f"選択中の制作伝票フォルダ ({len(self.selected_folders):,}件)")
        self.nb.tab(1, text=f"検出されるExcelファイル一覧 ({len(all_entries):,}件)")
        self.summary_var.set(
            f"選択中: {len(self.selected_folders):,} フォルダ / 検出Excel: {len(all_entries):,} ファイル（各フォルダ直下のみ・すべて読み取り専用で参照）"
        )

    def _render_files_chunk(self, all_entries: list, start_idx: int, chunk_size: int, token: int):
        self._files_chunk_job = None
        if self._dialog_closed or token != self._files_render_token:
            return
        try:
            if not self.files_tv.winfo_exists():
                return
        except Exception:
            return
        end_idx = min(start_idx + chunk_size, len(all_entries))
        tv_insert = self.files_tv.insert
        end_pos = tk.END
        from_ts = datetime.datetime.fromtimestamp
        for idx in range(start_idx, end_idx):
            ef, _sz, _mns, m_sec = all_entries[idx]
            parent_name = os.path.basename(os.path.dirname(ef)) or os.path.dirname(ef)
            mtime_str = from_ts(m_sec).strftime("%Y/%m/%d %H:%M:%S")
            tv_insert("", end_pos, values=(os.path.basename(ef), parent_name, mtime_str, ef))
        if end_idx < len(all_entries) and not self._dialog_closed:
            self._files_chunk_job = self.after(1, lambda: self._render_files_chunk(all_entries, end_idx, 600, token))

    def _apply(self):
        self.on_apply(list(self.selected_folders))
        self._close()


class ColumnOrderDialog(tk.Toplevel):
    """一覧の列の順番・表示設定ダイアログ"""

    def __init__(self, parent, current_order: list, visible_set: set, on_apply, display_unit: str = DEFAULT_TIME_UNIT):
        super().__init__(parent)
        self.title("列の順番・表示設定")
        self.geometry("520x560")
        self.minsize(460, 480)
        self.configure(bg=COLORS["bg_card"])
        self.transient(parent)
        self.grab_set()

        self.on_apply = on_apply
        self.display_unit = display_unit if display_unit in VALID_TIME_UNITS else DEFAULT_TIME_UNIT

        # 全利用可能列を順序リストに揃える
        seen = set()
        self.col_items = []
        for col in current_order:
            if col in ALL_AVAILABLE_COLUMNS and col not in seen:
                self.col_items.append([col, col in visible_set])
                seen.add(col)
        for col in ALL_AVAILABLE_COLUMNS:
            if col not in seen:
                self.col_items.append([col, col in visible_set])
                seen.add(col)

        frame = ttk.Frame(self, style="Card.TFrame", padding=14)
        frame.pack(fill=tk.BOTH, expand=True)

        footer = ttk.Frame(frame, style="Card.TFrame")
        footer.pack(side=tk.BOTTOM, fill=tk.X, pady=(12, 0))
        ttk.Button(footer, text="キャンセル", command=self.destroy).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(footer, text="保存して適用", style="Primary.TButton", command=self._ok).pack(side=tk.RIGHT)

        ttk.Label(frame, text="一覧に表示する列の順番と表示/非表示を設定します", style="Title.TLabel").pack(anchor=tk.W)
        ttk.Label(
            frame,
            text="列を選択して「上へ」「下へ」で並び替え、ダブルクリックまたは「表示切替」でON/OFFできます。",
            style="Mtime.TLabel",
        ).pack(anchor=tk.W, pady=(2, 10))

        body = ttk.Frame(frame, style="Card.TFrame")
        body.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        tv_frame = ttk.Frame(body, style="Card.TFrame")
        tv_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.tv = ttk.Treeview(tv_frame, columns=("状態", "列名", "種別"), show="headings", selectmode="browse", height=8)
        vsb = ttk.Scrollbar(tv_frame, orient=tk.VERTICAL, command=self.tv.yview)
        self.tv.configure(yscrollcommand=vsb.set)
        self.tv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tv.heading("状態", text="表示")
        self.tv.heading("列名", text="列名")
        self.tv.heading("種別", text="区分")
        self.tv.column("状態", width=55, anchor=tk.CENTER, stretch=False)
        self.tv.column("列名", width=180, anchor=tk.W)
        self.tv.column("種別", width=90, anchor=tk.CENTER, stretch=False)
        self.tv.bind("<Double-1>", lambda e: self._toggle_visible())
        self.tv.bind("<space>", lambda e: (self._toggle_visible(), "break")[1])

        ctrl_col = ttk.Frame(body, style="Card.TFrame", padding=(10, 0, 0, 0))
        ctrl_col.pack(side=tk.RIGHT, fill=tk.Y)

        ttk.Button(ctrl_col, text="▲▲ 一番上へ", width=14, command=lambda: self._move("top")).pack(pady=3)
        ttk.Button(ctrl_col, text="▲ 上へ", width=14, command=lambda: self._move("up")).pack(pady=3)
        ttk.Button(ctrl_col, text="▼ 下へ", width=14, command=lambda: self._move("down")).pack(pady=3)
        ttk.Button(ctrl_col, text="▼▼ 一番下へ", width=14, command=lambda: self._move("bottom")).pack(pady=3)
        ttk.Separator(ctrl_col, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        ttk.Button(ctrl_col, text="表示 / 非表示 切替", width=14, command=self._toggle_visible).pack(pady=3)
        ttk.Button(ctrl_col, text="すべて表示", width=14, command=self._show_all).pack(pady=3)
        ttk.Separator(ctrl_col, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        ttk.Button(ctrl_col, text="デフォルト順に戻す", width=14, command=self._reset_default).pack(pady=3)

        self._refresh(select_idx=0)
        self.bind("<Escape>", lambda e: self.destroy())

    def _refresh(self, select_idx: int = None):
        self.tv.delete(*self.tv.get_children())
        for i, (col, vis) in enumerate(self.col_items):
            kind = "集計・小計" if col in CALC_COLUMNS else "基本情報"
            mark = "☑ 表示" if vis else "☐ 非表示"
            disp_name = get_display_column_label(col, self.display_unit)
            self.tv.insert("", tk.END, iid=str(i), values=(mark, disp_name, kind))
        if select_idx is not None and 0 <= select_idx < len(self.col_items):
            self.tv.selection_set(str(select_idx))
            self.tv.see(str(select_idx))

    def _get_sel_index(self):
        sel = self.tv.selection()
        if not sel:
            return None
        return int(sel[0])

    def _move(self, direction: str):
        idx = self._get_sel_index()
        if idx is None:
            return
        new_idx = idx
        item = self.col_items.pop(idx)
        if direction == "up":
            new_idx = max(0, idx - 1)
        elif direction == "down":
            new_idx = min(len(self.col_items), idx + 1)
        elif direction == "top":
            new_idx = 0
        elif direction == "bottom":
            new_idx = len(self.col_items)
        self.col_items.insert(new_idx, item)
        self._refresh(select_idx=new_idx)

    def _toggle_visible(self):
        idx = self._get_sel_index()
        if idx is None:
            return
        self.col_items[idx][1] = not self.col_items[idx][1]
        self._refresh(select_idx=idx)

    def _show_all(self):
        idx = self._get_sel_index()
        for item in self.col_items:
            item[1] = True
        self._refresh(select_idx=idx)

    def _reset_default(self):
        self.col_items = [[c, True] for c in DEFAULT_COLUMN_ORDER] + [[c, False] for c in ALL_AVAILABLE_COLUMNS if c not in DEFAULT_COLUMN_ORDER]
        self._refresh(select_idx=0)

    def _ok(self):
        visible_cols = [c for c, vis in self.col_items if vis]
        if not visible_cols:
            messagebox.showwarning("確認", "少なくとも1列以上を表示状態にしてください。", parent=self)
            return
        full_order = [c for c, _ in self.col_items]
        self.on_apply(full_order, set(visible_cols))
        self.destroy()


class ExcelFilterPopup(tk.Toplevel):
    """
    Excelのオートフィルタのような直感的かつ高速な列フィルター＆ソートポップアップ
    - 下部アクションバー（OK/キャンセル/解除）を先に配置し、ボタン切れを完全防止
    - 日本語IME変換確定のEnterキーで画面が勝手に閉じないよう制御
    - 半角カナ/全角カナ・全角英数・カンマ有無・スペース区切り複数語に対応した検索
    - 1万件超のユニーク値でも固まらないチャンク描画とドラッグ連続チェックに対応
    """

    def __init__(
        self,
        parent,
        col_name: str,
        x_root: int,
        y_root: int,
        value_counts: Counter,
        is_numeric: bool,
        current_filter: dict,
        on_sort,
        on_apply_filter,
        on_clear_filter,
        display_col_name: str = "",
    ):
        super().__init__(parent)
        self.col_name = col_name
        self.display_col_name = display_col_name or col_name
        self.title(f"フィルター: {self.display_col_name}")
        self.configure(bg=COLORS["bg_card"])
        self.transient(parent)
        self.grab_set()

        self.is_numeric = is_numeric
        self.value_counts = value_counts
        self.on_sort = on_sort
        self.on_apply_filter = on_apply_filter
        self.on_clear_filter = on_clear_filter

        self._search_after_id = None
        self._chunk_after_id = None
        self._render_token = 0
        self._last_norm_kw = None
        self._drag_check_target = None

        # ウィンドウ位置とサイズ調整（画面外はみ出し・ボタン切れ防止）
        w = 430
        h = 560 if is_numeric else 500
        self.minsize(390, 440)
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        px = min(max(10, x_root), max(10, sw - w - 20))
        py = min(max(10, y_root), max(10, sh - h - 60))
        self.geometry(f"{w}x{h}+{px}+{py}")
        self.resizable(True, True)

        main = ttk.Frame(self, style="Card.TFrame", padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        # 1. 最下部フッターボタン（先に side=BOTTOM で配置し、どの解像度でも絶対に途切れないようにする）
        footer = tk.Frame(main, bg=COLORS["bg_header"], padx=8, pady=6)
        footer.pack(side=tk.BOTTOM, fill=tk.X, pady=(8, 0))
        ttk.Button(footer, text="フィルター解除", command=self._do_clear).pack(side=tk.LEFT)
        ttk.Button(footer, text="キャンセル", command=self._close).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(footer, text="OK（適用）", style="Primary.TButton", command=self._do_ok).pack(side=tk.RIGHT)

        # 2. ヘッダー＆ソートボタン
        ttk.Label(main, text=f"列「{self.display_col_name}」の並べ替え・絞り込み", style="Title.TLabel").pack(anchor=tk.W, pady=(0, 5))

        sort_f = ttk.Frame(main, style="Card.TFrame")
        sort_f.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(sort_f, text="↑ 昇順ソート", command=self._do_sort_asc).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 3))
        ttk.Button(sort_f, text="↓ 降順ソート", command=self._do_sort_desc).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(3, 0))

        # 3. 数値条件フィルター（数値列の場合）
        self.num_op_var = tk.StringVar(value=(current_filter or {}).get("num_op", "指定なし"))
        init_v1 = (current_filter or {}).get("num_v1")
        init_v2 = (current_filter or {}).get("num_v2")
        self.num_v1_var = tk.StringVar(value="" if init_v1 is None else str(init_v1))
        self.num_v2_var = tk.StringVar(value="" if init_v2 is None else str(init_v2))

        if is_numeric:
            num_box = ttk.LabelFrame(main, text="数値条件フィルター", padding=(6, 4))
            num_box.pack(fill=tk.X, pady=(0, 6))

            row1 = ttk.Frame(num_box, style="Card.TFrame")
            row1.pack(fill=tk.X, pady=2)
            op_cb = ttk.Combobox(
                row1,
                textvariable=self.num_op_var,
                values=["指定なし", "0より大きい (>0)", "0のみ (=0)", "以上 (>=)", "以下 (<=)", "より大きい (>)", "未満 (<)", "範囲 (最小～最大)"],
                state="readonly",
                width=18,
            )
            op_cb.pack(side=tk.LEFT, padx=(0, 4))
            op_cb.bind("<<ComboboxSelected>>", lambda e: self._toggle_num_inputs())

            self.ent_v1 = ttk.Entry(row1, textvariable=self.num_v1_var, width=10)
            self.ent_v1.pack(side=tk.LEFT, padx=2)
            self.ent_v1.bind("<Return>", lambda e: self._do_ok())
            self.lbl_tilde = ttk.Label(row1, text="～", style="Card.TLabel")
            self.ent_v2 = ttk.Entry(row1, textvariable=self.num_v2_var, width=10)
            self.ent_v2.bind("<Return>", lambda e: self._do_ok())
            self._toggle_num_inputs()

        # 4. 値検索ボックス（IME変換確定Enterで閉じない安全設計＋ワンクリック検索操作）
        search_box = ttk.LabelFrame(main, text="値リストから検索・選択（半角/全角・スペース区切り対応）", padding=(6, 4))
        search_box.pack(fill=tk.X, pady=(0, 4))

        search_f = ttk.Frame(search_box, style="Card.TFrame")
        search_f.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(search_f, text="検索:", style="Card.TLabel").pack(side=tk.LEFT, padx=(0, 4))
        self.search_var = tk.StringVar()
        self.ent_s = ttk.Entry(search_f, textvariable=self.search_var)
        self.ent_s.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        self.ent_s.bind("<KeyRelease>", self._on_search_key_release)
        self.ent_s.bind("<Return>", self._on_search_enter)
        self.ent_s.bind("<Down>", self._on_search_arrow_down)

        ttk.Button(search_f, text="クリア", width=6, command=self._clear_search_box).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Button(search_f, text="検索結果で即適用", style="Success.TButton", command=self._apply_search_matches_immediately).pack(side=tk.LEFT)

        # クイック選択ボタン行
        quick_row = ttk.Frame(search_box, style="Card.TFrame")
        quick_row.pack(fill=tk.X)

        self.all_selected_var = tk.BooleanVar(value=True)
        chk_all = tk.Checkbutton(
            quick_row,
            text="表示中を全選択",
            variable=self.all_selected_var,
            command=self._toggle_select_all,
            bg=COLORS["bg_card"],
            fg=COLORS["fg_main"],
            activebackground=COLORS["bg_card"],
            selectcolor="#FFFFFF",
            font=("Yu Gothic UI", 9, "bold"),
            anchor=tk.W,
        )
        chk_all.pack(side=tk.LEFT)

        ttk.Button(quick_row, text="表示中のみ選択", command=self._select_only_visible).pack(side=tk.RIGHT, padx=(3, 0))
        ttk.Button(quick_row, text="全解除", width=7, command=self._uncheck_all_values).pack(side=tk.RIGHT, padx=(3, 0))
        ttk.Button(quick_row, text="全選択", width=7, command=self._check_all_values).pack(side=tk.RIGHT)

        # 件数ステータス表示
        self.match_info_var = tk.StringVar(value="")
        tk.Label(
            main,
            textvariable=self.match_info_var,
            bg=COLORS["bg_card"],
            fg=COLORS["accent"],
            font=("Yu Gothic UI", 9, "bold"),
            anchor=tk.W,
        ).pack(fill=tk.X, pady=(0, 2))

        # 5. 値チェックリスト（高速Treeview）
        tv_frame = ttk.Frame(main, style="Card.TFrame")
        tv_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.tv = ttk.Treeview(tv_frame, columns=("chk", "val", "cnt"), show="headings", selectmode="browse", height=5)
        vsb = ttk.Scrollbar(tv_frame, orient=tk.VERTICAL, command=self.tv.yview)
        self.tv.configure(yscrollcommand=vsb.set)
        self.tv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tv.heading("chk", text="選択")
        self.tv.heading("val", text="値（クリック/ドラッグで切替）")
        self.tv.heading("cnt", text="件数")
        self.tv.column("chk", width=46, anchor=tk.CENTER, stretch=False)
        self.tv.column("val", width=230, anchor=tk.E if is_numeric else tk.W)
        self.tv.column("cnt", width=65, anchor=tk.E, stretch=False)

        self.tv.bind("<Button-1>", self._on_tv_click)
        self.tv.bind("<B1-Motion>", self._on_tv_drag)
        self.tv.bind("<ButtonRelease-1>", lambda e: setattr(self, "_drag_check_target", None))
        self.tv.bind("<space>", self._on_tv_space)
        self.tv.bind("<Return>", lambda e: self._do_ok())

        # 全ユニーク値のソートと検索用キャッシュの事前構築
        if is_numeric:
            self.all_values = sorted(
                value_counts.keys(),
                key=lambda v: (v == "", to_float(v, 0.0)),
            )
        else:
            self.all_values = sorted(value_counts.keys(), key=lambda v: (v == "", str(v)))

        self._disp_cache = {}
        self._norm_search_cache = {}
        for v in self.all_values:
            disp_s = self._format_display_val(v)
            self._disp_cache[v] = disp_s
            norm_s = normalize_search_text(disp_s)
            norm_raw = normalize_search_text(v)
            self._norm_search_cache[v] = f"{norm_s} {norm_raw} {norm_s.replace(',', '')}"

        selected_set = (current_filter or {}).get("selected")
        if selected_set is None:
            self.checked_map = {v: True for v in self.all_values}
            self._initial_all_checked = True
        else:
            self.checked_map = {v: (v in selected_set) for v in self.all_values}
            self._initial_all_checked = False
        self._user_toggled_items = False

        self._popup_closed = False
        self.protocol("WM_DELETE_WINDOW", self._close)
        self.visible_values = []
        self._rebuild_tree(from_search_typing=False)

        self.bind("<Escape>", lambda e: self._close())
        self.ent_s.focus_set()

    def _close(self):
        self._popup_closed = True
        self._render_token += 1
        for attr in ("_search_after_id", "_chunk_after_id"):
            job = getattr(self, attr, None)
            if job is not None:
                try:
                    self.after_cancel(job)
                except Exception:
                    pass
                setattr(self, attr, None)
        try:
            self.destroy()
        except Exception:
            pass

    def _toggle_num_inputs(self):
        if not self.is_numeric:
            return
        op = self.num_op_var.get()
        if op in ("指定なし", "0より大きい (>0)", "0のみ (=0)"):
            self.ent_v1.configure(state=tk.DISABLED)
            self.lbl_tilde.pack_forget()
            self.ent_v2.pack_forget()
        elif op == "範囲 (最小～最大)":
            self.ent_v1.configure(state=tk.NORMAL)
            self.lbl_tilde.pack(side=tk.LEFT, padx=2)
            self.ent_v2.pack(side=tk.LEFT, padx=2)
        else:
            self.ent_v1.configure(state=tk.NORMAL)
            self.lbl_tilde.pack_forget()
            self.ent_v2.pack_forget()

    def _format_display_val(self, raw_v: str) -> str:
        if raw_v == "":
            return "(空白)"
        if self.is_numeric:
            try:
                num = float(str(raw_v).replace(",", ""))
                if math.isfinite(num):
                    if abs(num - round(num)) < 1e-6:
                        return f"{int(round(num)):,}"
                    return f"{num:,.2f}"
            except ValueError:
                pass
        return raw_v

    def _on_search_key_release(self, event=None):
        """IME入力中も邪魔せずスムーズに候補を絞り込む（デバウンス付き）"""
        if event is not None and event.keysym in (
            "Return", "KP_Enter", "Escape", "Up", "Down", "Left", "Right",
            "Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R", "Tab",
        ):
            return
        norm_kw = normalize_search_text(self.search_var.get())
        if norm_kw == self._last_norm_kw:
            return
        if self._search_after_id is not None:
            try:
                self.after_cancel(self._search_after_id)
            except Exception:
                pass
            self._search_after_id = None
        if event is None:
            self._rebuild_tree(from_search_typing=True)
        else:
            self._search_after_id = self.after(50, lambda: self._rebuild_tree(from_search_typing=True))

    def _on_search_enter(self, event=None):
        """検索ボックスでのEnterキーはIME変換確定・検索更新のみ行い、ポップアップを閉じない"""
        if self._search_after_id is not None:
            try:
                self.after_cancel(self._search_after_id)
            except Exception:
                pass
            self._search_after_id = None
        self._rebuild_tree(from_search_typing=True)
        return "break"

    def _on_search_arrow_down(self, event=None):
        """検索ボックスで↓キーを押すと候補リストの先頭行へフォーカスを移す"""
        children = self.tv.get_children()
        if children:
            self.tv.focus_set()
            self.tv.selection_set(children[0])
            self.tv.focus(children[0])
            self.tv.see(children[0])
        return "break"

    def _clear_search_box(self):
        self.search_var.set("")
        self._rebuild_tree(from_search_typing=True)
        self.ent_s.focus_set()

    def _apply_search_matches_immediately(self):
        """現在の検索結果（表示中の候補）のみを選択状態にして即座にフィルター適用する"""
        if not self.visible_values:
            messagebox.showinfo("該当なし", "検索キーワードに一致する値がありません。", parent=self)
            return
        vis_set = set(self.visible_values)
        for v in self.all_values:
            self.checked_map[v] = (v in vis_set)
        self._user_toggled_items = True
        self._do_ok()

    def _update_match_info_label(self):
        vis_cnt = len(self.visible_values)
        total_cnt = len(self.all_values)
        vis_checked = sum(1 for v in self.visible_values if self.checked_map.get(v, False))
        total_checked = sum(1 for v in self.all_values if self.checked_map.get(v, False))
        self.all_selected_var.set(vis_cnt > 0 and vis_checked == vis_cnt)
        if vis_cnt < total_cnt:
            self.match_info_var.set(f"検索一致: {vis_cnt:,}件 / 全{total_cnt:,}件  （選択中: 表示内 {vis_checked:,}件 / 全体 {total_checked:,}件）")
        else:
            self.match_info_var.set(f"全 {total_cnt:,}件  （チェック選択中: {total_checked:,}件）")

    def _rebuild_tree(self, from_search_typing: bool = True):
        self._search_after_id = None
        if self._chunk_after_id is not None:
            try:
                self.after_cancel(self._chunk_after_id)
            except Exception:
                pass
            self._chunk_after_id = None

        norm_kw = normalize_search_text(self.search_var.get())
        self._last_norm_kw = norm_kw
        tokens = [t.replace(",", "") for t in norm_kw.split() if t]

        if not tokens:
            self.visible_values = list(self.all_values)
        else:
            norm_cache = self._norm_search_cache
            self.visible_values = [
                v for v in self.all_values
                if all(tok in norm_cache[v] for tok in tokens)
            ]

        # 初期全選択状態から検索語を入力した場合は、検索結果のみを自動的にチェック状態にする（Excel準拠）
        if from_search_typing and self._initial_all_checked and not self._user_toggled_items:
            if tokens:
                vis_set = set(self.visible_values)
                for v in self.all_values:
                    self.checked_map[v] = (v in vis_set)
            else:
                for v in self.all_values:
                    self.checked_map[v] = True

        self._update_match_info_label()

        self._render_token += 1
        token = self._render_token
        self.tv.delete(*self.tv.get_children())
        self._render_tree_chunk(0, 250, token)

    def _render_tree_chunk(self, start_idx: int, chunk_size: int, token: int):
        self._chunk_after_id = None
        if self._popup_closed or token != self._render_token:
            return
        try:
            if not self.tv.winfo_exists():
                return
        except Exception:
            return

        total_vis = len(self.visible_values)
        end_idx = min(start_idx + chunk_size, total_vis)
        tv_insert = self.tv.insert
        end_pos = tk.END
        checked_map = self.checked_map
        val_counts = self.value_counts
        disp_cache = self._disp_cache

        for idx in range(start_idx, end_idx):
            v = self.visible_values[idx]
            mark = "☑" if checked_map.get(v, False) else "☐"
            tv_insert("", end_pos, iid=str(idx), values=(mark, disp_cache[v], f"{val_counts[v]:,}"))

        if end_idx < total_vis and not self._popup_closed:
            self._chunk_after_id = self.after(1, lambda: self._render_tree_chunk(end_idx, 800, token))

    def _refresh_rendered_checkmarks(self):
        """現在Treeview上に描画されている行のチェック記号だけを一括高速更新する"""
        checked_map = self.checked_map
        val_counts = self.value_counts
        disp_cache = self._disp_cache
        for row_id in self.tv.get_children():
            idx = int(row_id)
            if 0 <= idx < len(self.visible_values):
                v = self.visible_values[idx]
                mark = "☑" if checked_map.get(v, False) else "☐"
                self.tv.item(row_id, values=(mark, disp_cache[v], f"{val_counts[v]:,}"))
        self._update_match_info_label()

    def _toggle_select_all(self):
        self._user_toggled_items = True
        target = self.all_selected_var.get()
        for v in self.visible_values:
            self.checked_map[v] = target
        self._refresh_rendered_checkmarks()

    def _select_only_visible(self):
        """現在表示中の検索候補のみを選択し、それ以外をすべて解除する"""
        self._user_toggled_items = True
        vis_set = set(self.visible_values)
        for v in self.all_values:
            self.checked_map[v] = (v in vis_set)
        self._refresh_rendered_checkmarks()

    def _check_all_values(self):
        self._user_toggled_items = True
        for v in self.all_values:
            self.checked_map[v] = True
        self._refresh_rendered_checkmarks()

    def _uncheck_all_values(self):
        self._user_toggled_items = True
        for v in self.all_values:
            self.checked_map[v] = False
        self._refresh_rendered_checkmarks()

    def _set_row_checked(self, row_id: str, new_state: bool):
        if not row_id:
            return
        idx = int(row_id)
        if 0 <= idx < len(self.visible_values):
            self._user_toggled_items = True
            val = self.visible_values[idx]
            if self.checked_map.get(val) == new_state:
                return
            self.checked_map[val] = new_state
            mark = "☑" if new_state else "☐"
            self.tv.item(row_id, values=(mark, self._disp_cache[val], f"{self.value_counts[val]:,}"))
            self._update_match_info_label()

    def _on_tv_click(self, event):
        region = self.tv.identify_region(event.x, event.y)
        if region not in ("cell", "tree"):
            return
        row_id = self.tv.identify_row(event.y)
        if not row_id:
            return
        idx = int(row_id)
        if 0 <= idx < len(self.visible_values):
            val = self.visible_values[idx]
            new_state = not self.checked_map.get(val, False)
            self._drag_check_target = new_state
            self._set_row_checked(row_id, new_state)
            self.tv.selection_set(row_id)
            self.tv.focus(row_id)

    def _on_tv_drag(self, event):
        if self._drag_check_target is None:
            return
        row_id = self.tv.identify_row(event.y)
        if row_id:
            self._set_row_checked(row_id, self._drag_check_target)

    def _on_tv_space(self, event):
        sel = self.tv.selection()
        if not sel:
            return "break"
        row_id = sel[0]
        idx = int(row_id)
        if 0 <= idx < len(self.visible_values):
            val = self.visible_values[idx]
            self._set_row_checked(row_id, not self.checked_map.get(val, False))
        return "break"

    def _do_sort_asc(self):
        self.on_sort(self.col_name, False)
        self._close()

    def _do_sort_desc(self):
        self.on_sort(self.col_name, True)
        self._close()

    def _do_clear(self):
        self.on_clear_filter(self.col_name)
        self._close()

    def _do_ok(self):
        kw = self.search_var.get().strip()
        checked_all = {v for v in self.all_values if self.checked_map.get(v, False)}

        # 全解除したあと検索ワードを入力してそのままOKを押した場合は、現在の検索一致候補を採用する
        if kw and not checked_all and self.visible_values:
            checked_all = set(self.visible_values)

        if len(checked_all) == len(self.all_values):
            selected_set = None
        else:
            selected_set = checked_all

        num_op = self.num_op_var.get()
        num_v1 = None
        num_v2 = None
        if self.is_numeric and num_op not in ("指定なし", "0より大きい (>0)", "0のみ (=0)"):
            s1 = self.num_v1_var.get().strip()
            if not s1:
                messagebox.showwarning("入力確認", "数値条件の値を入力してください。", parent=self)
                return
            try:
                num_v1 = float(s1.replace(",", ""))
                if not math.isfinite(num_v1):
                    raise ValueError("non-finite")
            except ValueError:
                messagebox.showwarning("入力確認", "正しい数値を入力してください。", parent=self)
                return
            if num_op == "範囲 (最小～最大)":
                s2 = self.num_v2_var.get().strip()
                if not s2:
                    messagebox.showwarning("入力確認", "範囲の最大値を入力してください。", parent=self)
                    return
                try:
                    num_v2 = float(s2.replace(",", ""))
                    if not math.isfinite(num_v2):
                        raise ValueError("non-finite")
                except ValueError:
                    messagebox.showwarning("入力確認", "正しい数値を入力してください。", parent=self)
                    return
                if num_v1 > num_v2:
                    num_v1, num_v2 = num_v2, num_v1

        if selected_set is None and (not self.is_numeric or num_op == "指定なし"):
            self.on_clear_filter(self.col_name)
        else:
            self.on_apply_filter(
                self.col_name,
                {
                    "selected": selected_set,
                    "num_op": num_op if self.is_numeric else "指定なし",
                    "num_v1": num_v1,
                    "num_v2": num_v2,
                },
            )
        self._close()


class ColumnCopyDialog(tk.Toplevel):
    """指定した列（単一または複数）をExcel貼り付け用形式でクリップボードにコピーするダイアログ"""

    def __init__(
        self,
        parent,
        visible_columns: list,
        displayed_count: int,
        selected_count: int,
        on_copy,
        display_unit: str = DEFAULT_TIME_UNIT,
    ):
        super().__init__(parent)
        self.title("列データのコピー（Excel貼付用）")
        self.geometry("500x560")
        self.minsize(420, 460)
        self.configure(bg=COLORS["bg_card"])
        self.transient(parent)
        self.grab_set()

        self.visible_columns = list(visible_columns)
        self.on_copy = on_copy
        self.display_unit = display_unit if display_unit in VALID_TIME_UNITS else DEFAULT_TIME_UNIT
        self.checked_map = {c: (c in CALC_COLUMNS or c == "管理No") for c in self.visible_columns}

        frame = ttk.Frame(self, style="Card.TFrame", padding=14)
        frame.pack(fill=tk.BOTH, expand=True)

        footer = ttk.Frame(frame, style="Flat.TFrame")
        footer.pack(side=tk.BOTTOM, fill=tk.X, pady=(6, 0))
        ttk.Button(footer, text="キャンセル", command=self.destroy).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(footer, text="クリップボードにコピー", style="Primary.TButton", command=self._do_copy).pack(side=tk.RIGHT)

        opt_box = ttk.LabelFrame(frame, text="コピー形式オプション", padding=(8, 6))
        opt_box.pack(side=tk.BOTTOM, fill=tk.X, pady=(8, 4))

        self.include_header_var = tk.BooleanVar(value=True)
        self.raw_number_var = tk.BooleanVar(value=True)
        self.scope_var = tk.StringVar(value="all" if selected_count <= 1 else "selected")

        tk.Checkbutton(
            opt_box,
            text="1行目に列見出しを含める",
            variable=self.include_header_var,
            bg=COLORS["bg_card"],
            fg=COLORS["fg_main"],
            activebackground=COLORS["bg_card"],
            selectcolor="#FFFFFF",
            font=("Yu Gothic UI", 9),
            anchor=tk.W,
        ).pack(fill=tk.X)
        tk.Checkbutton(
            opt_box,
            text="数値列はカンマ無しでコピー（Excel計算用）",
            variable=self.raw_number_var,
            bg=COLORS["bg_card"],
            fg=COLORS["fg_main"],
            activebackground=COLORS["bg_card"],
            selectcolor="#FFFFFF",
            font=("Yu Gothic UI", 9),
            anchor=tk.W,
        ).pack(fill=tk.X)

        scope_f = ttk.Frame(opt_box, style="Flat.TFrame")
        scope_f.pack(fill=tk.X, pady=(4, 0))
        tk.Radiobutton(
            scope_f,
            text=f"表示中の全行 ({displayed_count:,}件)",
            variable=self.scope_var,
            value="all",
            bg=COLORS["bg_card"],
            fg=COLORS["fg_main"],
            activebackground=COLORS["bg_card"],
            selectcolor="#FFFFFF",
            font=("Yu Gothic UI", 9),
        ).pack(side=tk.LEFT, padx=(0, 12))
        rb_sel = tk.Radiobutton(
            scope_f,
            text=f"選択中の行のみ ({selected_count:,}件)",
            variable=self.scope_var,
            value="selected",
            bg=COLORS["bg_card"],
            fg=COLORS["fg_main"],
            activebackground=COLORS["bg_card"],
            selectcolor="#FFFFFF",
            font=("Yu Gothic UI", 9),
        )
        rb_sel.pack(side=tk.LEFT)
        if selected_count == 0:
            rb_sel.configure(state=tk.DISABLED)

        ttk.Label(frame, text="コピーする列を選択してください", style="Title.TLabel").pack(anchor=tk.W)
        ttk.Label(
            frame,
            text="選択した列をタブ区切りでクリップボードへコピーします。別のExcelで Ctrl+V するだけで貼り付けできます。",
            style="Mtime.TLabel",
            wraplength=440,
        ).pack(anchor=tk.W, pady=(2, 8))

        quick_f = ttk.Frame(frame, style="Flat.TFrame")
        quick_f.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(quick_f, text="すべて選択", command=lambda: self._set_preset("all")).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(quick_f, text="すべて解除", command=lambda: self._set_preset("none")).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(quick_f, text="その他見出しのみ", command=lambda: self._set_preset("meta")).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(quick_f, text="集計・小計のみ", command=lambda: self._set_preset("calc")).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(quick_f, text="管理No＋集計", command=lambda: self._set_preset("kno_calc")).pack(side=tk.LEFT)

        tv_frame = ttk.Frame(frame, style="Card.TFrame")
        tv_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.tv = ttk.Treeview(tv_frame, columns=("chk", "col", "kind"), show="headings", selectmode="browse", height=6)
        vsb = ttk.Scrollbar(tv_frame, orient=tk.VERTICAL, command=self.tv.yview)
        self.tv.configure(yscrollcommand=vsb.set)
        self.tv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tv.heading("chk", text="選択")
        self.tv.heading("col", text="列名（ダブルクリックで即その1列のみコピー）")
        self.tv.heading("kind", text="区分")
        self.tv.column("chk", width=50, anchor=tk.CENTER, stretch=False)
        self.tv.column("col", width=260, anchor=tk.W)
        self.tv.column("kind", width=85, anchor=tk.CENTER, stretch=False)

        self.tv.bind("<Button-1>", self._on_click)
        self.tv.bind("<Double-1>", self._on_double_click)
        self.tv.bind("<space>", self._on_space)

        self._refresh()
        self.bind("<Escape>", lambda e: self.destroy())

    def _refresh(self):
        self.tv.delete(*self.tv.get_children())
        for idx, col in enumerate(self.visible_columns):
            mark = "☑" if self.checked_map.get(col, False) else "☐"
            kind = "集計・小計" if col in CALC_COLUMNS else "基本情報"
            disp_name = get_display_column_label(col, self.display_unit)
            self.tv.insert("", tk.END, iid=str(idx), values=(mark, disp_name, kind))

    def _set_preset(self, mode: str):
        for col in self.visible_columns:
            if mode == "all":
                self.checked_map[col] = True
            elif mode == "none":
                self.checked_map[col] = False
            elif mode == "meta":
                self.checked_map[col] = col not in CALC_COLUMNS
            elif mode == "calc":
                self.checked_map[col] = col in CALC_COLUMNS
            elif mode == "kno_calc":
                self.checked_map[col] = (col == "管理No" or col in CALC_COLUMNS)
        self._refresh()

    def _toggle_row(self, row_id: str):
        if not row_id:
            return
        idx = int(row_id)
        if 0 <= idx < len(self.visible_columns):
            col = self.visible_columns[idx]
            self.checked_map[col] = not self.checked_map.get(col, False)
            mark = "☑" if self.checked_map[col] else "☐"
            kind = "集計・小計" if col in CALC_COLUMNS else "基本情報"
            disp_name = get_display_column_label(col, self.display_unit)
            self.tv.item(row_id, values=(mark, disp_name, kind))

    def _on_click(self, event):
        region = self.tv.identify_region(event.x, event.y)
        if region in ("cell", "tree"):
            self._toggle_row(self.tv.identify_row(event.y))

    def _on_double_click(self, event):
        row_id = self.tv.identify_row(event.y)
        if not row_id:
            return
        idx = int(row_id)
        if 0 <= idx < len(self.visible_columns):
            col = self.visible_columns[idx]
            self.on_copy(
                [col],
                include_header=self.include_header_var.get(),
                raw_numbers=self.raw_number_var.get(),
                selected_only=(self.scope_var.get() == "selected"),
            )
            self.destroy()

    def _on_space(self, event):
        sel = self.tv.selection()
        if sel:
            self._toggle_row(sel[0])

    def _do_copy(self):
        cols = [c for c in self.visible_columns if self.checked_map.get(c, False)]
        if not cols:
            messagebox.showwarning("選択確認", "コピーする列を1つ以上選択してください。", parent=self)
            return
        self.on_copy(
            cols,
            include_header=self.include_header_var.get(),
            raw_numbers=self.raw_number_var.get(),
            selected_only=(self.scope_var.get() == "selected"),
        )
        self.destroy()


class WorkRateSettingsDialog(tk.Toplevel):
    """起動時のデフォルト表示単位（金額/時間/工数）と日報単価レート（円/分）設定ダイアログ"""

    def __init__(self, parent, current_rates: dict, default_unit: str, current_unit: str, on_apply):
        super().__init__(parent)
        self.title("作業単価レート・既定表示設定")
        self.geometry("560x525")
        self.minsize(520, 485)
        self.resizable(True, True)
        self.configure(bg=COLORS["bg_card"])
        self.transient(parent)
        self.grab_set()

        self.on_apply = on_apply

        frame = ttk.Frame(self, style="Card.TFrame", padding=16)
        frame.pack(fill=tk.BOTH, expand=True)

        # 最下部フッターボタンを先に side=BOTTOM で固定配置（高DPI環境でも見切れ完全防止）
        btn_f = ttk.Frame(frame, style="Flat.TFrame")
        btn_f.pack(side=tk.BOTTOM, fill=tk.X, pady=(12, 0))

        ttk.Label(frame, text="既定表示単位 ＆ 作業単価レート設定", style="Title.TLabel").pack(anchor=tk.W)
        ttk.Label(
            frame,
            text="アプリ起動時のデフォルト表示単位と、「金額(円)」表示時に使用する1分あたりの単価を設定します。",
            style="Mtime.TLabel",
            wraplength=500,
        ).pack(anchor=tk.W, pady=(2, 10))

        unit_box = ttk.LabelFrame(frame, text="起動時のデフォルト表示単位（設定に保存）", padding=(10, 6))
        unit_box.pack(fill=tk.X, pady=(0, 10))

        init_def_unit = default_unit if default_unit in VALID_TIME_UNITS else DEFAULT_TIME_UNIT
        self.default_unit_var = tk.StringVar(value=init_def_unit)
        unit_grid = ttk.Frame(unit_box, style="Flat.TFrame")
        unit_grid.pack(fill=tk.X)
        for idx_u, (u_val, u_lbl) in enumerate((
            ("金額(円)", "金額(円)（デフォルト）"),
            ("分(そのまま)", "時間：分(そのまま)"),
            ("時間(分÷60)", "時間：時間(分÷60)"),
            ("工数(8h=1)", "工数(1工数=8時間)"),
        )):
            tk.Radiobutton(
                unit_grid,
                text=u_lbl,
                variable=self.default_unit_var,
                value=u_val,
                bg=COLORS["bg_card"],
                fg=COLORS["fg_main"],
                activebackground=COLORS["bg_card"],
                selectcolor="#FFFFFF",
                font=("Yu Gothic UI", 9),
            ).grid(row=idx_u // 2, column=idx_u % 2, sticky=tk.W, padx=(0, 18), pady=1)

        rate_box = ttk.LabelFrame(frame, text="日報作業 単価レート（円 / 分）", padding=(10, 6))
        rate_box.pack(fill=tk.X)

        grid_f = ttk.Frame(rate_box, style="Flat.TFrame")
        grid_f.pack(fill=tk.X, pady=2)

        self.entry_vars = {}
        for idx, cat in enumerate(WORK_CATEGORIES):
            ttk.Label(grid_f, text=f"{cat} :", style="Card.TLabel", width=16).grid(
                row=idx, column=0, sticky=tk.W, padx=6, pady=5
            )
            cur_val = current_rates.get(cat, DEFAULT_WORK_RATES[cat])
            val_str = str(int(cur_val)) if abs(cur_val - round(cur_val)) < 1e-6 else str(cur_val)
            v = tk.StringVar(value=val_str)
            self.entry_vars[cat] = v
            ttk.Entry(grid_f, textvariable=v, width=14, justify=tk.RIGHT).grid(
                row=idx, column=1, sticky=tk.W, padx=6, pady=5
            )
            ttk.Label(grid_f, text="円/分", style="Card.TLabel").grid(row=idx, column=2, sticky=tk.W, padx=(0, 8))
            def_r = DEFAULT_WORK_RATES[cat]
            def_str = str(int(def_r)) if abs(def_r - round(def_r)) < 1e-6 else str(def_r)
            ttk.Label(grid_f, text=f"(デフォルト: {def_str} 円/分)", style="Mtime.TLabel").grid(
                row=idx, column=3, sticky=tk.W, padx=4, pady=5
            )

        self.switch_now_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            frame,
            text="保存後、現在の一覧表示も上で選んだデフォルト単位に切り替える",
            variable=self.switch_now_var,
            bg=COLORS["bg_card"],
            fg=COLORS["fg_main"],
            activebackground=COLORS["bg_card"],
            selectcolor="#FFFFFF",
            font=("Yu Gothic UI", 9),
            anchor=tk.W,
        ).pack(fill=tk.X, pady=(10, 0))

        def restore_defaults():
            self.default_unit_var.set(DEFAULT_TIME_UNIT)
            for cat, def_val in DEFAULT_WORK_RATES.items():
                s = str(int(def_val)) if abs(def_val - round(def_val)) < 1e-6 else str(def_val)
                self.entry_vars[cat].set(s)

        def save_and_apply():
            new_rates = {}
            for cat in WORK_CATEGORIES:
                raw_s = self.entry_vars[cat].get().strip().replace(",", "")
                try:
                    val = float(raw_s)
                    if not math.isfinite(val) or val < 0:
                        raise ValueError("invalid")
                except ValueError:
                    messagebox.showwarning(
                        "入力確認",
                        f"「{cat}」のレートには0以上の有限数値を入力してください。\n例: 62.5",
                        parent=self,
                    )
                    return
                new_rates[cat] = val
            self.on_apply(new_rates, self.default_unit_var.get(), self.switch_now_var.get())
            self.destroy()

        ttk.Button(btn_f, text="デフォルトに戻す", command=restore_defaults).pack(side=tk.LEFT)
        ttk.Button(btn_f, text="キャンセル", command=self.destroy).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(btn_f, text="保存して適用", style="Primary.TButton", command=save_and_apply).pack(side=tk.RIGHT)
        self.bind("<Escape>", lambda e: self.destroy())


class ErrorReportDialog(tk.Toplevel):
    """サンプルとフォーマットが異なる場合や読み込みエラー発生時に詳細を確認できるダイアログ"""

    def __init__(self, parent, error_items: list):
        super().__init__(parent)
        self.title(f"エラー内容の確認 ({len(error_items)}件) - {APP_NAME}")
        self.geometry("880x580")
        self.minsize(720, 440)
        self.configure(bg=COLORS["bg_main"])
        self.transient(parent)
        self.grab_set()

        self.error_items = [self._normalize_item(i, item) for i, item in enumerate(error_items, 1)]

        # 1. 上部ヘッダー
        hdr = tk.Frame(self, bg=COLORS["danger"], padx=16, pady=10)
        hdr.pack(side=tk.TOP, fill=tk.X)
        tk.Label(
            hdr,
            text=f"フォーマット・読み込みエラーが確認されました（全 {len(self.error_items)} 件）",
            bg=COLORS["danger"],
            fg="#FFFFFF",
            font=("Yu Gothic UI", 12, "bold"),
        ).pack(side=tk.LEFT)
        tk.Label(
            hdr,
            text="行をクリックすると下部に詳細（原因・期待形式・検出内容）が表示されます",
            bg=COLORS["danger"],
            fg="#FEE2E2",
            font=("Yu Gothic UI", 9),
        ).pack(side=tk.RIGHT)

        # 2. 下部フッター（先に side=BOTTOM で配置してボタン切れを防止）
        ftr = tk.Frame(self, bg=COLORS["bg_sub"], padx=14, pady=8)
        ftr.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Button(ftr, text="エラー内容をすべてコピー", command=self._copy_all).pack(side=tk.LEFT)
        self.copy_msg_var = tk.StringVar(value="")
        tk.Label(ftr, textvariable=self.copy_msg_var, bg=COLORS["bg_sub"], fg=COLORS["accent"], font=("Yu Gothic UI", 9, "bold")).pack(
            side=tk.LEFT, padx=10
        )
        ttk.Button(ftr, text="閉じる", style="Primary.TButton", command=self.destroy).pack(side=tk.RIGHT)

        # 3. 中央コンテンツ（上下分割：一覧テーブル＋詳細テキスト）
        body = ttk.Frame(self, style="TFrame", padding=10)
        body.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        paned = tk.PanedWindow(body, orient=tk.VERTICAL, bg=COLORS["border"], sashwidth=6, relief=tk.FLAT)
        paned.pack(fill=tk.BOTH, expand=True)

        top_box = ttk.LabelFrame(paned, text="エラー発生ファイル一覧", padding=(6, 4))
        paned.add(top_box, minsize=160, stretch="always")

        cols = ("no", "category", "file", "kind", "reason")
        self.tv = ttk.Treeview(top_box, columns=cols, show="headings", selectmode="browse", height=6)
        vsb = ttk.Scrollbar(top_box, orient=tk.VERTICAL, command=self.tv.yview)
        hsb = ttk.Scrollbar(top_box, orient=tk.HORIZONTAL, command=self.tv.xview)
        self.tv.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tv.grid(row=0, column=0, sticky=tk.NSEW)
        vsb.grid(row=0, column=1, sticky=tk.NS)
        hsb.grid(row=1, column=0, sticky=tk.EW)
        top_box.rowconfigure(0, weight=1)
        top_box.columnconfigure(0, weight=1)

        self.tv.heading("no", text="No")
        self.tv.heading("category", text="対象種別")
        self.tv.heading("file", text="ファイル名")
        self.tv.heading("kind", text="エラー区分")
        self.tv.heading("reason", text="エラー概要")

        self.tv.column("no", width=45, anchor=tk.CENTER, stretch=False)
        self.tv.column("category", width=120, anchor=tk.W, stretch=False)
        self.tv.column("file", width=180, anchor=tk.W, stretch=False)
        self.tv.column("kind", width=125, anchor=tk.CENTER, stretch=False)
        self.tv.column("reason", width=360, anchor=tk.W)

        self.tv.tag_configure("even", background="#FFFFFF", foreground=COLORS["fg_main"])
        self.tv.tag_configure("odd", background="#FEF2F2", foreground=COLORS["fg_main"])

        bottom_box = ttk.LabelFrame(paned, text="選択したエラーの詳細内容（サンプルとの違い・確認ポイント）", padding=(8, 6))
        paned.add(bottom_box, minsize=160, stretch="always")

        self.detail_text = tk.Text(
            bottom_box,
            bg="#FFFFFF",
            fg=COLORS["fg_main"],
            font=("Yu Gothic UI", 10),
            wrap=tk.WORD,
            relief=tk.SOLID,
            bd=1,
            padx=10,
            pady=8,
            height=8,
        )
        d_vsb = ttk.Scrollbar(bottom_box, orient=tk.VERTICAL, command=self.detail_text.yview)
        self.detail_text.configure(yscrollcommand=d_vsb.set)
        self.detail_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        d_vsb.pack(side=tk.RIGHT, fill=tk.Y)

        for idx, item in enumerate(self.error_items):
            self.tv.insert(
                "",
                tk.END,
                iid=str(idx),
                values=(item["no"], item["category"], item["file"], item["kind"], item["reason"]),
                tags=("even" if idx % 2 == 0 else "odd",),
            )

        self.tv.bind("<<TreeviewSelect>>", self._on_select)
        if self.error_items:
            self.tv.selection_set("0")
            self._show_detail(0)

        self.bind("<Escape>", lambda e: self.destroy())

    @staticmethod
    def _normalize_item(no: int, item) -> dict:
        if isinstance(item, dict):
            return {
                "no": no,
                "category": item.get("category", "ファイル読込"),
                "file": item.get("file", "") or (os.path.basename(item.get("path", "")) if item.get("path") else "-"),
                "path": item.get("path", ""),
                "kind": item.get("kind", "エラー"),
                "reason": item.get("reason", "") or item.get("message", ""),
                "expected": item.get("expected", ""),
                "detected": item.get("detected", ""),
            }
        return {
            "no": no,
            "category": "ファイル読込",
            "file": "-",
            "path": "",
            "kind": "読込エラー",
            "reason": str(item),
            "expected": "",
            "detected": "",
        }

    def _format_item_detail(self, item: dict) -> str:
        lines = [
            f"【エラーNo】 {item['no']} / {len(self.error_items)}",
            f"【対象種別】 {item['category']}   （区分: {item['kind']}）",
            f"【ファイル名】 {item['file']}",
        ]
        if item.get("path"):
            lines.append(f"【ファイルパス】 {item['path']}")
        lines.append("")
        lines.append(f"【エラー内容・原因】\n  {item['reason']}")
        if item.get("expected"):
            lines.append("")
            lines.append(f"【期待されるサンプルフォーマット】\n  {item['expected']}")
        if item.get("detected"):
            lines.append("")
            lines.append(f"【ファイル内で検出された内容】\n  {item['detected']}")
        return "\n".join(lines)

    def _on_select(self, event=None):
        sel = self.tv.selection()
        if not sel:
            return
        try:
            idx = int(sel[0])
            if 0 <= idx < len(self.error_items):
                self._show_detail(idx)
        except ValueError:
            pass

    def _show_detail(self, idx: int):
        txt = self._format_item_detail(self.error_items[idx])
        self.detail_text.configure(state=tk.NORMAL)
        self.detail_text.delete("1.0", tk.END)
        self.detail_text.insert(tk.END, txt)
        self.detail_text.configure(state=tk.DISABLED)

    def _copy_all(self):
        blocks = [self._format_item_detail(item) for item in self.error_items]
        full_text = ("\n" + ("-" * 72) + "\n").join(blocks)
        self.clipboard_clear()
        self.clipboard_append(full_text)
        self.copy_msg_var.set("エラー詳細をクリップボードへコピーしました")


class HelpDialog(tk.Toplevel):
    """見やすく整理された目次付き・カード型ヘルプダイアログ（動的折り返し対応）"""

    def __init__(self, parent):
        super().__init__(parent)
        self.title(f"使い方ガイド - {APP_NAME}")
        self.geometry("1040x720")
        self.minsize(860, 560)
        self.configure(bg=COLORS["bg_main"])
        self.transient(parent)

        self._wrap_targets = []  # list of (label_widget, ratio_float, subtract_px)
        self._table_specs = []   # list of (grid_frame, ratios_list, cell_labels_by_col)
        self._last_canvas_w = 0

        # 上部ヘッダー
        hdr = tk.Frame(self, bg=COLORS["accent"], padx=18, pady=12)
        hdr.pack(side=tk.TOP, fill=tk.X)
        tk.Label(
            hdr,
            text=f"{APP_NAME}  操作ガイド＆ヘルプ",
            bg=COLORS["accent"],
            fg="#FFFFFF",
            font=("Yu Gothic UI", 14, "bold"),
        ).pack(side=tk.LEFT)
        tk.Label(
            hdr,
            text="制作伝票一覧 集計システム",
            bg=COLORS["accent"],
            fg="#DBEAFE",
            font=("Yu Gothic UI", 10),
        ).pack(side=tk.RIGHT)

        # フッター（先に side=BOTTOM で配置してボタンが途切れないようにする）
        ftr = tk.Frame(self, bg=COLORS["bg_sub"], padx=12, pady=8)
        ftr.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Button(ftr, text="閉じる", style="Primary.TButton", command=self._close).pack(side=tk.RIGHT)

        # メイン2ペイン（左：目次、右：スクロール可能コンテンツ）
        body = ttk.Frame(self, style="TFrame")
        body.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)

        nav_frame = ttk.Frame(body, style="Card.TFrame", padding=8)
        nav_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))

        ttk.Label(nav_frame, text="目次", style="Title.TLabel", padding=(4, 2, 4, 8)).pack(anchor=tk.W)

        content_outer = ttk.Frame(body, style="Card.TFrame")
        content_outer.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(content_outer, bg=COLORS["bg_card"], highlightthickness=0)
        vsb = ttk.Scrollbar(content_outer, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.inner = tk.Frame(self.canvas, bg=COLORS["bg_card"], padx=20, pady=16)
        self.canvas_win = self.canvas.create_window((0, 0), window=self.inner, anchor=tk.NW)
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        # グローバルの bind_all("<MouseWheel>") を破壊しないよう、専用属性＋Toplevelローカルバインドで安全にスクロール連動
        self._wheel_scroll_canvas = self.canvas
        self._dialog_closed = False
        self._init_wrap_job = None
        self.bind("<MouseWheel>", self._on_mousewheel)
        self.bind("<Escape>", lambda e: self._close())
        self.protocol("WM_DELETE_WINDOW", self._close)

        self.sections = {}
        self.nav_buttons = {}

        sections_meta = [
            ("overview", "1. アプリの概要"),
            ("files", "2. ファイル選択と保護"),
            ("codes", "3. 作業コード・レート設定"),
            ("filter", "4. 一覧とExcel風フィルター"),
            ("copy", "5. セル・範囲・列コピー"),
            ("columns", "6. 列の順番変更と保存"),
            ("config", "7. 設定の保存先と復元"),
            ("export", "8. Excel新規ブックで開く"),
            ("smartscreen", "9. 初回起動時の警告回避"),
        ]
        for key, label in sections_meta:
            btn = tk.Button(
                nav_frame,
                text=label,
                anchor=tk.W,
                relief=tk.FLAT,
                bg=COLORS["bg_card"],
                fg=COLORS["fg_main"],
                activebackground=COLORS["accent_light"],
                activeforeground=COLORS["accent"],
                font=("Yu Gothic UI", 10),
                padx=10,
                pady=7,
                width=20,
                cursor="hand2",
                command=lambda k=key: self._scroll_to(k),
            )
            btn.pack(fill=tk.X, pady=2)
            self.nav_buttons[key] = btn

        self._build_sections()
        self._init_wrap_job = self.after(30, self._run_init_wrap)

    def _run_init_wrap(self):
        self._init_wrap_job = None
        if self._dialog_closed:
            return
        try:
            if self.winfo_exists() and self.canvas.winfo_exists():
                self._update_wraplengths(max(620, self.canvas.winfo_width()))
        except Exception:
            pass

    def _close(self):
        self._dialog_closed = True
        if self._init_wrap_job is not None:
            try:
                self.after_cancel(self._init_wrap_job)
            except Exception:
                pass
            self._init_wrap_job = None
        try:
            self.destroy()
        except Exception:
            pass

    def _on_canvas_configure(self, event):
        w = max(360, int(event.width))
        self.canvas.itemconfigure(self.canvas_win, width=w)
        if abs(w - self._last_canvas_w) >= 4:
            self._update_wraplengths(w)

    def _update_wraplengths(self, canvas_w: int):
        self._last_canvas_w = canvas_w
        content_w = max(340, canvas_w - 44)
        for lbl, ratio, sub_px in self._wrap_targets:
            try:
                if lbl.winfo_exists():
                    lbl.configure(wraplength=max(120, int(content_w * ratio) - sub_px))
            except Exception:
                pass
        for grid, ratios, col_labels in self._table_specs:
            try:
                if not grid.winfo_exists():
                    continue
                tbl_w = max(320, content_w - 6)
                for c_i, r in enumerate(ratios):
                    col_px = max(46, int(tbl_w * r))
                    grid.columnconfigure(c_i, weight=max(1, int(round(r * 100))), minsize=col_px)
                    cell_wrap = max(34, col_px - 18)
                    for lbl in col_labels[c_i]:
                        if lbl.winfo_exists():
                            lbl.configure(wraplength=cell_wrap)
            except Exception:
                pass
        try:
            self.inner.update_idletasks()
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        except Exception:
            pass

    def _on_mousewheel(self, event):
        try:
            if self._dialog_closed or not self.winfo_exists():
                return
            delta = -1 if event.delta > 0 else 1
            steps = max(1, abs(int(event.delta / 120))) * 3
            self.canvas.yview_scroll(delta * steps, "units")
            return "break"
        except Exception:
            pass

    def _scroll_to(self, key: str):
        widget = self.sections.get(key)
        if not widget:
            return
        self.inner.update_idletasks()
        total_h = max(1, self.inner.winfo_height())
        y_pos = widget.winfo_y()
        self.canvas.yview_moveto(y_pos / total_h)
        for k, b in self.nav_buttons.items():
            if k == key:
                b.configure(bg=COLORS["accent_light"], fg=COLORS["accent"], font=("Yu Gothic UI", 10, "bold"))
            else:
                b.configure(bg=COLORS["bg_card"], fg=COLORS["fg_main"], font=("Yu Gothic UI", 10))

    def _add_section_header(self, key: str, num: str, title: str):
        box = tk.Frame(self.inner, bg=COLORS["bg_card"], pady=6)
        box.pack(fill=tk.X, pady=(10, 4))
        self.sections[key] = box

        badge = tk.Label(
            box,
            text=num,
            bg=COLORS["accent"],
            fg="#FFFFFF",
            font=("Yu Gothic UI", 10, "bold"),
            padx=8,
            pady=2,
        )
        badge.pack(side=tk.LEFT, padx=(0, 8))
        tk.Label(
            box,
            text=title,
            bg=COLORS["bg_card"],
            fg=COLORS["fg_main"],
            font=("Yu Gothic UI", 12, "bold"),
        ).pack(side=tk.LEFT)

        tk.Frame(self.inner, bg=COLORS["accent_light"], height=2).pack(fill=tk.X, pady=(0, 8))

    def _add_paragraph(self, text: str):
        lbl = tk.Label(
            self.inner,
            text=text,
            bg=COLORS["bg_card"],
            fg=COLORS["fg_main"],
            font=("Yu Gothic UI", 10),
            justify=tk.LEFT,
            anchor=tk.W,
            wraplength=700,
        )
        lbl.pack(fill=tk.X, pady=(0, 6))
        self._wrap_targets.append((lbl, 1.0, 4))

    def _add_callout(self, title: str, text: str, bg_color: str, border_color: str):
        card = tk.Frame(self.inner, bg=bg_color, highlightbackground=border_color, highlightthickness=1, padx=12, pady=8)
        card.pack(fill=tk.X, pady=(4, 10))
        tk.Label(card, text=title, bg=bg_color, fg=COLORS["fg_main"], font=("Yu Gothic UI", 10, "bold"), anchor=tk.W).pack(fill=tk.X)
        lbl = tk.Label(
            card,
            text=text,
            bg=bg_color,
            fg=COLORS["fg_sub"],
            font=("Yu Gothic UI", 9),
            justify=tk.LEFT,
            anchor=tk.W,
            wraplength=670,
        )
        lbl.pack(fill=tk.X, pady=(2, 0))
        self._wrap_targets.append((lbl, 1.0, 30))

    def _add_table(self, headers: list, rows: list, col_ratios: list = None):
        n_cols = len(headers)
        if not col_ratios or len(col_ratios) != n_cols:
            col_ratios = [1.0 / n_cols] * n_cols
        else:
            s = sum(col_ratios) or 1.0
            col_ratios = [r / s for r in col_ratios]

        tbl = tk.Frame(self.inner, bg=COLORS["border"], padx=1, pady=1)
        tbl.pack(fill=tk.X, pady=(4, 10))
        grid = tk.Frame(tbl, bg=COLORS["border"])
        grid.pack(fill=tk.X)

        col_labels = [[] for _ in range(n_cols)]
        init_tbl_w = 700

        for c_i, h in enumerate(headers):
            col_px = max(46, int(init_tbl_w * col_ratios[c_i]))
            anchor_h = tk.CENTER if (c_i == 0 and col_ratios[c_i] <= 0.12) else tk.W
            lbl = tk.Label(
                grid,
                text=h,
                bg=COLORS["bg_header"],
                fg=COLORS["accent"],
                font=("Yu Gothic UI", 9, "bold"),
                padx=8,
                pady=6,
                anchor=anchor_h,
                justify=tk.LEFT,
                wraplength=max(34, col_px - 18),
            )
            lbl.grid(row=0, column=c_i, sticky=tk.NSEW, padx=(0 if c_i == 0 else 1, 0), pady=(0, 1))
            grid.columnconfigure(c_i, weight=max(1, int(round(col_ratios[c_i] * 100))), minsize=col_px)
            col_labels[c_i].append(lbl)

        for r_i, r_vals in enumerate(rows, 1):
            bg = "#FFFFFF" if r_i % 2 == 1 else COLORS["bg_sub"]
            for c_i, val in enumerate(r_vals):
                col_px = max(46, int(init_tbl_w * col_ratios[c_i]))
                anchor_c = tk.N if (c_i == 0 and col_ratios[c_i] <= 0.12) else tk.NW
                lbl = tk.Label(
                    grid,
                    text=val,
                    bg=bg,
                    fg=COLORS["fg_main"],
                    font=("Yu Gothic UI", 9),
                    padx=8,
                    pady=5,
                    anchor=anchor_c,
                    justify=tk.LEFT,
                    wraplength=max(34, col_px - 18),
                )
                lbl.grid(row=r_i, column=c_i, sticky=tk.NSEW, padx=(0 if c_i == 0 else 1, 0), pady=(0, 1 if r_i < len(rows) else 0))
                col_labels[c_i].append(lbl)

        self._table_specs.append((grid, col_ratios, col_labels))

    def _build_sections(self):
        # 1. 概要
        self._add_section_header("overview", "STEP 1", "アプリの概要と基本の流れ")
        self._add_paragraph(
            "DenpyoHub_Miya は、日報ファイル(CSV)・購入品(XLSM/XLSX)・外注加工(XLSM/XLSX)・受注表(XLSM/XLSX)および"
            "制作伝票フォルダ（指定フォルダ直下の .xls / .xlsx / .xlsm・サブフォルダは含まない）を読み込み、管理Noごとに作業時間・金額と購入・外注・材料費用を自動集計するデスクトップアプリです。"
        )
        self._add_table(
            ["順番", "操作", "内容"],
            [
                ["1", "参照ファイルを選択", "起動時は自動集計しません。各ファイル・制作伝票フォルダを選択します（複数選択対応・制作伝票は指定フォルダ直下のみ参照）。"],
                ["2", "「集計・一覧更新」を押す", "ボタンを押すとファイルを検証・読み込み、管理Noごとに一覧テーブルとサマリーを作成します。"],
                ["3", "絞り込み・確認", "Excel風フィルターや行リンク表示、小計1/2の内訳開閉ボタンで、一覧と明細（日報・購入・外注・受注・制作伝票）を確認します。"],
                ["4", "Excel新規ブックで開く", "現在の一覧（列順番・フィルター・表示単位・見出し反映済み）をExcelの新規ブックとして直接開きます。"],
            ],
            col_ratios=[0.08, 0.26, 0.66],
        )

        # 2. ファイル複数選択と読み取り専用保護・フォーマット検証
        self._add_section_header("files", "STEP 2", "参照ファイル＆制作伝票フォルダ（読み取り専用保護・フォーマット検証）")
        self._add_paragraph(
            "・日報／購入品／外注加工／受注表は各行の「選択(複数可)...」や「追加...」で複数ファイルを選択できます。\n"
            "・制作伝票フォルダの行では「選択(複数可)」から複数フォルダをチェック選択・追加・削除できます。選択した各フォルダ直下にある .xls / .xlsx / .xlsm ファイルのみを自動検出し（サブフォルダは自動探索しません）、調整用(21行)・材料費各種(28～35行)・設計費(36行)・重機/運搬費(AA17)・車両/交通費(AA18)・出張費(AA19)・総請求金額(AA30)・数量(AD9)を読み込み、小計1・消耗品費(6%)・管理費・小計2・小計1+小計2・諸経費(8%)・合計・合計/数量・調整額・総請求金額/数量を自動算出します。\n"
            "・各サンプルファイルとフォーマットが異なるファイルが含まれる場合はエラー通知が表示され、「エラー確認」ボタンから原因・期待形式・検出内容を詳しく確認できます。"
        )
        self._add_callout(
            "参照元ファイル・フォルダは変更されないよう厳重に保護されています",
            "選択した日報・購入品・外注加工・受注表・制作伝票フォルダ内の全ファイルは、常に共有読み取り専用で参照され、他のユーザーがExcelで開いている状態でも支障なく集計・更新できます（アプリから元のファイルやフォルダを書き換えることは一切ありません）。\n"
            "また、設定保存や出力時に参照元ファイルや制作伝票フォルダ内への書き込みを指定した場合も自動で厳重にブロックします。",
            COLORS["bg_header"],
            COLORS["accent"],
        )

        # 3. 作業コード・レート設定
        self._add_section_header("codes", "STEP 3", "表示切替（金額 / 時間 / 工数）と単価レート・作業コード設定")
        self._add_paragraph(
            "・初期設定では一覧の作業データは「金額(円)」で表示され、見出しも「作業費(円)」「移動費(円)」「修理改造費(円)」等に分かりやすく切り替わります。\n"
            "・一覧上部の「一括表示」ボタン（金額(円) / 時間(分) / 時間(h) / 工数(8h)）からワンクリックで一括切り替えできます（1工数＝8時間）。\n"
            "・「表示・レート設定...」から起動時のデフォルト表示単位と各区分レート（円/分）を変更・保存できます。\n"
            "・「作業コード設定...」から日報の作業コード割り当て範囲も変更できます。"
        )
        self._add_table(
            ["集計項目名", "デフォルト範囲", "デフォルトレート", "小計の扱い"],
            [
                ["作業時間", "20～49", "62.5 円/分", "作業小計 に加算"],
                ["移動時間", "90～99", "50 円/分", "作業小計 に加算"],
                ["修理改造時間", "50～69", "70 円/分", "作業小計 に加算"],
                ["設計関係", "1～19", "62.5 円/分", "作業小計 に加算"],
                ["その他作業", "70～89", "62.5 円/分", "作業小計 に加算"],
                ["購入品 / 外注加工", "Q列（金額 単価X必要数）", "-", "購入品小計 に加算"],
                ["制作伝票（フォルダ読込）", "21行, 28～36行, AA17～19, AA30, AD9", "-", "小計1 / 消耗品費(6%) / 管理費 / 小計2 / 小計1+小計2 / 諸経費(8%) / 合計 / 合計÷数量 / 調整額 / 総請求金額÷数量"],
            ],
            col_ratios=[0.22, 0.24, 0.14, 0.40],
        )

        # 4. 一覧とフィルター
        self._add_section_header("filter", "STEP 4", "左右分割一覧（行リンク連動・内訳開閉）とExcel風フィルター")
        self._add_paragraph(
            "・画面中央の一覧は左（その他見出し・受注情報）と右（日報/購入/外注/制作伝票の算出データ・各小計）に分かれています。\n"
            "・初期状態では「小計1」「小計2」の内訳列を隠してスッキリ表示しています。一覧上部の「▶ 小計1内訳を表示」「▶ 小計2内訳を表示」「全内訳を表示」「小計のみ表示」ボタンでいつでも開閉できます。\n"
            "・「行リンク連動」がONの間は、縦スクロール・並び替え・選択行のハイライトが左右で自動的に揃うため、横並びで比較しやすくなります（横スクロールやセル選択・コピーは左右別々に操作できます）。\n"
            "・セルや行をクリックすると、左右の該当行が青く連動強調され、画面下部のタブにその管理Noの明細が表示されます。\n"
            "・ファイル設定エリアは通常閉じており、右上の「▼ ファイル設定を開く」ボタンを押したときだけ開きます（もう一度押すと畳みます）。"
        )
        self._add_callout(
            "Excel風オートフィルタ＆全体検索の使い方",
            "各列の見出しの右端「▽」をクリック（または見出し上で右クリック）すると、絞り込みポップアップが開きます。\n"
            "半角・全角カタカナや英数字の違い、数値のカンマ有無を気にせず検索でき、「検索結果で即適用」ボタンや複数キーワードの連続チェック選択、数値条件（0より大きい、範囲指定など）にも対応しています。",
            COLORS["warning_bg"],
            COLORS["warning"],
        )

        # 5. セル・範囲・列データのコピー
        self._add_section_header("copy", "STEP 5", "ボタン＆ショートカットでのセル・範囲・列コピー")
        self._add_paragraph(
            "ショートカットだけでなく、各一覧下部の専用ボタンからも簡単にExcel貼り付け用コピーができます。\n"
            "・ボタン操作：「選択コピー」「見出し付コピー」「この列を全コピー」「全セル選択」「選択解除」\n"
            "・マウス＆ショートカット：クリック（1セル）、ドラッグ / Shift+クリック（矩形範囲）、Ctrl+クリック（複数追加）、\n"
            "  見出しをCtrl+クリック（列まるごと選択）、Ctrl+C（コピー）、Ctrl+Shift+C（見出し付きコピー）、Ctrl+A（全選択）\n"
            "・上部の「列を選んでコピー...」ボタン：左右をまたいで複数列をまとめて選択コピー\n"
            "・下部の内訳明細タブ（日報・購入品・外注加工・受注表・制作伝票）：明細行の Ctrl+C または右クリックメニューからコピー可能"
        )

        # 6. 列順変更
        self._add_section_header("columns", "STEP 6", "列の順番変更と表示設定")
        self._add_paragraph(
            "画面上部の「列の順番・表示...」ボタン（または列見出しの右クリックメニュー）から、\n"
            "すべての列の並び順や表示/非表示を自由に変更できます。変更した順番は保存され、Excel表示時にも同じ順番で開きます。"
        )

        # 7. 設定保存
        self._add_section_header("config", "STEP 7", "設定の保存（デフォルト：アプリと同じ場所）")
        self._add_paragraph(
            "参照ファイル一覧、作業コード設定、単価レート設定、列の順番、列幅、分割比率、表示設定などは、\n"
            "初期状態でアプリと同じフォルダの「denpyohub_config.json」に安全なアトミック書き込みで自動保存されます。\n"
            "「設定保存 / 読込...」から明示的に保存したり、別の場所へ設定ファイルを保存/読み込み・初期化することも可能です。"
        )

        # 8. Excel新規ブック
        self._add_section_header("export", "STEP 8", "一覧および内訳明細をExcelの新規ブックとして開く（印刷設定済み）")
        self._add_paragraph(
            "・一覧上部の「Excel新規ブックで開く」ボタンを押すと、現在表示されているフィルター条件・表示単位・列の順番のまま、A3横向き（非対応プリンタではA4横）・1ページ幅収め・列グループ化（小計1/2内訳の開閉連動）設定済みの新規Excelブックとして開きます。\n"
            "・下部「選択した管理Noの内訳明細」にある「表示中タブをExcelで開く」「内訳明細をExcelで開く（全タブ）」ボタンを押すと、選択中の管理Noの明細（日報・購入品・外注加工・受注表詳細・制作伝票）を、タイトル・合計行・A4横1ページ幅収め・見出し行繰り返し・ページレイアウト表示設定済みのExcel新規ブックで開きます。"
        )

        # 9. SmartScreen
        self._add_section_header("smartscreen", "補足", "初回起動時に青い警告画面が出た場合")
        self._add_callout(
            "WindowsによってPCが保護されました（SmartScreen）の回避方法",
            "1. 青い画面内の「詳細情報」をクリックします。\n"
            "2. 画面下部に表示される「実行」ボタンをクリックします。\n"
            "※ 2回目以降は通常通りすぐに起動します。",
            "#FEF2F2",
            COLORS["danger"],
        )


class SplitGridPane(ttk.LabelFrame):
    """独立したスクロール・並替・セル選択・ボタンコピー機能を持つ一覧テーブルペイン"""

    def __init__(self, parent, title: str, pane_id: str, app):
        super().__init__(parent, text=title, padding=(6, 4))
        self.app = app
        self.pane_id = pane_id
        self.visible_columns = []
        self.rows = []
        self.kno_to_row_idx = {}
        self.sort_col = "管理No"
        self.sort_desc = True
        self._linked_kno = None

        # セル選択管理
        self.selected_cells = set()       # {(row_idx, col_name), ...}
        self.anchor_cell = None
        self.active_cell = None
        self._drag_selecting = False
        self._drag_ctrl = False
        self._drag_base_cells = set()
        self._overlay_borders = []
        self._overlay_after_id = None
        self._chunk_after_id = None
        self._render_job_id = 0

        self.cell_info_var = tk.StringVar(value="件数: 0件  |  選択: なし（セルクリック/ドラッグで選択してボタンコピー可）")

        # 下部：選択セル情報＋ボタンコピー操作バー（先に side=BOTTOM で配置してボタンが途切れないようにする）
        footer = tk.Frame(self, bg=COLORS["bg_header"], padx=6, pady=4)
        footer.pack(side=tk.BOTTOM, fill=tk.X, pady=(3, 0))

        tk.Label(
            footer,
            textvariable=self.cell_info_var,
            bg=COLORS["bg_header"],
            fg=COLORS["accent"],
            font=("Yu Gothic UI", 9, "bold"),
            anchor=tk.W,
        ).pack(side=tk.TOP, fill=tk.X, pady=(0, 2))

        btn_bar = tk.Frame(footer, bg=COLORS["bg_header"])
        btn_bar.pack(side=tk.TOP, fill=tk.X)

        ttk.Button(btn_bar, text="選択コピー", command=lambda: self.copy_selection(include_header=False)).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Button(btn_bar, text="見出し付コピー", command=lambda: self.copy_selection(include_header=True)).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_bar, text="この列を全コピー", command=self.copy_active_columns_action).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_bar, text="全セル選択", command=self.select_all_cells).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_bar, text="選択解除", command=self.clear_selection).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_bar, text="列幅調整", command=self.autofit_all_columns).pack(side=tk.RIGHT, padx=(3, 0))

        # 上部：テーブル本体（左右ペインで完全に同じY座標・高さに揃える）
        tree_container = ttk.Frame(self, style="Card.TFrame")
        tree_container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(
            tree_container,
            columns=(),
            show="headings",
            selectmode="extended",
            height=5,
        )
        vsb = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self._on_vsb_scroll)
        hsb = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL, command=self._on_hsb_scroll)

        def on_yscroll(first, last):
            vsb.set(first, last)
            self._schedule_overlay_update()
            self.app._on_pane_y_scroll(self.pane_id, first)

        def on_xscroll(first, last):
            hsb.set(first, last)
            self._schedule_overlay_update()

        self.tree.configure(yscrollcommand=on_yscroll, xscrollcommand=on_xscroll)
        self.tree.grid(row=0, column=0, sticky=tk.NSEW)
        vsb.grid(row=0, column=1, sticky=tk.NS)
        hsb.grid(row=1, column=0, sticky=tk.EW)
        tree_container.rowconfigure(0, weight=1)
        tree_container.columnconfigure(0, weight=1)

        self.tree.tag_configure("even", background=COLORS["row_even"], foreground=COLORS["fg_main"])
        self.tree.tag_configure("odd", background=COLORS["row_odd"], foreground=COLORS["fg_main"])
        self.tree.tag_configure("active_both", background="#EFF6FF", foreground=COLORS["fg_main"])
        self.tree.tag_configure("linked_row", background="#DBEAFE", foreground="#1E3A8A")

        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Button-1>", self._on_left_click)
        self.tree.bind("<B1-Motion>", self._on_left_drag)
        self.tree.bind("<Button-3>", self._on_right_click)
        self.tree.bind("<ButtonRelease-1>", self._on_button_release)
        self.tree.bind("<Configure>", lambda e: self._schedule_overlay_update())
        self.tree.bind("<Control-c>", lambda e: self.copy_selection(include_header=False))
        self.tree.bind("<Control-C>", lambda e: self.copy_selection(include_header=False))
        self.tree.bind("<Control-Shift-C>", lambda e: self.copy_selection(include_header=True))
        self.tree.bind("<Control-a>", lambda e: self.select_all_cells())
        self.tree.bind("<Control-A>", lambda e: self.select_all_cells())

        self.row_menu = tk.Menu(self, tearoff=0, bg="#FFFFFF", fg=COLORS["fg_main"])
        self.header_menu = tk.Menu(self, tearoff=0, bg="#FFFFFF", fg=COLORS["fg_main"])

    def _width_key(self, col: str) -> str:
        """右ペインの識別用管理No列幅を左ペインと独立して保存するキーを返す"""
        if self.pane_id == "calc" and col == "管理No":
            return "_calc_管理No"
        return col

    def configure_columns(self, columns: list):
        """表示列と保存済み列幅・見出しテキストを適用する"""
        self.visible_columns = list(columns)
        self.tree.configure(columns=self.visible_columns)
        for col in self.visible_columns:
            anchor = tk.E if col in NUMERIC_COLUMNS else tk.W
            wk = self._width_key(col)
            w = self.app.column_widths.get(wk, self.app.column_widths.get(col, DEFAULT_COL_WIDTHS.get(col, 90)))
            self.tree.column(col, width=w, minwidth=45, anchor=anchor, stretch=False)
        self.update_heading_texts()

    def update_heading_texts(self):
        """各列見出し（金額/分/時間の単位付き）にソート記号とフィルター記号を表示する"""
        linked = self.app.link_rows_var.get()
        cur_sort_col = self.app.sort_col if linked else self.sort_col
        cur_sort_desc = self.app.sort_desc if linked else self.sort_desc
        cur_unit = self.app.time_unit_var.get()
        for col in self.visible_columns:
            disp = get_display_column_label(col, cur_unit)
            parts = [disp]
            if cur_sort_col == col:
                parts.append("↓" if cur_sort_desc else "↑")
            if col in self.app.column_filters:
                parts.append("▼[絞込]")
            else:
                parts.append("▽")
            self.tree.heading(col, text=" ".join(parts))

    def capture_column_widths(self) -> dict:
        """現在の列幅辞書を返す"""
        widths = {}
        try:
            for col in self.visible_columns:
                w = self.tree.column(col, "width")
                if isinstance(w, int) and w >= 35:
                    widths[self._width_key(col)] = w
        except Exception:
            pass
        return widths

    def _on_vsb_scroll(self, *args):
        self.tree.yview(*args)
        self._schedule_overlay_update()

    def _on_hsb_scroll(self, *args):
        self.tree.xview(*args)
        self._schedule_overlay_update()

    def _schedule_overlay_update(self):
        if self.app._closing:
            return
        try:
            if not self.tree.winfo_exists():
                return
        except Exception:
            return
        if self._overlay_after_id is not None:
            try:
                self.app.root.after_cancel(self._overlay_after_id)
            except Exception:
                pass
            self._overlay_after_id = None
        try:
            self._overlay_after_id = self.app.root.after(15, self._update_selection_overlay)
        except Exception:
            self._overlay_after_id = None

    def cancel_pending_jobs(self):
        self._render_job_id += 1
        for attr in ("_overlay_after_id", "_chunk_after_id"):
            job = getattr(self, attr, None)
            if job is not None:
                try:
                    self.app.root.after_cancel(job)
                except Exception:
                    pass
                setattr(self, attr, None)

    def _clear_overlay_borders(self):
        for w in self._overlay_borders:
            try:
                w.place_forget()
            except Exception:
                pass

    def _get_overlay_border_widget(self, idx: int) -> tk.Frame:
        while len(self._overlay_borders) <= idx:
            f = tk.Frame(self.tree, bg=COLORS["accent"], bd=0, highlightthickness=0)
            self._overlay_borders.append(f)
        return self._overlay_borders[idx]

    def _get_tree_header_height(self) -> int:
        for y in range(18, 60, 2):
            if self.tree.identify_region(10, y) != "heading":
                return y
        return 28

    def _update_selection_overlay(self):
        """選択セル群の外周をExcel風の青い枠線で高速描画する（表示範囲行のみ）"""
        self._overlay_after_id = None
        if self.app._closing:
            return
        try:
            if not self.tree.winfo_exists():
                return
        except Exception:
            return

        self._clear_overlay_borders()
        if not self.selected_cells or not self.visible_columns:
            return

        tree_w = self.tree.winfo_width()
        tree_h = self.tree.winfo_height()
        if tree_w <= 10 or tree_h <= 40:
            return

        hdr_h = self._get_tree_header_height()
        vis_cols = self.visible_columns
        col_to_idx = {c: i for i, c in enumerate(vis_cols)}

        sel_idx_set = set()
        for r_idx, c_name in self.selected_cells:
            c_idx = col_to_idx.get(c_name)
            if c_idx is not None and 0 <= r_idx < len(self.rows):
                sel_idx_set.add((r_idx, c_idx))
        if not sel_idx_set:
            return

        visible_row_indices = []
        seen_r = set()
        for y in range(hdr_h + 2, tree_h, 14):
            iid = self.tree.identify_row(y)
            if iid:
                r_idx = self.kno_to_row_idx.get(iid)
                if r_idx is not None and r_idx not in seen_r:
                    seen_r.add(r_idx)
                    visible_row_indices.append((r_idx, iid))
        iid_bottom = self.tree.identify_row(tree_h - 4)
        if iid_bottom:
            r_idx = self.kno_to_row_idx.get(iid_bottom)
            if r_idx is not None and r_idx not in seen_r:
                seen_r.add(r_idx)
                visible_row_indices.append((r_idx, iid_bottom))

        border_idx = 0
        line_w = 2
        for r_idx, iid in visible_row_indices:
            for c_idx, c_name in enumerate(vis_cols):
                if (r_idx, c_idx) not in sel_idx_set:
                    continue
                bbox = self.tree.bbox(iid, c_name)
                if not bbox:
                    continue
                x, y, w, h = bbox
                if x + w <= 0 or x >= tree_w or y + h <= hdr_h or y >= tree_h:
                    continue
                top_y = max(y, hdr_h)
                vis_h = (y + h) - top_y
                if vis_h <= 1:
                    continue

                if (r_idx - 1, c_idx) not in sel_idx_set and y >= hdr_h:
                    self._get_overlay_border_widget(border_idx).place(x=x, y=y, width=w, height=line_w)
                    border_idx += 1
                if (r_idx + 1, c_idx) not in sel_idx_set:
                    self._get_overlay_border_widget(border_idx).place(x=x, y=y + h - line_w, width=w, height=line_w)
                    border_idx += 1
                if (r_idx, c_idx - 1) not in sel_idx_set:
                    self._get_overlay_border_widget(border_idx).place(x=x, y=top_y, width=line_w, height=vis_h)
                    border_idx += 1
                if (r_idx, c_idx + 1) not in sel_idx_set:
                    self._get_overlay_border_widget(border_idx).place(x=x + w - line_w, y=top_y, width=line_w, height=vis_h)
                    border_idx += 1

    def update_cell_info_bar(self):
        """このペイン内の選択セル数・数値合計・平均を単位種別ごとに正確かつ高速に表示する"""
        row_total = len(self.rows)
        if not self.selected_cells:
            self.cell_info_var.set(f"件数: {row_total:,}件  |  選択: なし（セルクリック/ドラッグで選択してボタンコピー可）")
            return
        cnt = len(self.selected_cells)
        rows_cnt = len({r for r, _ in self.selected_cells})
        cols_cnt = len({c for _, c in self.selected_cells})

        unit = self.app.time_unit_var.get()
        rates_tuple = self.app._get_rates_tuple()
        work_unit_lbl = "円" if unit == "金額(円)" else ("h" if unit == "時間(分÷60)" else ("工数" if unit == "工数(8h=1)" else "分"))
        rows_ref = self.rows
        n_rows = len(rows_ref)
        get_w = self.app._get_work_numeric_val

        # 選択された数値セルを実単位（円 / 工数 / h / 分 / 数量）ごとに集計し、異単位混在時の誤合算を防止
        sums_by_unit = {}
        cnts_by_unit = {}
        for r_idx, c_name in self.selected_cells:
            if 0 <= r_idx < n_rows and c_name in NUMERIC_COLUMNS:
                r_dict = rows_ref[r_idx]
                if c_name in WORK_ONLY_COLUMNS:
                    u_lbl = work_unit_lbl
                    val = get_w(r_dict, c_name, rates_tuple=rates_tuple, unit=unit)
                elif c_name in YEN_ONLY_COLUMNS:
                    u_lbl = "円"
                    if c_name in DERIVED_YEN_COLUMNS:
                        val = get_w(r_dict, c_name, rates_tuple=rates_tuple, unit=unit)
                    else:
                        val = to_float(r_dict.get(c_name, 0.0), 0.0)
                else:
                    u_lbl = "数量"
                    val = to_float(r_dict.get(c_name, 0.0), 0.0)
                sums_by_unit[u_lbl] = sums_by_unit.get(u_lbl, 0.0) + val
                cnts_by_unit[u_lbl] = cnts_by_unit.get(u_lbl, 0) + 1

        def fmt_num(v: float) -> str:
            return f"{int(round(v)):,}" if abs(v - round(v)) < 1e-6 else f"{v:,.2f}"

        if len(sums_by_unit) == 1:
            u_lbl, s_val = next(iter(sums_by_unit.items()))
            avg_val = s_val / max(1, cnts_by_unit[u_lbl])
            self.cell_info_var.set(
                f"件数: {row_total:,}件  |  選択: {rows_cnt:,}行×{cols_cnt}列 ({cnt:,}セル)  |  合計: {fmt_num(s_val)}  平均: {fmt_num(avg_val)} ({u_lbl})"
            )
        elif len(sums_by_unit) > 1:
            parts = []
            for u_lbl in ("円", "工数", "h", "分", "数量"):
                if u_lbl in sums_by_unit:
                    suffix = "" if u_lbl == "数量" else u_lbl
                    title = "数量計" if u_lbl == "数量" else ("金額計" if u_lbl == "円" else f"{u_lbl}計")
                    parts.append(f"{title}: {fmt_num(sums_by_unit[u_lbl])}{suffix}")
            self.cell_info_var.set(
                f"件数: {row_total:,}件  |  選択: {rows_cnt:,}行×{cols_cnt}列 ({cnt:,}セル)  |  " + " / ".join(parts)
            )
        elif cnt == 1:
            r_idx, c_name = next(iter(self.selected_cells))
            val_txt = self.app._get_cell_display_str(rows_ref[r_idx], c_name) if 0 <= r_idx < n_rows else ""
            if len(val_txt) > 30:
                val_txt = val_txt[:27] + "..."
            disp_c = get_display_column_label(c_name, unit)
            self.cell_info_var.set(f"件数: {row_total:,}件  |  選択 [{disp_c}]: {val_txt}")
        else:
            self.cell_info_var.set(f"件数: {row_total:,}件  |  選択: {rows_cnt:,}行×{cols_cnt}列 ({cnt:,}セル)")

    def _compute_rect_cells(self, cell_a: tuple, cell_b: tuple) -> set:
        if not cell_a or not cell_b:
            return set()
        r1, c1_name = cell_a
        r2, c2_name = cell_b
        if c1_name not in self.visible_columns or c2_name not in self.visible_columns:
            return set()
        c1 = self.visible_columns.index(c1_name)
        c2 = self.visible_columns.index(c2_name)
        r_min, r_max = min(r1, r2), max(r1, r2)
        c_min, c_max = min(c1, c2), max(c1, c2)
        result = set()
        for r in range(max(0, r_min), min(len(self.rows) - 1, r_max) + 1):
            for c in range(c_min, c_max + 1):
                result.add((r, self.visible_columns[c]))
        return result

    def _resolve_tree_column(self, col_id_str: str) -> str:
        if not col_id_str or not col_id_str.startswith("#"):
            return ""
        try:
            idx = int(col_id_str[1:]) - 1
            if 0 <= idx < len(self.visible_columns):
                return self.visible_columns[idx]
        except ValueError:
            pass
        return ""

    def _get_header_cell_bounds(self, col_idx: int):
        x = 0
        first_frac, _ = self.tree.xview()
        total_w = sum(self.tree.column(c, "width") for c in self.visible_columns)
        offset = int(round(first_frac * total_w))
        for i, col in enumerate(self.visible_columns):
            cw = self.tree.column(col, "width")
            if i == col_idx:
                return x - offset, x + cw - offset
            x += cw
        return 0, 0

    def _on_tree_select(self, event=None):
        self.app.active_pane = self
        sel = self.tree.selection()
        if sel:
            self.app._on_row_focused(self.pane_id, sel[0])

    def _on_button_release(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "separator":
            self.app._capture_current_column_widths()
            self.app.save_config()
            self._schedule_overlay_update()
            return
        if self._drag_selecting:
            self._drag_selecting = False
            if self.active_cell:
                r_idx, _ = self.active_cell
                if 0 <= r_idx < len(self.rows):
                    self.app._on_row_focused(self.pane_id, self.rows[r_idx]["管理No"])
            self.update_cell_info_bar()
            self._schedule_overlay_update()

    def set_linked_row(self, kno: str):
        """対向ペインと揃えて指定管理Noの行を強調表示する"""
        if self._linked_kno and self.tree.exists(self._linked_kno):
            cur_tags = tuple(t for t in self.tree.item(self._linked_kno, "tags") if t != "linked_row")
            self.tree.item(self._linked_kno, tags=cur_tags)
        self._linked_kno = kno
        if kno and self.tree.exists(kno):
            cur_tags = tuple(t for t in self.tree.item(kno, "tags") if t != "linked_row")
            self.tree.item(kno, tags=cur_tags + ("linked_row",))

    def select_entire_column(self, col_name: str, add_mode: bool = False):
        self.app.active_pane = self
        if not self.rows or col_name not in self.visible_columns:
            return
        col_cells = {(r_idx, col_name) for r_idx in range(len(self.rows))}
        if add_mode:
            if col_cells.issubset(self.selected_cells):
                self.selected_cells -= col_cells
            else:
                self.selected_cells |= col_cells
        else:
            self.selected_cells = col_cells
        self.anchor_cell = (0, col_name)
        self.active_cell = (0, col_name)
        self.tree.selection_remove(self.tree.selection())
        self.update_cell_info_bar()
        self._schedule_overlay_update()

    def _on_left_click(self, event):
        self.app.active_pane = self
        self.tree.focus_set()
        region = self.tree.identify_region(event.x, event.y)
        col_id = self.tree.identify_column(event.x)
        col_name = self._resolve_tree_column(col_id)

        if region == "heading" and col_name:
            col_idx = self.visible_columns.index(col_name)
            _, col_end_x = self._get_header_cell_bounds(col_idx)
            is_ctrl = bool(event.state & 0x0004)
            filter_hit_w = 54 if col_name in self.app.column_filters else 36
            if event.x >= col_end_x - filter_hit_w and col_end_x > 0:
                self.app.open_excel_filter_popup(col_name, event.x_root, self.tree.winfo_rooty() + 30)
                return "break"
            if is_ctrl:
                self.select_entire_column(col_name, add_mode=bool(event.state & 0x0001))
                return "break"
            self.sort_by_column(col_name)
            return "break"

        if region == "cell":
            row_id = self.tree.identify_row(event.y)
            if not row_id or row_id not in self.kno_to_row_idx or not col_name:
                return
            r_idx = self.kno_to_row_idx[row_id]
            cell = (r_idx, col_name)
            is_ctrl = bool(event.state & 0x0004)
            is_shift = bool(event.state & 0x0001)

            if is_shift and self.anchor_cell is not None:
                rect = self._compute_rect_cells(self.anchor_cell, cell)
                self.selected_cells = (self.selected_cells | rect) if is_ctrl else rect
                self.active_cell = cell
                self._drag_selecting = False
            elif is_ctrl:
                if cell in self.selected_cells:
                    self.selected_cells.remove(cell)
                    self._drag_base_cells = set(self.selected_cells)
                else:
                    self._drag_base_cells = set(self.selected_cells)
                    self.selected_cells.add(cell)
                self.anchor_cell = cell
                self.active_cell = cell
                self._drag_selecting = True
                self._drag_ctrl = True
            else:
                self.selected_cells = {cell}
                self._drag_base_cells = set()
                self.anchor_cell = cell
                self.active_cell = cell
                self._drag_selecting = True
                self._drag_ctrl = False

            self.tree.selection_remove(self.tree.selection())
            self.app._on_row_focused(self.pane_id, row_id)
            self.update_cell_info_bar()
            self._schedule_overlay_update()
            return "break"

    def _on_left_drag(self, event):
        if not self._drag_selecting or self.anchor_cell is None:
            return
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        col_name = self._resolve_tree_column(col_id)
        if not row_id or row_id not in self.kno_to_row_idx or not col_name:
            return
        r_idx = self.kno_to_row_idx[row_id]
        cell = (r_idx, col_name)
        if cell == self.active_cell:
            return
        self.active_cell = cell
        rect = self._compute_rect_cells(self.anchor_cell, cell)
        self.selected_cells = (self._drag_base_cells | rect) if self._drag_ctrl else rect
        self.update_cell_info_bar()
        self._schedule_overlay_update()

    def _on_right_click(self, event):
        self.app.active_pane = self
        region = self.tree.identify_region(event.x, event.y)
        col_id = self.tree.identify_column(event.x)
        col_name = self._resolve_tree_column(col_id)

        if region == "heading" and col_name:
            disp_col = get_display_column_label(col_name, self.app.time_unit_var.get())
            self.header_menu.delete(0, tk.END)
            self.header_menu.add_command(
                label=f"「{disp_col}」列をまるごと選択",
                command=lambda c=col_name: self.select_entire_column(c, add_mode=False),
            )
            self.header_menu.add_command(
                label=f"「{disp_col}」列をまるごと追加選択 (Ctrl相当)",
                command=lambda c=col_name: self.select_entire_column(c, add_mode=True),
            )
            self.header_menu.add_separator()
            self.header_menu.add_command(
                label=f"「{disp_col}」列の全データをコピー（見出しなし・Excel貼付用）",
                command=lambda c=col_name: self.copy_columns_to_clipboard([c], include_header=False, raw_numbers=True, selected_only=False),
            )
            self.header_menu.add_command(
                label=f"「{disp_col}」列の全データをコピー（見出し付き・Excel貼付用）",
                command=lambda c=col_name: self.copy_columns_to_clipboard([c], include_header=True, raw_numbers=True, selected_only=False),
            )
            self.header_menu.add_command(
                label="複数の列を選んでコピー（Excel貼付用）...",
                command=self.app.open_column_copy_dialog,
            )
            self.header_menu.add_separator()
            self.header_menu.add_command(
                label=f"「{disp_col}」でExcel風フィルターを開く...",
                command=lambda c=col_name: self.app.open_excel_filter_popup(c, event.x_root, self.tree.winfo_rooty() + 30),
            )
            if col_name in self.app.column_filters:
                self.header_menu.add_command(
                    label=f"「{disp_col}」のフィルターを解除",
                    command=lambda c=col_name: self.app._clear_single_filter(c),
                )
            self.header_menu.add_separator()
            self.header_menu.add_command(label=f"この一覧を「{disp_col}」昇順に並べ替え", command=lambda c=col_name: self.set_sort(c, False))
            self.header_menu.add_command(label=f"この一覧を「{disp_col}」降順に並べ替え", command=lambda c=col_name: self.set_sort(c, True))
            if not (self.pane_id == "calc" and col_name == "管理No"):
                self.header_menu.add_separator()
                self.header_menu.add_command(label="この列を左へ移動", command=lambda c=col_name: self.app.shift_column(c, -1, self.pane_id))
                self.header_menu.add_command(label="この列を右へ移動", command=lambda c=col_name: self.app.shift_column(c, 1, self.pane_id))
                self.header_menu.add_command(label="列の順番・表示設定...", command=self.app.open_column_order_dialog)
            self.header_menu.add_separator()
            self.header_menu.add_command(label="この列の幅を自動調整", command=lambda c=col_name: self.autofit_column(c))
            self.header_menu.tk_popup(event.x_root, event.y_root)
            return

        row_id = self.tree.identify_row(event.y)
        if row_id and row_id in self.kno_to_row_idx:
            r_idx = self.kno_to_row_idx[row_id]
            clicked_cell = (r_idx, col_name) if col_name else None
            if clicked_cell and clicked_cell not in self.selected_cells:
                self.selected_cells = {clicked_cell}
                self.anchor_cell = clicked_cell
                self.active_cell = clicked_cell
                self.tree.selection_remove(self.tree.selection())
                self.app._on_row_focused(self.pane_id, row_id)
                self.update_cell_info_bar()
                self._schedule_overlay_update()

            sel_cnt = len(self.selected_cells)
            sel_rows = len({r for r, _ in self.selected_cells})
            disp_col = get_display_column_label(col_name, self.app.time_unit_var.get()) if col_name else ""
            self.row_menu.delete(0, tk.END)
            if sel_cnt > 0:
                self.row_menu.add_command(
                    label=f"選択中の範囲をコピー ({sel_cnt}セル・Excel貼付用) [Ctrl+C]",
                    command=lambda: self.copy_selection(include_header=False),
                )
                self.row_menu.add_command(
                    label=f"選択中の範囲を見出し付きでコピー ({sel_cnt}セル)",
                    command=lambda: self.copy_selection(include_header=True),
                )
                self.row_menu.add_separator()
            if col_name:
                cell_val = self._get_cell_value_for_copy(row_id, col_name, raw_numbers=True)
                self.row_menu.add_command(
                    label=f"このセル（{disp_col}）の値をコピー",
                    command=lambda v=cell_val: self.app._copy_text(v),
                )
            self.row_menu.add_command(label=f"管理No「{row_id}」をコピー", command=lambda k=row_id: self.app._copy_text(k))
            self.row_menu.add_command(
                label=f"反対側の一覧でも「{row_id}」を表示",
                command=lambda k=row_id: self.app.focus_kno_in_other_pane(self.pane_id, k),
            )
            self.row_menu.add_separator()
            if col_name:
                self.row_menu.add_command(
                    label=f"「{disp_col}」列をまるごと選択",
                    command=lambda c=col_name: self.select_entire_column(c, add_mode=False),
                )
                self.row_menu.add_command(
                    label=f"「{disp_col}」列の全データをコピー（見出しなし）",
                    command=lambda c=col_name: self.copy_columns_to_clipboard([c], include_header=False, raw_numbers=True, selected_only=False),
                )
                self.row_menu.add_command(
                    label=f"「{disp_col}」列の全データをコピー（見出し付き）",
                    command=lambda c=col_name: self.copy_columns_to_clipboard([c], include_header=True, raw_numbers=True, selected_only=False),
                )
            self.row_menu.add_command(
                label=f"選択行 ({sel_rows}件) のこの一覧の全列をコピー",
                command=lambda: self.copy_columns_to_clipboard(self.visible_columns, include_header=False, raw_numbers=True, selected_only=True),
            )
            self.row_menu.add_command(
                label="複数の列を選んでコピー（左右またぎ可）...",
                command=self.app.open_column_copy_dialog,
            )
            if col_name:
                self.row_menu.add_separator()
                self.row_menu.add_command(
                    label=f"「{disp_col}」列でフィルター...",
                    command=lambda c=col_name: self.app.open_excel_filter_popup(c, event.x_root, event.y_root),
                )
                self.row_menu.add_command(label=f"「{disp_col}」列の幅を自動調整", command=lambda c=col_name: self.autofit_column(c))
            self.row_menu.add_separator()
            self.row_menu.add_command(label="この一覧の列幅を自動調整", command=self.autofit_all_columns)
            self.row_menu.tk_popup(event.x_root, event.y_root)

    def select_all_cells(self):
        self.app.active_pane = self
        if not self.rows or not self.visible_columns:
            return "break"
        self.selected_cells = {
            (r_idx, col)
            for r_idx in range(len(self.rows))
            for col in self.visible_columns
        }
        self.anchor_cell = (0, self.visible_columns[0])
        self.active_cell = (0, self.visible_columns[0])
        self.tree.selection_remove(self.tree.selection())
        self.update_cell_info_bar()
        self._schedule_overlay_update()
        return "break"

    def clear_selection(self):
        """選択セルと枠線を解除する"""
        self.selected_cells.clear()
        self.anchor_cell = None
        self.active_cell = None
        self._drag_selecting = False
        try:
            self.tree.selection_remove(self.tree.selection())
        except Exception:
            pass
        self._clear_overlay_borders()
        self.update_cell_info_bar()
        return "break"

    def _get_cell_value_for_copy(self, kno: str, col: str, raw_numbers: bool = True) -> str:
        r_idx = self.kno_to_row_idx.get(kno)
        row_dict = self.rows[r_idx] if (r_idx is not None and r_idx < len(self.rows)) else None
        if row_dict is None:
            for r in self.rows:
                if r["管理No"] == kno:
                    row_dict = r
                    break
        if row_dict is None:
            return ""
        if raw_numbers:
            if col in RATE_DEPENDENT_COLUMNS:
                val = self.app._get_work_numeric_val(row_dict, col)
                return str(int(round(val))) if abs(val - round(val)) < 1e-6 else f"{val:.2f}"
            elif col in COST_YEN_COLUMNS:
                return str(int(round(row_dict.get(col, 0.0))))
        return self.app._get_cell_display_str(row_dict, col)

    def copy_selection(self, include_header: bool = False, raw_numbers: bool = True):
        """この一覧で選択中のセル・範囲・列をTSVでクリップボードへコピーする"""
        self.app.active_pane = self
        if not self.selected_cells:
            if self.tree.selection():
                self.copy_columns_to_clipboard(self.visible_columns, include_header=include_header, raw_numbers=raw_numbers, selected_only=True)
                return "break"
            messagebox.showinfo("選択なし", "コピーするセルをクリックまたはドラッグで選択してください。", parent=self.app.root)
            return "break"

        rows_sorted = sorted({r for r, _ in self.selected_cells if 0 <= r < len(self.rows)})
        cols_sorted = [c for c in self.visible_columns if any(col == c for _, col in self.selected_cells)]
        if not rows_sorted or not cols_sorted:
            return "break"

        lines = []
        unit_now = self.app.time_unit_var.get()
        rates_tuple = self.app._get_rates_tuple()
        get_w = self.app._get_work_numeric_val
        get_disp = self.app._get_cell_display_str
        if include_header:
            lines.append("\t".join(escape_tsv_cell(get_display_column_label(c, unit_now)) for c in cols_sorted))
        for r_idx in rows_sorted:
            r_dict = self.rows[r_idx]
            row_cells = []
            for col in cols_sorted:
                if (r_idx, col) not in self.selected_cells:
                    row_cells.append("")
                    continue
                if raw_numbers:
                    if col in RATE_DEPENDENT_COLUMNS:
                        val = get_w(r_dict, col, rates_tuple=rates_tuple, unit=unit_now)
                        cell_str = str(int(round(val))) if abs(val - round(val)) < 1e-6 else f"{val:.2f}"
                    elif col in COST_YEN_COLUMNS:
                        cell_str = str(int(round(r_dict.get(col, 0.0))))
                    else:
                        cell_str = str(r_dict.get(col, ""))
                else:
                    cell_str = get_disp(r_dict, col, unit=unit_now, rates_tuple=rates_tuple)
                row_cells.append(escape_tsv_cell(cell_str))
            lines.append("\t".join(row_cells))

        self.app.root.clipboard_clear()
        self.app.root.clipboard_append("\r\n".join(lines))
        self.app.status_var.set(
            f"コピーしました（{len(rows_sorted):,}行 × {len(cols_sorted)}列 / {len(self.selected_cells):,}セル・Excelへ Ctrl+V で貼付可能）"
        )
        return "break"

    def copy_active_columns_action(self):
        """選択中セルが属する列の全表示行をまとめてコピー（未選択時は列選択ダイアログを開く）"""
        self.app.active_pane = self
        if not self.rows:
            messagebox.showinfo("コピー対象なし", "コピーするデータがありません。先に集計を実行してください。", parent=self.app.root)
            return
        cols_sorted = [c for c in self.visible_columns if any(col == c for _, col in self.selected_cells)]
        if not cols_sorted and self.active_cell and self.active_cell[1] in self.visible_columns:
            cols_sorted = [self.active_cell[1]]
        if not cols_sorted:
            self.app.open_column_copy_dialog()
            return
        self.copy_columns_to_clipboard(cols_sorted, include_header=True, raw_numbers=True, selected_only=False)

    def copy_columns_to_clipboard(self, columns: list, include_header: bool = False, raw_numbers: bool = True, selected_only: bool = False):
        if not self.rows:
            messagebox.showinfo("コピー対象なし", "コピーするデータがありません。先に集計を実行してください。", parent=self.app.root)
            return
        if not columns:
            return
        if selected_only:
            sel_row_indices = sorted({r for r, _ in self.selected_cells if 0 <= r < len(self.rows)})
            if sel_row_indices:
                target_rows = [self.rows[i] for i in sel_row_indices]
            else:
                sel_set = set(self.tree.selection())
                target_rows = [r for r in self.rows if r["管理No"] in sel_set]
            if not target_rows:
                messagebox.showinfo("選択なし", "コピーする行またはセルを選択してください。", parent=self.app.root)
                return
        else:
            target_rows = self.rows

        lines = []
        unit_now = self.app.time_unit_var.get()
        rates_tuple = self.app._get_rates_tuple()
        get_w = self.app._get_work_numeric_val
        get_disp = self.app._get_cell_display_str
        if include_header:
            lines.append("\t".join(escape_tsv_cell(get_display_column_label(c, unit_now)) for c in columns))
        for r in target_rows:
            row_cells = []
            for col in columns:
                if raw_numbers:
                    if col in RATE_DEPENDENT_COLUMNS:
                        val = get_w(r, col, rates_tuple=rates_tuple, unit=unit_now)
                        cell_str = str(int(round(val))) if abs(val - round(val)) < 1e-6 else f"{val:.2f}"
                    elif col in COST_YEN_COLUMNS:
                        cell_str = str(int(round(r.get(col, 0.0))))
                    else:
                        cell_str = str(r.get(col, ""))
                else:
                    cell_str = get_disp(r, col, unit=unit_now, rates_tuple=rates_tuple)
                row_cells.append(escape_tsv_cell(cell_str))
            lines.append("\t".join(row_cells))

        self.app.root.clipboard_clear()
        self.app.root.clipboard_append("\r\n".join(lines))
        col_label = get_display_column_label(columns[0], unit_now) if len(columns) == 1 else f"{len(columns)}列"
        self.app.status_var.set(f"クリップボードにコピーしました（{col_label} × {len(target_rows):,}行・Excelへ Ctrl+V で貼付可能）")

    def sort_by_column(self, col: str):
        if self.app.link_rows_var.get():
            if self.app.sort_col == col:
                new_desc = not self.app.sort_desc
            else:
                new_desc = col in NUMERIC_COLUMNS
            self.app.set_sort(col, new_desc)
            return
        if self.sort_col == col:
            self.sort_desc = not self.sort_desc
        else:
            self.sort_col = col
            self.sort_desc = col in NUMERIC_COLUMNS
        self.refresh_rows(self.app.displayed_rows)

    def set_sort(self, col: str, desc: bool, sync: bool = True):
        if sync and self.app.link_rows_var.get():
            self.app.set_sort(col, desc)
            return
        self.sort_col = col
        self.sort_desc = desc
        self.refresh_rows(self.app.displayed_rows)

    def autofit_column(self, col_name: str, save_after: bool = True):
        if col_name not in self.visible_columns:
            return
        disp_name = get_display_column_label(col_name, self.app.time_unit_var.get())
        max_chars = len(disp_name) * 2 + 5
        for r in self.rows[:300]:
            s = self.app._get_cell_display_str(r, col_name)
            char_len = sum(2 if ord(ch) > 127 else 1 for ch in s)
            if char_len > max_chars:
                max_chars = char_len
        width = min(max(max_chars * 7 + 16, 65), 380)
        self.app.column_widths[self._width_key(col_name)] = width
        self.tree.column(col_name, width=width)
        if save_after:
            self.app.save_config()

    def autofit_all_columns(self):
        for col in self.visible_columns:
            self.autofit_column(col, save_after=False)
        self.app.save_config()
        self._schedule_overlay_update()
        self.app.status_var.set("列幅を自動調整しました。")

    def scroll_to_kno(self, kno: str):
        if kno in self.kno_to_row_idx and self.tree.exists(kno):
            self.tree.see(kno)
            self.tree.selection_set(kno)
            r_idx = self.kno_to_row_idx[kno]
            if self.visible_columns:
                self.selected_cells = {(r_idx, self.visible_columns[0])}
                self.anchor_cell = (r_idx, self.visible_columns[0])
                self.active_cell = (r_idx, self.visible_columns[0])
                self.update_cell_info_bar()
                self._schedule_overlay_update()

    def refresh_rows(self, base_rows: list, pre_sorted: bool = False):
        """並び順（行リンクON時は左右共通）を適用してTreeviewを高速更新する"""
        self.update_heading_texts()
        if pre_sorted:
            sorted_rows = list(base_rows)
        else:
            if self.app.link_rows_var.get():
                col = self.app.sort_col
                desc = self.app.sort_desc
            else:
                col = self.sort_col if self.sort_col in self.visible_columns else (self.visible_columns[0] if self.visible_columns else "管理No")
                desc = self.sort_desc

            if col in RATE_DEPENDENT_COLUMNS:
                get_w = self.app._get_work_numeric_val
                rates_tuple = self.app._get_rates_tuple()
                cur_unit = self.app.time_unit_var.get()
                sort_key = lambda item: get_w(item, col, rates_tuple=rates_tuple, unit=cur_unit)
            elif col in COST_YEN_COLUMNS:
                sort_key = lambda item: item.get(col, 0.0)
            elif col in NUMERIC_COLUMNS:
                sort_key = lambda item: to_float(item.get(col, 0.0), 0.0)
            else:
                sort_key = lambda item: str(item.get(col, ""))

            sorted_rows = list(base_rows)
            sorted_rows.sort(key=sort_key, reverse=desc)

        self.rows = sorted_rows
        self.kno_to_row_idx = {r["管理No"]: idx for idx, r in enumerate(self.rows)}
        self._linked_kno = None

        self.selected_cells.clear()
        self.anchor_cell = None
        self.active_cell = None
        self._drag_selecting = False
        self._clear_overlay_borders()
        self.update_cell_info_bar()

        self._render_job_id += 1
        if self._chunk_after_id is not None:
            try:
                self.app.root.after_cancel(self._chunk_after_id)
            except Exception:
                pass
            self._chunk_after_id = None
        render_id = self._render_job_id
        self.tree.delete(*self.tree.get_children())

        total_rows = len(self.rows)
        if total_rows <= 600:
            self._insert_chunk(0, total_rows, render_id)
        else:
            # 最初の画面表示分（150行）を即座に描画し、残りは1200行単位で高速バッチ挿入
            self._insert_chunk(0, 150, render_id)

    def _insert_chunk(self, start_idx: int, chunk_size: int, render_id: int):
        self._chunk_after_id = None
        if self.app._closing or render_id != self._render_job_id:
            return
        try:
            if not self.tree.winfo_exists():
                return
        except Exception:
            return
        rows_ref = self.rows
        total_len = len(rows_ref)
        end_idx = min(start_idx + chunk_size, total_len)
        vis_cols = self.visible_columns
        get_disp = self.app._get_cell_display_str
        cur_unit = self.app.time_unit_var.get()
        rates_tuple = self.app._get_rates_tuple()
        tree_insert = self.tree.insert
        end_pos = tk.END

        for idx in range(start_idx, end_idx):
            r = rows_ref[idx]
            values = [get_disp(r, c_name, unit=cur_unit, rates_tuple=rates_tuple) for c_name in vis_cols]
            tag = "active_both" if (r["作業小計"] > 0 and r["購入品小計"] > 0) else ("even" if (idx & 1) == 0 else "odd")
            tree_insert("", end_pos, iid=r["管理No"], values=values, tags=(tag,))

        if end_idx < total_len and not self.app._closing:
            try:
                self._chunk_after_id = self.app.root.after(1, lambda: self._insert_chunk(end_idx, 1200, render_id))
            except Exception:
                self._chunk_after_id = None
        else:
            if self.app.link_rows_var.get() and self.app._current_detail_kno:
                self.set_linked_row(self.app._current_detail_kno)
            self._schedule_overlay_update()


class DenpyoHubApp:
    """制作伝票一覧作成メインアプリケーションクラス (DenpyoHub_Miya)"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1420x880")
        self.root.minsize(1100, 700)
        self.root.configure(bg=COLORS["bg_main"])

        # アイコン設定
        icon_path = get_resource_path("icon.ico")
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except Exception:
                pass

        # 終了・非同期ジョブ管理フラグ
        self._closing = False
        self._active_job_id = 0
        self._is_loading = False
        self._ui_queue = queue.Queue()
        self._syncing_yscroll = False
        self._last_saved_config_json = ""
        self._last_search_kw = ""
        self.last_errors = []
        self._poll_after_id = None
        self._max_after_id = None
        self._startup_after_id = None
        self._sash_after_id = None
        self._restore_y_after_id = None
        self._search_debounce_job = None
        self._progress_hide_job = None

        # 設定ファイルの管理（デフォルトはアプリと同じ場所・旧ファイル名からの移行読込に対応）
        self.default_config_path = os.path.join(get_base_dir(), CONFIG_FILENAME)
        self.config_path = self.default_config_path
        self.main_file_cache = MainFileParseCache(max_entries=32)
        self.seizo_cache = SeizoCacheStore(os.path.join(get_base_dir(), SEIZO_CACHE_FILENAME))
        if os.path.exists(self.default_config_path):
            self.config = self.load_config(self.default_config_path)
        else:
            legacy_path = os.path.join(get_base_dir(), LEGACY_CONFIG_FILENAME)
            self.config = self.load_config(legacy_path) if os.path.exists(legacy_path) else {}

        # 参照ファイルリスト（複数選択対応 + 制作伝票フォルダ）
        self.file_lists = {
            "daily": self._normalize_file_list(self.config.get("files", {}).get("daily", [])),
            "kounyu": self._normalize_file_list(self.config.get("files", {}).get("kounyu", [])),
            "gaichu": self._normalize_file_list(self.config.get("files", {}).get("gaichu", [])),
            "juchu": self._normalize_file_list(self.config.get("files", {}).get("juchu", [])),
            "seizo": self._normalize_file_list(self.config.get("files", {}).get("seizo", [])),
        }

        # UI表示用の要約文字列と更新日変数
        self.path_summary_vars = {k: tk.StringVar() for k in ("daily", "kounyu", "gaichu", "juchu", "seizo")}
        self.mtime_vars = {k: tk.StringVar(value="未選択") for k in ("daily", "kounyu", "gaichu", "juchu", "seizo")}

        # 作業コード範囲設定
        saved_ranges = self.config.get("code_ranges", {})
        self.code_ranges = {cat: saved_ranges.get(cat, def_val) for cat, def_val in DEFAULT_CODE_RANGES.items()}

        # 作業単価レート設定（円/分）
        saved_rates = self.config.get("work_rates", {})
        self.work_rates = {}
        for cat, def_rate in DEFAULT_WORK_RATES.items():
            try:
                val = float(saved_rates.get(cat, def_rate))
                self.work_rates[cat] = val if (math.isfinite(val) and val >= 0) else def_rate
            except (TypeError, ValueError):
                self.work_rates[cat] = def_rate

        # 列の順番・表示・幅設定（小計1・小計2の内訳列はデフォルトでは非表示にし、ボタンで展開可能）
        raw_saved_order = self.config.get("column_order", list(DEFAULT_COLUMN_ORDER))
        saved_order_set = set(raw_saved_order)
        self.column_order = [c for c in raw_saved_order if c in ALL_AVAILABLE_COLUMNS]
        for c in ALL_AVAILABLE_COLUMNS:
            if c not in self.column_order:
                self.column_order.append(c)

        has_sub_detail_flags = ("show_sub1_details" in self.config) or ("show_sub2_details" in self.config)
        self.show_sub1_details = bool(self.config.get("show_sub1_details", False))
        self.show_sub2_details = bool(self.config.get("show_sub2_details", False))

        saved_visible = self.config.get("visible_columns", list(DEFAULT_VISIBLE_COLUMNS))
        self.visible_set = {c for c in saved_visible if c in ALL_AVAILABLE_COLUMNS}
        if not self.visible_set:
            self.visible_set = set(DEFAULT_VISIBLE_COLUMNS)
        for c in ALL_AVAILABLE_COLUMNS:
            if c not in saved_order_set and c not in DEFAULT_HIDDEN_COLUMNS:
                self.visible_set.add(c)
        if not has_sub_detail_flags:
            self.visible_set.difference_update(SUBTOTAL1_DETAIL_COLUMNS)
            self.visible_set.difference_update(SUBTOTAL2_DETAIL_COLUMNS)
        self.visible_columns = [c for c in self.column_order if c in self.visible_set]

        saved_widths = self.config.get("column_widths", {})
        self.column_widths = dict(DEFAULT_COL_WIDTHS)
        for k, v in saved_widths.items():
            if isinstance(v, (int, float)) and 35 <= int(v) <= 800:
                self.column_widths[k] = int(v)

        # フィルタ・表示設定変数（デフォルトは「金額(円)」・設定で変更可能）
        self.search_var = tk.StringVar(value="")
        self.filter_mode_var = tk.StringVar(value=self.config.get("filter_mode", "実績ありのみ"))
        saved_def_unit = self.config.get("default_time_unit", DEFAULT_TIME_UNIT)
        if saved_def_unit not in VALID_TIME_UNITS:
            saved_def_unit = DEFAULT_TIME_UNIT
        self.default_time_unit = saved_def_unit
        self.time_unit_var = tk.StringVar(value=self.default_time_unit)
        self.code_summary_var = tk.StringVar(value="")
        self.filter_status_var = tk.StringVar(value="列フィルター: なし")
        self.status_var = tk.StringVar(value="起動しています...")
        self.progress_msg_var = tk.StringVar(value="")
        self.progress_pct_var = tk.StringVar(value="0%")
        self.detail_state_var = tk.StringVar(value="")
        self.summary_var = tk.StringVar(value="表示件数: 0 件")
        self.config_info_var = tk.StringVar(value=f"設定ファイル: {self.config_path}")

        # 列ごとのExcel風フィルター条件
        self.column_filters = {}

        # 左右分割ペイン・行リンク連動・分割比率管理
        self.active_pane = None
        self._current_detail_kno = None
        self.unit_buttons = {}
        self.link_rows_var = tk.BooleanVar(value=bool(self.config.get("link_rows", True)))
        self.h_sash_ratio = min(max(to_float(self.config.get("h_sash_ratio", 0.46), 0.46), 0.20), 0.80)
        self.v_sash_ratio = min(max(to_float(self.config.get("v_sash_ratio", 0.66), 0.66), 0.35), 0.85)

        # ファイル設定エリア開閉状態（通常は閉じており、ボタンを押したときだけ開く）
        self.file_panel_expanded = False

        # 内部保持データ
        self.raw_daily = []
        self.raw_kounyu = []
        self.raw_gaichu = []
        self.raw_juchu = []
        self.raw_seizo = []
        self.daily_by_kno = defaultdict(list)
        self.kounyu_by_kno = defaultdict(list)
        self.gaichu_by_kno = defaultdict(list)
        self.juchu_by_kno = {}
        self.seizo_by_kno = defaultdict(list)
        self.aggregated_rows = []
        self.displayed_rows = []
        saved_sort_col = self.config.get("sort_col", "管理No")
        self.sort_col = saved_sort_col if saved_sort_col in ALL_AVAILABLE_COLUMNS else "管理No"
        self.sort_desc = bool(self.config.get("sort_desc", True))
        self._linked_kno = None

        self._setup_light_style()
        self._build_ui()
        if not self.file_panel_expanded:
            self.file_group.pack_forget()
            self.btn_toggle_files.configure(text="▼ ファイル設定を開く")
        self._refresh_file_display_vars(skip_io=True)
        self._update_code_summary_label()
        self._bind_global_mousewheel()

        # 終了時にも列幅・設定を保存
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # UI描画後に最大化およびバックグラウンド起動初期化を実行（メインスレッドを絶対にブロックしない）
        self._poll_after_id = self.root.after(40, self._poll_ui_queue)
        self._max_after_id = self.root.after(20, self._maximize_window)
        self._startup_after_id = self.root.after(120, self._startup_async_init)

    def _cancel_after(self, attr_name: str):
        """指定属性に保持された after タイマーを安全にキャンセルしてクリアする"""
        job = getattr(self, attr_name, None)
        if job is not None:
            try:
                self.root.after_cancel(job)
            except Exception:
                pass
            setattr(self, attr_name, None)

    def _poll_ui_queue(self):
        """ワーカースレッドからのUI更新タスクをメインスレッドで安全に消化する"""
        self._poll_after_id = None
        if self._closing:
            return
        try:
            while True:
                fn = self._ui_queue.get_nowait()
                try:
                    fn()
                except Exception:
                    pass
        except queue.Empty:
            pass
        if not self._closing:
            self._poll_after_id = self.root.after(40, self._poll_ui_queue)

    def _maximize_window(self):
        """ウィンドウを最大化して起動する"""
        self._max_after_id = None
        if self._closing:
            return
        try:
            if sys.platform == "win32":
                self.root.state("zoomed")
            else:
                self.root.attributes("-zoomed", True)
        except Exception:
            pass

    @staticmethod
    def _normalize_file_list(val) -> list:
        """文字列またはリストのファイル設定を重複のないリストに正規化する"""
        if not val:
            return []
        if isinstance(val, str):
            items = [val.strip()] if val.strip() else []
        elif isinstance(val, list):
            items = [str(x).strip() for x in val if str(x).strip()]
        else:
            items = []
        seen = set()
        result = []
        for p in items:
            norm = os.path.normpath(p)
            key = os.path.normcase(os.path.abspath(norm))
            if key not in seen:
                seen.add(key)
                result.append(norm)
        return result

    def _setup_light_style(self):
        """システム設定に関わらず常に清潔なライトモードにするスタイル定義"""
        style = ttk.Style(self.root)
        style.theme_use("clam")

        style.configure(".", background=COLORS["bg_main"], foreground=COLORS["fg_main"], font=("Yu Gothic UI", 10))
        style.configure("TFrame", background=COLORS["bg_main"])
        style.configure("Card.TFrame", background=COLORS["bg_card"], relief="solid", borderwidth=1)
        style.configure("Flat.TFrame", background=COLORS["bg_card"])
        style.configure("TLabel", background=COLORS["bg_main"], foreground=COLORS["fg_main"], font=("Yu Gothic UI", 10))
        style.configure("Card.TLabel", background=COLORS["bg_card"], foreground=COLORS["fg_main"], font=("Yu Gothic UI", 10))
        style.configure("Title.TLabel", background=COLORS["bg_card"], foreground=COLORS["accent"], font=("Yu Gothic UI", 12, "bold"))
        style.configure("Mtime.TLabel", background=COLORS["bg_card"], foreground=COLORS["fg_sub"], font=("Yu Gothic UI", 9))
        style.configure("Badge.TLabel", background=COLORS["bg_header"], foreground=COLORS["accent"], font=("Yu Gothic UI", 9, "bold"))
        style.configure("FilterActive.TLabel", background=COLORS["warning_bg"], foreground="#92400E", font=("Yu Gothic UI", 9, "bold"))

        style.configure(
            "TLabelframe",
            background=COLORS["bg_card"],
            foreground=COLORS["fg_main"],
            bordercolor=COLORS["border"],
            lightcolor=COLORS["border"],
            darkcolor=COLORS["border"],
        )
        style.configure(
            "TLabelframe.Label",
            background=COLORS["bg_card"],
            foreground=COLORS["accent"],
            font=("Yu Gothic UI", 10, "bold"),
        )

        style.configure(
            "TButton",
            background="#E2E8F0",
            foreground=COLORS["fg_main"],
            bordercolor=COLORS["border"],
            lightcolor="#E2E8F0",
            darkcolor="#E2E8F0",
            focuscolor="none",
            padding=(10, 4),
            font=("Yu Gothic UI", 9),
        )
        style.map(
            "TButton",
            background=[("active", "#CBD5E1"), ("pressed", "#94A3B8")],
            foreground=[("disabled", "#94A3B8")],
        )

        style.configure(
            "Primary.TButton",
            background=COLORS["accent"],
            foreground="#FFFFFF",
            bordercolor=COLORS["accent"],
            lightcolor=COLORS["accent"],
            darkcolor=COLORS["accent"],
            focuscolor="none",
            padding=(13, 5),
            font=("Yu Gothic UI", 10, "bold"),
        )
        style.map(
            "Primary.TButton",
            background=[("active", COLORS["accent_hover"]), ("pressed", "#1E40AF")],
            foreground=[("active", "#FFFFFF")],
        )

        style.configure(
            "Success.TButton",
            background=COLORS["success"],
            foreground="#FFFFFF",
            bordercolor=COLORS["success"],
            lightcolor=COLORS["success"],
            darkcolor=COLORS["success"],
            focuscolor="none",
            padding=(12, 5),
            font=("Yu Gothic UI", 9, "bold"),
        )
        style.map(
            "Success.TButton",
            background=[("active", COLORS["success_hover"]), ("pressed", "#166534")],
            foreground=[("active", "#FFFFFF")],
        )

        style.configure(
            "Danger.TButton",
            background=COLORS["danger"],
            foreground="#FFFFFF",
            bordercolor=COLORS["danger"],
            lightcolor=COLORS["danger"],
            darkcolor=COLORS["danger"],
            focuscolor="none",
            padding=(11, 4),
            font=("Yu Gothic UI", 9, "bold"),
        )
        style.map(
            "Danger.TButton",
            background=[("active", "#B91C1C"), ("pressed", "#991B1B")],
            foreground=[("active", "#FFFFFF")],
        )

        style.configure(
            "TEntry",
            fieldbackground="#FFFFFF",
            background="#FFFFFF",
            foreground=COLORS["fg_main"],
            bordercolor=COLORS["border"],
            lightcolor="#FFFFFF",
            darkcolor=COLORS["border"],
            insertcolor=COLORS["fg_main"],
            padding=4,
        )

        style.configure(
            "TCombobox",
            fieldbackground="#FFFFFF",
            background="#E2E8F0",
            foreground=COLORS["fg_main"],
            selectbackground=COLORS["accent_light"],
            selectforeground=COLORS["fg_main"],
            bordercolor=COLORS["border"],
            arrowcolor=COLORS["fg_main"],
            padding=3,
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", "#FFFFFF")],
            selectbackground=[("readonly", "#FFFFFF")],
            selectforeground=[("readonly", COLORS["fg_main"])],
        )

        style.configure(
            "Treeview",
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
            foreground=COLORS["fg_main"],
            rowheight=26,
            bordercolor=COLORS["border"],
            lightcolor=COLORS["border"],
            darkcolor=COLORS["border"],
            font=("Yu Gothic UI", 9),
        )
        style.configure(
            "Treeview.Heading",
            background="#E2E8F0",
            foreground=COLORS["fg_main"],
            bordercolor=COLORS["border"],
            lightcolor="#F1F5F9",
            darkcolor="#CBD5E1",
            font=("Yu Gothic UI", 9, "bold"),
            padding=(6, 5),
        )
        style.map(
            "Treeview",
            background=[("selected", COLORS["accent_light"])],
            foreground=[("selected", COLORS["fg_main"])],
        )
        style.map(
            "Treeview.Heading",
            background=[("active", "#CBD5E1")],
        )

        # スクロールバーを太く掴みやすく設定
        style.configure("Vertical.TScrollbar", width=16, background="#CBD5E1", troughcolor="#F1F5F9", arrowsize=14)
        style.configure("Horizontal.TScrollbar", width=16, background="#CBD5E1", troughcolor="#F1F5F9", arrowsize=14)

        style.configure("TNotebook", background=COLORS["bg_main"], bordercolor=COLORS["border"])
        style.configure(
            "TNotebook.Tab",
            background="#E2E8F0",
            foreground=COLORS["fg_sub"],
            padding=(12, 4),
            font=("Yu Gothic UI", 9, "bold"),
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", "#FFFFFF"), ("active", "#F1F5F9")],
            foreground=[("selected", COLORS["accent"])],
        )

    def _build_ui(self):
        """メイン画面のUI構築（下部バーやボタンが絶対に途切れないよう順序と余白を最適化）"""
        # 1. 上部ヘッダーバー
        header_frame = ttk.Frame(self.root, style="Card.TFrame", padding=(12, 6))
        header_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(6, 3))

        ttk.Label(
            header_frame,
            text="制作伝票一覧 集計システム (DenpyoHub_Miya)",
            style="Title.TLabel",
        ).pack(side=tk.LEFT)

        ttk.Button(header_frame, text="使い方・ヘルプ", command=self.open_help_dialog).pack(side=tk.RIGHT, padx=(3, 0))
        ttk.Button(header_frame, text="設定保存 / 読込", command=self.open_config_manager_dialog).pack(side=tk.RIGHT, padx=3)
        ttk.Button(header_frame, text="列の順番・表示", command=self.open_column_order_dialog).pack(side=tk.RIGHT, padx=3)
        ttk.Button(header_frame, text="表示・レート設定", command=self.open_rate_settings_dialog).pack(side=tk.RIGHT, padx=3)
        ttk.Button(header_frame, text="作業コード設定", command=self.open_code_settings_dialog).pack(side=tk.RIGHT, padx=3)

        self.btn_toggle_files = ttk.Button(header_frame, text="▼ ファイル設定を開く", command=self.toggle_file_panel)
        self.btn_toggle_files.pack(side=tk.RIGHT, padx=3)

        # 2. ファイル選択＆更新日表示エリア（通常は閉じており、ボタンを押したときだけ展開）
        self.file_group = ttk.LabelFrame(
            self.root,
            text="参照ファイル設定（複数選択可・制作伝票フォルダは指定フォルダ直下のみ参照・すべて読み取り専用保護）",
            padding=(10, 4),
        )

        file_defs = [
            ("daily", "日報ファイル (daily_report.csv)", [("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]),
            ("kounyu", "購入品ファイル (購入品.xlsm)", [("Excelファイル", "*.xlsm *.xlsx"), ("すべてのファイル", "*.*")]),
            ("gaichu", "外注加工ファイル (外注加工.xlsm)", [("Excelファイル", "*.xlsm *.xlsx"), ("すべてのファイル", "*.*")]),
            ("juchu", "受注表ファイル (受注票頚城.xlsm)", [("Excelファイル", "*.xlsm *.xlsx"), ("すべてのファイル", "*.*")]),
        ]

        for row_idx, (key, label_text, filetypes) in enumerate(file_defs):
            ttk.Label(self.file_group, text=label_text, style="Card.TLabel", width=27).grid(
                row=row_idx, column=0, sticky=tk.W, padx=(4, 6), pady=2
            )
            entry = ttk.Entry(self.file_group, textvariable=self.path_summary_vars[key], state="readonly")
            entry.grid(row=row_idx, column=1, sticky=tk.EW, padx=4, pady=2)
            entry.bind("<Double-1>", lambda e, k=key, l=label_text, ft=filetypes: self.open_file_list_dialog(k, l, ft))

            btn_box = ttk.Frame(self.file_group, style="Flat.TFrame")
            btn_box.grid(row=row_idx, column=2, padx=2, pady=2)

            ttk.Button(
                btn_box,
                text="選択(複数可)",
                width=12,
                command=lambda k=key, ft=filetypes: self.select_files_replace(k, ft),
            ).pack(side=tk.LEFT, padx=1)

            ttk.Button(
                btn_box,
                text="追加",
                width=6,
                command=lambda k=key, ft=filetypes: self.select_files_append(k, ft),
            ).pack(side=tk.LEFT, padx=1)

            ttk.Button(
                btn_box,
                text="一覧",
                width=6,
                command=lambda k=key, l=label_text, ft=filetypes: self.open_file_list_dialog(k, l, ft),
            ).pack(side=tk.LEFT, padx=1)

            ttk.Button(
                btn_box,
                text="クリア",
                width=6,
                command=lambda k=key: self.clear_files(k),
            ).pack(side=tk.LEFT, padx=1)

            mtime_lbl = ttk.Label(
                self.file_group,
                textvariable=self.mtime_vars[key],
                style="Mtime.TLabel",
            )
            mtime_lbl.grid(row=row_idx, column=3, sticky=tk.W, padx=(8, 6), pady=2)

        # 5行目: 制作伝票フォルダ (指定フォルダ直下の .xls / .xlsx / .xlsm・サブフォルダ除く)
        seizo_row = len(file_defs)
        ttk.Label(self.file_group, text="制作伝票フォルダ (.xls/.xlsx/.xlsm)", style="Card.TLabel", width=27).grid(
            row=seizo_row, column=0, sticky=tk.W, padx=(4, 6), pady=2
        )
        seizo_entry = ttk.Entry(self.file_group, textvariable=self.path_summary_vars["seizo"], state="readonly")
        seizo_entry.grid(row=seizo_row, column=1, sticky=tk.EW, padx=4, pady=2)
        seizo_entry.bind("<Double-1>", lambda e: self.open_seizo_file_list_dialog())

        s_btn_box = ttk.Frame(self.file_group, style="Flat.TFrame")
        s_btn_box.grid(row=seizo_row, column=2, padx=2, pady=2)

        ttk.Button(
            s_btn_box,
            text="選択(複数可)",
            width=12,
            command=self.select_seizo_folder_replace,
        ).pack(side=tk.LEFT, padx=1)

        ttk.Button(
            s_btn_box,
            text="追加",
            width=6,
            command=self.select_seizo_folder_append,
        ).pack(side=tk.LEFT, padx=1)

        ttk.Button(
            s_btn_box,
            text="一覧",
            width=6,
            command=self.open_seizo_file_list_dialog,
        ).pack(side=tk.LEFT, padx=1)

        ttk.Button(
            s_btn_box,
            text="クリア",
            width=6,
            command=lambda: self.clear_files("seizo"),
        ).pack(side=tk.LEFT, padx=1)

        ttk.Label(
            self.file_group,
            textvariable=self.mtime_vars["seizo"],
            style="Mtime.TLabel",
        ).grid(row=seizo_row, column=3, sticky=tk.W, padx=(8, 6), pady=2)

        self.file_group.columnconfigure(1, weight=1, minsize=180)
        self.file_group.columnconfigure(3, minsize=350)

        # 3. 操作・フィルタ・作業コード範囲確認バー
        self.ctrl_frame = ttk.Frame(self.root, style="Card.TFrame", padding=(10, 5))
        self.ctrl_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=2)

        top_ctrl = ttk.Frame(self.ctrl_frame, style="Flat.TFrame")
        top_ctrl.pack(fill=tk.X, pady=(0, 3))

        self.btn_aggregate = ttk.Button(
            top_ctrl,
            text="集計・一覧更新",
            style="Primary.TButton",
            command=self.run_aggregation,
        )
        self.btn_aggregate.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(top_ctrl, text="全体検索:", style="Card.TLabel").pack(side=tk.LEFT, padx=(2, 2))
        self.search_entry = ttk.Entry(top_ctrl, textvariable=self.search_var, width=20)
        self.search_entry.pack(side=tk.LEFT, padx=(0, 3))
        self.search_entry.bind("<KeyRelease>", self._on_search_key_release)
        self.search_entry.bind("<Return>", self._run_search_now)
        self.search_entry.bind("<Escape>", lambda e: self._clear_search())

        ttk.Button(top_ctrl, text="検索", width=5, command=self._run_search_now).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Button(top_ctrl, text="クリア", width=6, command=self._clear_search).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(top_ctrl, text="表示対象:", style="Card.TLabel").pack(side=tk.LEFT, padx=(2, 2))
        filter_cb = ttk.Combobox(
            top_ctrl,
            textvariable=self.filter_mode_var,
            values=[
                "すべて表示",
                "実績ありのみ",
                "作業実績ありのみ",
                "購入・外注実績ありのみ",
                "作業＋購入・外注の両方あり",
            ],
            state="readonly",
            width=21,
        )
        filter_cb.pack(side=tk.LEFT, padx=(0, 10))
        filter_cb.bind("<<ComboboxSelected>>", lambda e: self.apply_filter_and_refresh(save_cfg=True))

        ttk.Label(top_ctrl, text="一括表示:", style="Card.TLabel").pack(side=tk.LEFT, padx=(2, 4))
        unit_seg = tk.Frame(top_ctrl, bg=COLORS["border"], padx=1, pady=1)
        unit_seg.pack(side=tk.LEFT, padx=(0, 6))
        self.unit_buttons = {}
        for u_val, u_text in (("金額(円)", "金額(円)"), ("分(そのまま)", "時間(分)"), ("時間(分÷60)", "時間(h)"), ("工数(8h=1)", "工数(8h)")):
            b = tk.Button(
                unit_seg,
                text=u_text,
                relief=tk.FLAT,
                bd=0,
                padx=9,
                pady=3,
                font=("Yu Gothic UI", 9, "bold"),
                cursor="hand2",
                command=lambda u=u_val: self.set_time_unit(u),
            )
            b.pack(side=tk.LEFT)
            self.unit_buttons[u_val] = b
        self._update_unit_button_styles()

        ttk.Button(
            top_ctrl,
            text="レート・既定設定",
            command=self.open_rate_settings_dialog,
        ).pack(side=tk.LEFT, padx=(0, 6))

        self.btn_excel_open = ttk.Button(
            top_ctrl,
            text="Excel新規ブックで開く",
            style="Success.TButton",
            command=self.open_excel_new_workbook_action,
        )
        self.btn_excel_open.pack(side=tk.RIGHT, padx=(4, 0))

        ttk.Button(
            top_ctrl,
            text="列を選んでコピー",
            command=self.open_column_copy_dialog,
        ).pack(side=tk.RIGHT, padx=4)

        bottom_ctrl = ttk.Frame(self.ctrl_frame, style="Flat.TFrame")
        bottom_ctrl.pack(fill=tk.X)

        ttk.Label(
            bottom_ctrl,
            textvariable=self.code_summary_var,
            style="Badge.TLabel",
            padding=(6, 2),
        ).pack(side=tk.LEFT)

        self.btn_error_report = ttk.Button(
            bottom_ctrl,
            text="エラー確認 (0件)",
            command=self.open_error_report_dialog,
        )
        self.btn_error_report.pack(side=tk.RIGHT, padx=(6, 0))

        ttk.Button(
            bottom_ctrl,
            text="全列フィルター解除",
            command=self.clear_all_column_filters,
        ).pack(side=tk.RIGHT, padx=(6, 0))

        ttk.Button(
            bottom_ctrl,
            text="全列幅自動調整",
            command=self.autofit_all_columns,
        ).pack(side=tk.RIGHT, padx=(6, 0))

        self.lbl_filter_status = ttk.Label(
            bottom_ctrl,
            textvariable=self.filter_status_var,
            style="Badge.TLabel",
            padding=(8, 2),
        )
        self.lbl_filter_status.pack(side=tk.RIGHT)

        tk.Checkbutton(
            bottom_ctrl,
            text="行リンク連動 (縦スクロール・並び順・行強調を左右で揃える)",
            variable=self.link_rows_var,
            command=self._on_link_rows_toggled,
            bg=COLORS["bg_card"],
            fg=COLORS["accent"],
            activebackground=COLORS["bg_card"],
            selectcolor="#FFFFFF",
            font=("Yu Gothic UI", 9, "bold"),
        ).pack(side=tk.RIGHT, padx=(0, 8))

        # 集計・一覧更新の進捗バーパネル（処理中のみ操作バー下部に大きく見やすく表示）
        self.progress_panel = tk.Frame(
            self.ctrl_frame,
            bg=COLORS["accent_light"],
            highlightbackground=COLORS["accent"],
            highlightthickness=1,
            padx=10,
            pady=4,
        )
        tk.Label(
            self.progress_panel,
            textvariable=self.progress_msg_var,
            bg=COLORS["accent_light"],
            fg=COLORS["accent"],
            font=("Yu Gothic UI", 9, "bold"),
            anchor=tk.W,
            width=46,
        ).pack(side=tk.LEFT, padx=(0, 8))
        self.progress_bar = ttk.Progressbar(
            self.progress_panel,
            mode="determinate",
            maximum=100,
            value=0,
        )
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(
            self.progress_panel,
            text="中止",
            width=6,
            command=self.cancel_aggregation,
        ).pack(side=tk.RIGHT, padx=(6, 0))
        tk.Label(
            self.progress_panel,
            textvariable=self.progress_pct_var,
            bg=COLORS["accent_light"],
            fg=COLORS["accent"],
            font=("Yu Gothic UI", 9, "bold"),
            width=5,
            anchor=tk.E,
        ).pack(side=tk.RIGHT)

        # 4. 最下部ステータスバー（メイン分割領域より先に side=BOTTOM で確保し、途切れを防止）
        status_bar = tk.Frame(self.root, bg="#E2E8F0", padx=10, pady=3)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        tk.Label(
            status_bar,
            textvariable=self.status_var,
            bg="#E2E8F0",
            fg=COLORS["fg_main"],
            font=("Yu Gothic UI", 9),
            anchor=tk.W,
        ).pack(side=tk.LEFT)
        tk.Label(
            status_bar,
            textvariable=self.config_info_var,
            bg="#E2E8F0",
            fg=COLORS["fg_sub"],
            font=("Yu Gothic UI", 9),
            anchor=tk.E,
        ).pack(side=tk.RIGHT)

        # 5. メインスプリットペイン（上：左右の制作伝票一覧テーブル、下：選択した管理Noの内訳明細タブ）
        self.v_paned = tk.PanedWindow(self.root, orient=tk.VERTICAL, bg=COLORS["border"], sashwidth=6, relief=tk.FLAT)
        self.v_paned.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=3)

        # 上部：左右2分割（左：その他見出し / 右：算出データ・各小計）
        list_outer = ttk.Frame(self.v_paned, style="Flat.TFrame")
        self.v_paned.add(list_outer, minsize=250, stretch="always")

        # 総合計サマリーバー（h_paned より先に side=BOTTOM で確保し、ウィンドウ幅に応じた折り返しに対応）
        summary_bar = tk.Frame(list_outer, bg=COLORS["total_bg"], padx=8, pady=4)
        summary_bar.pack(side=tk.BOTTOM, fill=tk.X, pady=(3, 0))
        self.summary_lbl = tk.Label(
            summary_bar,
            textvariable=self.summary_var,
            bg=COLORS["total_bg"],
            fg=COLORS["fg_main"],
            font=("Yu Gothic UI", 9, "bold"),
            justify=tk.LEFT,
            anchor=tk.W,
        )
        self.summary_lbl.pack(fill=tk.X)
        summary_bar.bind("<Configure>", lambda e: self.summary_lbl.configure(wraplength=max(400, e.width - 16)))

        # 算出データ内訳表示トグルバー（h_paned の外側に配置し、左右ペインの行高・開始Y座標を完全一致させる）
        detail_toggle_bar = tk.Frame(list_outer, bg=COLORS["bg_header"], padx=8, pady=3)
        detail_toggle_bar.pack(side=tk.TOP, fill=tk.X, pady=(0, 3))
        tk.Label(
            detail_toggle_bar,
            text="算出データの内訳表示:",
            bg=COLORS["bg_header"],
            fg=COLORS["fg_sub"],
            font=("Yu Gothic UI", 9, "bold"),
        ).pack(side=tk.LEFT, padx=(0, 6))
        self.btn_toggle_sub1 = ttk.Button(
            detail_toggle_bar,
            text="▶ 小計1内訳を表示",
            command=self.toggle_subtotal1_details,
        )
        self.btn_toggle_sub1.pack(side=tk.LEFT, padx=(0, 4))
        self.btn_toggle_sub2 = ttk.Button(
            detail_toggle_bar,
            text="▶ 小計2内訳を表示",
            command=self.toggle_subtotal2_details,
        )
        self.btn_toggle_sub2.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(
            detail_toggle_bar,
            text="全内訳を表示",
            command=lambda: self.set_subtotal_details_visibility(True, True),
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            detail_toggle_bar,
            text="小計のみ表示",
            command=lambda: self.set_subtotal_details_visibility(False, False),
        ).pack(side=tk.LEFT, padx=2)
        tk.Label(
            detail_toggle_bar,
            textvariable=self.detail_state_var,
            bg=COLORS["bg_header"],
            fg=COLORS["accent"],
            font=("Yu Gothic UI", 9, "bold"),
            anchor=tk.E,
        ).pack(side=tk.RIGHT, padx=(8, 2))

        self.h_paned = tk.PanedWindow(list_outer, orient=tk.HORIZONTAL, bg=COLORS["border"], sashwidth=7, relief=tk.FLAT)
        self.h_paned.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.meta_pane = SplitGridPane(
            self.h_paned,
            "その他見出し・受注基本情報（横スクロール・コピー個別可）",
            "meta",
            self,
        )
        self.calc_pane = SplitGridPane(
            self.h_paned,
            "算出データ・各小計（日報・購入品・外注加工・制作伝票）",
            "calc",
            self,
        )
        self.h_paned.add(self.meta_pane, minsize=300, stretch="always")
        self.h_paned.add(self.calc_pane, minsize=360, stretch="always")
        self.active_pane = self.calc_pane

        self.root.bind("<Control-f>", lambda e: (self.search_entry.focus_set(), self.search_entry.select_range(0, tk.END)))
        self.root.bind("<F5>", lambda e: None if self._is_loading else self.run_aggregation())

        self._configure_main_tree_columns()
        self._update_unit_button_styles()
        self._sash_after_id = self.root.after(100, self._init_horizontal_sash)

        # 下部：選択した管理Noの内訳明細ビュー
        self.detail_frame = ttk.LabelFrame(self.v_paned, text="選択した管理Noの内訳明細", padding=(6, 3))
        self.v_paned.add(self.detail_frame, minsize=145, stretch="never")

        detail_action_bar = tk.Frame(self.detail_frame, bg=COLORS["bg_header"], padx=6, pady=2)
        detail_action_bar.pack(side=tk.TOP, fill=tk.X, pady=(0, 3))
        self.detail_bar_info_var = tk.StringVar(value="一覧から管理Noを選択すると内訳明細が表示されます（Excel出力・印刷対応）")
        tk.Label(
            detail_action_bar,
            textvariable=self.detail_bar_info_var,
            bg=COLORS["bg_header"],
            fg=COLORS["fg_sub"],
            font=("Yu Gothic UI", 9, "bold"),
            anchor=tk.W,
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.btn_detail_excel_all = ttk.Button(
            detail_action_bar,
            text="内訳明細をExcelで開く（全タブ）",
            style="Success.TButton",
            command=lambda: self.open_detail_excel_action(only_active_tab=False),
        )
        self.btn_detail_excel_all.pack(side=tk.RIGHT, padx=(4, 0))
        self.btn_detail_excel_tab = ttk.Button(
            detail_action_bar,
            text="表示中タブをExcelで開く",
            command=lambda: self.open_detail_excel_action(only_active_tab=True),
        )
        self.btn_detail_excel_tab.pack(side=tk.RIGHT, padx=4)

        self.detail_notebook = ttk.Notebook(self.detail_frame)
        self.detail_notebook.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # タブ1: 日報明細
        tab_daily = ttk.Frame(self.detail_notebook, style="Card.TFrame")
        self.detail_notebook.add(tab_daily, text="日報データ明細 (0件)")
        daily_cols = ("更新日", "作業者", "作業コード", "作業名", "分類区分", "時間(分)", "時間(h)", "工数(h/8)", "金額(円)", "出荷先", "製品詳細", "元ファイル")
        self.tree_daily = self._create_detail_tree(
            tab_daily,
            daily_cols,
            [135, 115, 72, 110, 95, 78, 78, 82, 90, 95, 200, 140],
            right_cols={"作業コード", "時間(分)", "時間(h)", "工数(h/8)", "金額(円)"},
        )

        # タブ2: 購入品明細
        tab_kounyu = ttk.Frame(self.detail_notebook, style="Card.TFrame")
        self.detail_notebook.add(tab_kounyu, text="購入品明細 (0件)")
        kounyu_cols = ("図面番号", "品番", "品名", "メーカー", "仕様", "仕入先", "単価", "必要数", "金額(単価X必要数)", "注文日", "請求日", "備考", "元ファイル")
        self.tree_kounyu = self._create_detail_tree(
            tab_kounyu,
            kounyu_cols,
            [80, 60, 130, 90, 155, 110, 80, 65, 120, 85, 85, 130, 130],
            right_cols={"単価", "必要数", "金額(単価X必要数)"},
        )

        # タブ3: 外注加工明細
        tab_gaichu = ttk.Frame(self.detail_notebook, style="Card.TFrame")
        self.detail_notebook.add(tab_gaichu, text="外注加工明細 (0件)")
        self.tree_gaichu = self._create_detail_tree(
            tab_gaichu,
            kounyu_cols,
            [80, 60, 130, 90, 155, 110, 80, 65, 120, 85, 85, 130, 130],
            right_cols={"単価", "必要数", "金額(単価X必要数)"},
        )

        # タブ4: 受注表全項目
        tab_juchu = ttk.Frame(self.detail_notebook, style="Card.TFrame")
        self.detail_notebook.add(tab_juchu, text="受注表詳細情報")
        juchu_cols = ("項目名", "内容", "項目名2", "内容2")
        self.tree_juchu = self._create_detail_tree(
            tab_juchu,
            juchu_cols,
            [140, 300, 140, 300],
            right_cols=set(),
        )

        # タブ5: 制作伝票明細
        tab_seizo = ttk.Frame(self.detail_notebook, style="Card.TFrame")
        self.detail_notebook.add(tab_seizo, text="制作伝票明細 (0件)")
        seizo_cols = ("行", "項目名", "材質・仕様", "単価 / レート", "時間 / 重量", "金額(円)", "更新日時", "元ファイル")
        self.tree_seizo = self._create_detail_tree(
            tab_seizo,
            seizo_cols,
            [52, 165, 110, 100, 100, 120, 145, 220],
            right_cols={"行", "単価 / レート", "時間 / 重量", "金額(円)"},
        )

    def _create_detail_tree(self, parent, columns, widths, right_cols):
        """明細表示用Treeviewを作成する（Ctrl+Cおよび右クリックでのExcel貼付用コピー対応）"""
        frame = ttk.Frame(parent, style="Card.TFrame")
        frame.pack(fill=tk.BOTH, expand=True)
        tv = ttk.Treeview(frame, columns=columns, show="headings", selectmode="extended", height=4)
        vsb = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tv.yview)
        hsb = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tv.xview)
        tv.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tv.grid(row=0, column=0, sticky=tk.NSEW)
        vsb.grid(row=0, column=1, sticky=tk.NS)
        hsb.grid(row=1, column=0, sticky=tk.EW)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        for col, w in zip(columns, widths):
            anchor = tk.E if col in right_cols else tk.W
            tv.heading(col, text=col)
            tv.column(col, width=w, minwidth=40, anchor=anchor, stretch=False)

        tv.tag_configure("even", background=COLORS["row_even"], foreground=COLORS["fg_main"])
        tv.tag_configure("odd", background=COLORS["row_odd"], foreground=COLORS["fg_main"])

        def copy_detail_rows(all_rows: bool = False, include_header: bool = True):
            items = tv.get_children() if all_rows else tv.selection()
            if not items:
                items = tv.get_children()
            if not items:
                return "break"
            lines = []
            if include_header:
                lines.append("\t".join(escape_tsv_cell(c) for c in columns))
            for iid in items:
                vals = tv.item(iid, "values")
                lines.append("\t".join(escape_tsv_cell(v) for v in vals))
            self.root.clipboard_clear()
            self.root.clipboard_append("\r\n".join(lines))
            self.status_var.set(f"内訳明細をクリップボードにコピーしました（{len(items):,}行・Excelへ Ctrl+V で貼付可能）")
            return "break"

        menu = tk.Menu(tv, tearoff=0, bg="#FFFFFF", fg=COLORS["fg_main"])
        menu.add_command(label="選択行をコピー（見出し付き・Excel貼付用）[Ctrl+C]", command=lambda: copy_detail_rows(all_rows=False, include_header=True))
        menu.add_command(label="選択行をコピー（見出しなし）", command=lambda: copy_detail_rows(all_rows=False, include_header=False))
        menu.add_separator()
        menu.add_command(label="このタブの全行をコピー（見出し付き）", command=lambda: copy_detail_rows(all_rows=True, include_header=True))
        menu.add_command(label="このタブの全行を選択 [Ctrl+A]", command=lambda: tv.selection_set(tv.get_children()))

        def on_detail_right_click(event):
            row_id = tv.identify_row(event.y)
            if row_id and row_id not in tv.selection():
                tv.selection_set(row_id)
            if tv.get_children():
                menu.tk_popup(event.x_root, event.y_root)

        tv.bind("<Control-c>", lambda e: copy_detail_rows(all_rows=False, include_header=True))
        tv.bind("<Control-C>", lambda e: copy_detail_rows(all_rows=False, include_header=True))
        tv.bind("<Control-a>", lambda e: (tv.selection_set(tv.get_children()), "break")[1])
        tv.bind("<Control-A>", lambda e: (tv.selection_set(tv.get_children()), "break")[1])
        tv.bind("<Button-3>", on_detail_right_click)
        return tv

    def _bind_global_mousewheel(self):
        """カーソル下のTreeviewおよびスクロール対応ダイアログ（ヘルプ等）にマウスホイールを滑らかに効かせる"""
        def on_wheel(event):
            try:
                x, y = self.root.winfo_pointerxy()
                target = self.root.winfo_containing(x, y)
            except Exception:
                return
            if not target:
                return
            w = target
            while w is not None:
                if isinstance(w, ttk.Treeview):
                    delta = -1 if event.delta > 0 else 1
                    steps = max(1, abs(int(event.delta / 120))) * 3
                    if event.state & 0x0001:  # Shiftキー押下で横スクロール
                        w.xview_scroll(delta * steps, "units")
                    else:
                        w.yview_scroll(delta * steps, "units")
                    return "break"
                scroll_cv = getattr(w, "_wheel_scroll_canvas", None)
                if scroll_cv is not None:
                    try:
                        if scroll_cv.winfo_exists():
                            delta = -1 if event.delta > 0 else 1
                            steps = max(1, abs(int(event.delta / 120))) * 3
                            scroll_cv.yview_scroll(delta * steps, "units")
                            return "break"
                    except Exception:
                        pass
                w = w.master

        self.root.bind_all("<MouseWheel>", on_wheel, add="+")
        self.root.bind_all("<Shift-MouseWheel>", on_wheel, add="+")

    def toggle_file_panel(self):
        """ファイル設定エリアの折りたたみ・展開を切り替える"""
        if self.file_panel_expanded:
            self.file_group.pack_forget()
            self.btn_toggle_files.configure(text="▼ ファイル設定を開く")
            self.file_panel_expanded = False
        else:
            self.file_group.pack(side=tk.TOP, fill=tk.X, padx=10, pady=2, before=self.ctrl_frame)
            self.btn_toggle_files.configure(text="▲ ファイル設定を畳む")
            self.file_panel_expanded = True
        self.save_config()

    def _init_horizontal_sash(self):
        """初回表示時に左右分割ペインおよび上下分割ペインの比率を見やすく設定する"""
        self._sash_after_id = None
        if self._closing:
            return
        try:
            total_w = self.h_paned.winfo_width()
            if total_w > 400:
                self.h_paned.sash_place(0, int(total_w * self.h_sash_ratio), 0)
            total_h = self.v_paned.winfo_height()
            if total_h > 350:
                self.v_paned.sash_place(0, 0, int(total_h * self.v_sash_ratio))
        except Exception:
            pass

    def _get_active_pane(self) -> SplitGridPane:
        return self.active_pane if self.active_pane is not None else self.calc_pane

    def _update_unit_button_styles(self):
        """一括表示切替ボタンのスタイルと算出ペイン見出しを更新する（常時ライトモード）"""
        cur = self.time_unit_var.get()
        if self.unit_buttons:
            for u_val, btn in self.unit_buttons.items():
                if u_val == cur:
                    btn.configure(bg=COLORS["accent"], fg="#FFFFFF", activebackground=COLORS["accent_hover"], activeforeground="#FFFFFF")
                else:
                    btn.configure(bg="#FFFFFF", fg=COLORS["fg_main"], activebackground="#E2E8F0", activeforeground=COLORS["fg_main"])
        if hasattr(self, "calc_pane") and self.calc_pane is not None:
            pane_titles = {
                "金額(円)": "算出データ・各小計【金額表示 (円)】（作業費・購入品費・外注加工費）",
                "分(そのまま)": "算出データ・各小計【時間表示 (分)】（作業時間・購入品費・外注加工費）",
                "時間(分÷60)": "算出データ・各小計【時間表示 (h)】（作業時間・購入品費・外注加工費）",
                "工数(8h=1)": "算出データ・各小計【工数表示 (1工数=8h)】（作業工数・購入品費・外注加工費）",
            }
            self.calc_pane.configure(text=pane_titles.get(cur, "算出データ・各小計"))

    def set_time_unit(self, unit: str):
        """一覧全体の表示単位（金額／時間／工数）をボタンから一括切り替えする"""
        if unit not in VALID_TIME_UNITS:
            return
        self.time_unit_var.set(unit)
        self._update_unit_button_styles()
        self.apply_filter_and_refresh()
        if self._current_detail_kno:
            self._show_detail_for_kno(self._current_detail_kno)
        label_map = {"金額(円)": "金額(円)", "分(そのまま)": "時間(分)", "時間(分÷60)": "時間(h)", "工数(8h=1)": "工数(1工数=8時間)"}
        self.status_var.set(f"一覧の表示を「{label_map.get(unit, unit)}」に一括切り替えしました。")

    def _configure_main_tree_columns(self):
        """その他見出しペイン（左）と算出データ・小計ペイン（右）の列構成を更新する"""
        meta_cols = [c for c in self.column_order if c in self.visible_set and c not in CALC_COLUMNS]
        if not meta_cols:
            meta_cols = ["管理No"]
        calc_only = [c for c in self.column_order if c in self.visible_set and c in CALC_COLUMNS]
        calc_cols = ["管理No"] + calc_only

        self.meta_pane.configure_columns(meta_cols)
        self.calc_pane.configure_columns(calc_cols)
        self._sync_subtotal_detail_buttons()

    def _sync_subtotal_detail_buttons(self):
        """小計1・小計2の内訳表示トグルボタンの表示テキスト・スタイル・状態バッジを現在の表示列と同期する"""
        self.show_sub1_details = any(c in self.visible_set for c in SUBTOTAL1_DETAIL_COLUMNS)
        self.show_sub2_details = any(c in self.visible_set for c in SUBTOTAL2_DETAIL_COLUMNS)
        b1 = getattr(self, "btn_toggle_sub1", None)
        b2 = getattr(self, "btn_toggle_sub2", None)
        if b1 is not None:
            b1.configure(
                text="▼ 小計1内訳を隠す" if self.show_sub1_details else "▶ 小計1内訳を表示",
                style="Primary.TButton" if self.show_sub1_details else "TButton",
            )
        if b2 is not None:
            b2.configure(
                text="▼ 小計2内訳を隠す" if self.show_sub2_details else "▶ 小計2内訳を表示",
                style="Primary.TButton" if self.show_sub2_details else "TButton",
            )
        calc_cnt = len(self.calc_pane.visible_columns) - 1 if hasattr(self, "calc_pane") else 0
        if self.show_sub1_details and self.show_sub2_details:
            state_txt = f"表示中: 全内訳展開 ({calc_cnt}列)"
        elif self.show_sub1_details:
            state_txt = f"表示中: 小計1内訳展開 ({calc_cnt}列)"
        elif self.show_sub2_details:
            state_txt = f"表示中: 小計2内訳展開 ({calc_cnt}列)"
        else:
            state_txt = f"表示中: 小計・合計のみ ({calc_cnt}列)"
        self.detail_state_var.set(state_txt)

    def set_subtotal_details_visibility(self, show_sub1: bool, show_sub2: bool):
        """小計1・小計2の内訳項目の表示・非表示を一括切り替えして画面と設定を更新する"""
        prev_y = 0.0
        try:
            prev_y = self.meta_pane.tree.yview()[0]
        except Exception:
            pass
        self.show_sub1_details = bool(show_sub1)
        self.show_sub2_details = bool(show_sub2)
        if self.show_sub1_details:
            self.visible_set.update(SUBTOTAL1_DETAIL_COLUMNS)
        else:
            self.visible_set.difference_update(SUBTOTAL1_DETAIL_COLUMNS)
        if self.show_sub2_details:
            self.visible_set.update(SUBTOTAL2_DETAIL_COLUMNS)
        else:
            self.visible_set.difference_update(SUBTOTAL2_DETAIL_COLUMNS)
        self.visible_columns = [c for c in self.column_order if c in self.visible_set]
        self._configure_main_tree_columns()
        self._refresh_both_panes_sorted()
        self._cancel_after("_restore_y_after_id")
        if prev_y > 0.0 and not self._closing:
            def restore_y():
                self._restore_y_after_id = None
                if self._closing:
                    return
                try:
                    self._syncing_yscroll = True
                    self.meta_pane.tree.yview_moveto(prev_y)
                    self.calc_pane.tree.yview_moveto(prev_y)
                except Exception:
                    pass
                finally:
                    self._syncing_yscroll = False
            self._restore_y_after_id = self.root.after(20, restore_y)
        self.save_config()
        s1_txt = "表示" if self.show_sub1_details else "非表示"
        s2_txt = "表示" if self.show_sub2_details else "非表示"
        self.status_var.set(f"内訳列の表示を切り替えました（小計1内訳: {s1_txt} / 小計2内訳: {s2_txt}）")

    def toggle_subtotal1_details(self):
        """小計1に対する内訳項目の表示・非表示を切り替える"""
        self.set_subtotal_details_visibility(not self.show_sub1_details, self.show_sub2_details)

    def toggle_subtotal2_details(self):
        """小計2に対する内訳項目の表示・非表示を切り替える"""
        self.set_subtotal_details_visibility(self.show_sub1_details, not self.show_sub2_details)

    def _capture_current_column_widths(self):
        """左右両ペインの現在列幅をメモリに記録する"""
        self.column_widths.update(self.meta_pane.capture_column_widths())
        self.column_widths.update(self.calc_pane.capture_column_widths())

    def _on_pane_y_scroll(self, src_pane_id: str, first_frac: str):
        """行リンクON時、片方の縦スクロールにもう片方を連動させる"""
        if self._syncing_yscroll or not self.link_rows_var.get():
            return
        self._syncing_yscroll = True
        try:
            other = self.calc_pane if src_pane_id == "meta" else self.meta_pane
            other.tree.yview_moveto(first_frac)
            other._schedule_overlay_update()
        except Exception:
            pass
        finally:
            self._syncing_yscroll = False

    def _on_row_focused(self, src_pane_id: str, kno: str):
        """行・セルクリック時：下部明細を更新し、行リンクONなら左右の同じ管理No行を強調する"""
        self._show_detail_for_kno(kno)
        if self.link_rows_var.get():
            self.meta_pane.set_linked_row(kno)
            self.calc_pane.set_linked_row(kno)
        else:
            pane = self.meta_pane if src_pane_id == "meta" else self.calc_pane
            other = self.calc_pane if src_pane_id == "meta" else self.meta_pane
            pane.set_linked_row(kno)
            other.set_linked_row(None)

    def _on_link_rows_toggled(self):
        """行リンク連動のON/OFF切り替え"""
        if self.link_rows_var.get():
            pane = self._get_active_pane()
            self.sort_col = pane.sort_col if pane.sort_col in ALL_AVAILABLE_COLUMNS else "管理No"
            self.sort_desc = pane.sort_desc
            self.apply_filter_and_refresh(save_cfg=True)
            if self._current_detail_kno:
                self.meta_pane.set_linked_row(self._current_detail_kno)
                self.calc_pane.set_linked_row(self._current_detail_kno)
            self.status_var.set("行リンク連動をONにしました（縦スクロール・並び順・選択行が左右で連動します）。")
        else:
            self.meta_pane.sort_col, self.meta_pane.sort_desc = self.sort_col, self.sort_desc
            self.calc_pane.sort_col, self.calc_pane.sort_desc = self.sort_col, self.sort_desc
            self.save_config()
            self.status_var.set("行リンク連動をOFFにしました（左右それぞれ個別に縦スクロール・並替できます）。")

    def focus_kno_in_other_pane(self, src_pane_id: str, kno: str):
        """片方の一覧で選んだ管理Noを、もう一方の一覧でも表示・選択する"""
        target = self.calc_pane if src_pane_id == "meta" else self.meta_pane
        target.scroll_to_kno(kno)
        target.set_linked_row(kno)
        self.active_pane = target

    def _copy_text(self, text: str):
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.status_var.set(f"クリップボードにコピーしました: {text}")

    def copy_columns_to_clipboard(self, columns: list, include_header: bool = False, raw_numbers: bool = True, selected_only: bool = False):
        """指定列（左右ペインまたぎ対応）をExcel貼付用TSVでコピーする"""
        if not self.displayed_rows:
            messagebox.showinfo("コピー対象なし", "コピーするデータがありません。先に集計を実行してください。", parent=self.root)
            return
        if not columns:
            return

        pane = self._get_active_pane()
        source_rows = pane.rows if pane.rows else self.displayed_rows
        if selected_only:
            sel_row_indices = sorted({r for r, _ in pane.selected_cells if 0 <= r < len(source_rows)})
            if sel_row_indices:
                target_rows = [source_rows[i] for i in sel_row_indices]
            else:
                sel_set = set(pane.tree.selection())
                target_rows = [r for r in source_rows if r["管理No"] in sel_set]
            if not target_rows:
                messagebox.showinfo("選択なし", "コピーする行またはセルを選択してください。", parent=self.root)
                return
        else:
            target_rows = source_rows

        lines = []
        unit_now = self.time_unit_var.get()
        rates_tuple = self._get_rates_tuple()
        if include_header:
            lines.append("\t".join(escape_tsv_cell(get_display_column_label(c, unit_now)) for c in columns))

        for r in target_rows:
            row_cells = []
            for col in columns:
                if raw_numbers:
                    if col in RATE_DEPENDENT_COLUMNS:
                        val = self._get_work_numeric_val(r, col, rates_tuple=rates_tuple, unit=unit_now)
                        cell_str = str(int(round(val))) if abs(val - round(val)) < 1e-6 else f"{val:.2f}"
                    elif col in COST_YEN_COLUMNS:
                        cell_str = str(int(round(r.get(col, 0.0))))
                    else:
                        cell_str = str(r.get(col, ""))
                else:
                    cell_str = self._get_cell_display_str(r, col, unit=unit_now, rates_tuple=rates_tuple)
                row_cells.append(escape_tsv_cell(cell_str))
            lines.append("\t".join(row_cells))

        self.root.clipboard_clear()
        self.root.clipboard_append("\r\n".join(lines))
        col_label = get_display_column_label(columns[0], unit_now) if len(columns) == 1 else f"{len(columns)}列"
        self.status_var.set(f"クリップボードにコピーしました（{col_label} × {len(target_rows):,}行・Excelへ Ctrl+V で貼付可能）")

    def open_column_copy_dialog(self):
        """複数列・単一列のコピーダイアログを開く（左右両ペインの列を選択可能）"""
        if not self.displayed_rows:
            messagebox.showinfo("コピー対象なし", "コピーするデータがありません。先に集計を実行してください。", parent=self.root)
            return
        pane = self._get_active_pane()
        sel_row_count = len({r for r, _ in pane.selected_cells}) or len(pane.tree.selection())
        ColumnCopyDialog(
            self.root,
            self.visible_columns,
            len(self.displayed_rows),
            sel_row_count,
            self.copy_columns_to_clipboard,
            display_unit=self.time_unit_var.get(),
        )

    def get_protected_source_paths(self) -> set:
        """参照ファイル（絶対変更禁止）の正規化パス集合を返す（フォルダ展開I/Oはせず、フォルダは get_protected_source_dirs で保護）"""
        protected = set()
        for k, paths in self.file_lists.items():
            for p in paths:
                if not p:
                    continue
                if k == "seizo" and os.path.isdir(p):
                    continue
                protected.add(os.path.normcase(os.path.abspath(os.path.normpath(p))))
        return protected

    def get_protected_source_dirs(self) -> set:
        """参照フォルダ（制作伝票フォルダおよびサブフォルダ・書き込み禁止）の正規化パス集合を返す"""
        return expand_seizo_protected_dirs(self.file_lists.get("seizo", []))

    def is_protected_source_file(self, filepath: str) -> bool:
        """指定パスが参照元ファイルまたは参照元フォルダ内に該当するか判定する"""
        return is_same_or_protected_path(
            filepath,
            self.get_protected_source_paths(),
            self.get_protected_source_dirs(),
        )

    def _capture_current_sash_ratios(self):
        """現在の左右分割および上下分割ペインの位置比率を記録する"""
        try:
            if hasattr(self, "h_paned") and self.h_paned.winfo_exists():
                tw = self.h_paned.winfo_width()
                if tw > 400:
                    sx, _ = self.h_paned.sash_coord(0)
                    if 80 <= sx <= tw - 80:
                        self.h_sash_ratio = round(min(max(sx / float(tw), 0.20), 0.80), 4)
            if hasattr(self, "v_paned") and self.v_paned.winfo_exists():
                th = self.v_paned.winfo_height()
                if th > 350:
                    _, sy = self.v_paned.sash_coord(0)
                    if 100 <= sy <= th - 100:
                        self.v_sash_ratio = round(min(max(sy / float(th), 0.35), 0.85), 4)
        except Exception:
            pass

    def load_config(self, path: str) -> dict:
        """設定ファイル（JSON）を安全に読み込む（破損時はバックアップも参照）"""
        for candidate in (path, f"{path}.bak" if path else ""):
            if candidate and os.path.exists(candidate):
                try:
                    with open(candidate, "r", encoding="utf-8") as f:
                        loaded = json.load(f)
                        if isinstance(loaded, dict):
                            return loaded
                except Exception:
                    continue
        return {}

    def save_config(self, target_path: str = None, show_message: bool = False):
        """
        現在のファイルパス・作業コード・単価レート・列順番・列幅・分割比率・表示設定を
        一時ファイル経由のアトミック書き込みで安全かつ高速にJSONへ保存する。
        参照元ファイルおよび参照元フォルダ内への書き込みは厳重にブロックする。
        """
        self._capture_current_column_widths()
        self._capture_current_sash_ratios()
        save_to = target_path or self.config_path or self.default_config_path

        if self.is_protected_source_file(save_to):
            if show_message:
                messagebox.showerror(
                    "書き込みブロック",
                    f"参照元ファイル・参照元フォルダを保護するため、以下のパスへの保存をブロックしました:\n{save_to}",
                    parent=self.root,
                )
            return

        data = {
            "schema_version": 2,
            "files": self.file_lists,
            "code_ranges": self.code_ranges,
            "work_rates": self.work_rates,
            "default_time_unit": self.default_time_unit,
            "link_rows": self.link_rows_var.get(),
            "column_order": self.column_order,
            "visible_columns": [c for c in self.column_order if c in self.visible_set],
            "show_sub1_details": self.show_sub1_details,
            "show_sub2_details": self.show_sub2_details,
            "column_widths": self.column_widths,
            "filter_mode": self.filter_mode_var.get(),
            "time_unit": self.time_unit_var.get(),
            "sort_col": self.sort_col,
            "sort_desc": self.sort_desc,
            "file_panel_expanded": self.file_panel_expanded,
            "h_sash_ratio": self.h_sash_ratio,
            "v_sash_ratio": self.v_sash_ratio,
        }
        try:
            json_str = json.dumps(data, ensure_ascii=False, indent=2)
            # 変更がない場合の不要なディスクI/Oをスキップして高速化
            if (
                not show_message
                and save_to == self.config_path
                and json_str == self._last_saved_config_json
                and os.path.exists(save_to)
            ):
                return

            parent_dir = os.path.dirname(os.path.abspath(save_to))
            if parent_dir and not os.path.exists(parent_dir):
                os.makedirs(parent_dir, exist_ok=True)

            tmp_save = f"{save_to}.tmp.{os.getpid()}"
            try:
                with open(tmp_save, "w", encoding="utf-8") as f:
                    f.write(json_str)
                    f.flush()
                    try:
                        os.fsync(f.fileno())
                    except Exception:
                        pass
                try:
                    os.replace(tmp_save, save_to)
                except OSError:
                    time.sleep(0.06)
                    os.replace(tmp_save, save_to)
            except Exception:
                if os.path.exists(tmp_save):
                    try:
                        os.remove(tmp_save)
                    except Exception:
                        pass
                raise

            self._last_saved_config_json = json_str
            self.config_path = save_to
            self.config_info_var.set(f"設定ファイル: {self.config_path}")
            if show_message:
                messagebox.showinfo("設定保存", f"現在の設定を保存しました:\n{save_to}", parent=self.root)
        except Exception as e:
            if show_message:
                messagebox.showerror("保存エラー", f"設定ファイルの保存に失敗しました:\n{e}", parent=self.root)

    def reset_config_to_defaults(self, parent_win=None):
        """表示設定・作業コード・レート・列順・列幅を初期状態に戻して保存する"""
        if not messagebox.askyesno(
            "設定の初期化確認",
            "作業コード範囲・単価レート・既定表示単位・列の順番・列幅を初期設定に戻しますか？\n（選択中の参照ファイルパスは保持されます）",
            parent=parent_win or self.root,
        ):
            return
        self.code_ranges = dict(DEFAULT_CODE_RANGES)
        self.work_rates = dict(DEFAULT_WORK_RATES)
        self.default_time_unit = DEFAULT_TIME_UNIT
        self.time_unit_var.set(DEFAULT_TIME_UNIT)
        self.link_rows_var.set(True)
        self.column_order = list(DEFAULT_COLUMN_ORDER) + [c for c in ALL_AVAILABLE_COLUMNS if c not in DEFAULT_COLUMN_ORDER]
        self.show_sub1_details = False
        self.show_sub2_details = False
        self.visible_set = set(DEFAULT_VISIBLE_COLUMNS)
        self.visible_columns = [c for c in self.column_order if c in self.visible_set]
        self.column_widths = dict(DEFAULT_COL_WIDTHS)
        self.filter_mode_var.set("実績ありのみ")
        self.sort_col = "管理No"
        self.sort_desc = True
        self._update_code_summary_label()
        self._configure_main_tree_columns()
        self._update_unit_button_styles()
        self.save_config(target_path=self.default_config_path, show_message=False)
        if self.raw_daily or self.raw_kounyu or self.raw_gaichu or self.raw_juchu or self.raw_seizo:
            self._recalculate_from_loaded_data()
        else:
            self.apply_filter_and_refresh()
        self.status_var.set("設定を初期状態に戻して保存しました。")

    def open_config_manager_dialog(self):
        """設定の保存・保存先変更・読み込み・初期化ダイアログ（デフォルトはアプリと同じ場所）"""
        dlg = tk.Toplevel(self.root)
        dlg.title("設定ファイルの保存・読み込み")
        dlg.geometry("700x340")
        dlg.minsize(640, 310)
        dlg.configure(bg=COLORS["bg_card"])
        dlg.transient(self.root)
        dlg.grab_set()

        frame = ttk.Frame(dlg, style="Card.TFrame", padding=16)
        frame.pack(fill=tk.BOTH, expand=True)

        btn_row2 = ttk.Frame(frame, style="Flat.TFrame")
        btn_row2.pack(side=tk.BOTTOM, fill=tk.X, pady=(12, 0))
        ttk.Button(btn_row2, text="設定を初期状態に戻す", command=lambda: self.reset_config_to_defaults(dlg)).pack(side=tk.LEFT)
        ttk.Button(btn_row2, text="閉じる", command=dlg.destroy).pack(side=tk.RIGHT)

        ttk.Label(frame, text="設定の保存・読み込み管理", style="Title.TLabel").pack(anchor=tk.W)
        ttk.Label(
            frame,
            text=(
                "参照ファイル・制作伝票フォルダ・作業コード設定・単価レート・既定表示単位・行リンク連動・\n"
                f"列の順番・列幅・分割比率を自動保存します。デフォルトの保存先はアプリと同じ場所（{CONFIG_FILENAME}）です。"
            ),
            style="Mtime.TLabel",
            justify=tk.LEFT,
        ).pack(anchor=tk.W, pady=(4, 12))

        path_f = ttk.Frame(frame, style="Flat.TFrame")
        path_f.pack(fill=tk.X, pady=4)
        ttk.Label(path_f, text="現在の保存先:", style="Card.TLabel", width=13).pack(side=tk.LEFT)
        cur_var = tk.StringVar(value=self.config_path)
        ttk.Entry(path_f, textvariable=cur_var, state="readonly").pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)

        btn_row1 = ttk.Frame(frame, style="Flat.TFrame")
        btn_row1.pack(fill=tk.X, pady=(14, 6))

        def save_default():
            self.save_config(target_path=self.default_config_path, show_message=True)
            cur_var.set(self.config_path)

        def save_as():
            selected = filedialog.asksaveasfilename(
                parent=dlg,
                title="設定ファイルの保存先を選択",
                initialdir=get_base_dir(),
                initialfile=CONFIG_FILENAME,
                defaultextension=".json",
                filetypes=[("JSON設定ファイル", "*.json")],
            )
            if selected:
                self.save_config(target_path=os.path.normpath(selected), show_message=True)
                cur_var.set(self.config_path)

        def load_from():
            selected = filedialog.askopenfilename(
                parent=dlg,
                title="読み込む設定ファイルを選択",
                initialdir=get_base_dir(),
                filetypes=[("JSON設定ファイル", "*.json"), ("すべてのファイル", "*.*")],
            )
            if selected:
                self._apply_loaded_config_file(os.path.normpath(selected))
                cur_var.set(self.config_path)
                dlg.destroy()

        ttk.Button(btn_row1, text="アプリと同じ場所に保存（デフォルト）", style="Primary.TButton", command=save_default).pack(side=tk.LEFT)
        ttk.Button(btn_row1, text="別の場所に保存...", command=save_as).pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_row1, text="設定ファイルを読み込む...", command=load_from).pack(side=tk.LEFT)
        dlg.bind("<Escape>", lambda e: dlg.destroy())

    def _apply_loaded_config_file(self, filepath: str):
        """外部から読み込んだ設定JSONをアプリ状態へ適用し、デフォルト設定先とも同期する"""
        cfg = self.load_config(filepath)
        if not cfg:
            messagebox.showwarning("読込失敗", "有効な設定ファイルを読み込めませんでした。", parent=self.root)
            return
        self.config_path = filepath
        self.config_info_var.set(f"設定ファイル: {self.config_path}")

        files = cfg.get("files", {})
        for k in ("daily", "kounyu", "gaichu", "juchu", "seizo"):
            self.file_lists[k] = self._normalize_file_list(files.get(k, []))

        ranges = cfg.get("code_ranges", {})
        for cat, def_val in DEFAULT_CODE_RANGES.items():
            self.code_ranges[cat] = ranges.get(cat, def_val)

        rates = cfg.get("work_rates", {})
        for cat, def_rate in DEFAULT_WORK_RATES.items():
            val = to_float(rates.get(cat, def_rate), def_rate)
            self.work_rates[cat] = val if val >= 0 else def_rate

        saved_order = cfg.get("column_order", list(DEFAULT_COLUMN_ORDER))
        saved_order_set = set(saved_order)
        self.column_order = [c for c in saved_order if c in ALL_AVAILABLE_COLUMNS]
        for c in ALL_AVAILABLE_COLUMNS:
            if c not in self.column_order:
                self.column_order.append(c)

        has_sub_detail_flags = ("show_sub1_details" in cfg) or ("show_sub2_details" in cfg)
        self.show_sub1_details = bool(cfg.get("show_sub1_details", False))
        self.show_sub2_details = bool(cfg.get("show_sub2_details", False))

        saved_visible = cfg.get("visible_columns", list(DEFAULT_VISIBLE_COLUMNS))
        self.visible_set = {c for c in saved_visible if c in ALL_AVAILABLE_COLUMNS} or set(DEFAULT_VISIBLE_COLUMNS)
        for c in ALL_AVAILABLE_COLUMNS:
            if c not in saved_order_set and c not in DEFAULT_HIDDEN_COLUMNS:
                self.visible_set.add(c)
        if not has_sub_detail_flags:
            self.visible_set.difference_update(SUBTOTAL1_DETAIL_COLUMNS)
            self.visible_set.difference_update(SUBTOTAL2_DETAIL_COLUMNS)
        self.visible_columns = [c for c in self.column_order if c in self.visible_set]

        saved_widths = cfg.get("column_widths", {})
        self.column_widths = dict(DEFAULT_COL_WIDTHS)
        for k, v in saved_widths.items():
            if isinstance(v, (int, float)) and 35 <= int(v) <= 800:
                self.column_widths[k] = int(v)

        if "filter_mode" in cfg:
            self.filter_mode_var.set(cfg["filter_mode"])
        if "link_rows" in cfg:
            self.link_rows_var.set(bool(cfg["link_rows"]))
        if "default_time_unit" in cfg and cfg["default_time_unit"] in VALID_TIME_UNITS:
            self.default_time_unit = cfg["default_time_unit"]
        else:
            self.default_time_unit = DEFAULT_TIME_UNIT
        if "time_unit" in cfg and cfg["time_unit"] in VALID_TIME_UNITS:
            self.time_unit_var.set(cfg["time_unit"])
        else:
            self.time_unit_var.set(self.default_time_unit)
        if "sort_col" in cfg and cfg["sort_col"] in ALL_AVAILABLE_COLUMNS:
            self.sort_col = cfg["sort_col"]
        if "sort_desc" in cfg:
            self.sort_desc = bool(cfg["sort_desc"])
        if "h_sash_ratio" in cfg:
            self.h_sash_ratio = min(max(to_float(cfg["h_sash_ratio"], 0.46), 0.20), 0.80)
        if "v_sash_ratio" in cfg:
            self.v_sash_ratio = min(max(to_float(cfg["v_sash_ratio"], 0.66), 0.35), 0.85)
        self._update_unit_button_styles()

        # 次回起動時にも反映されるようデフォルトパスにも保存
        self.save_config(target_path=self.default_config_path, show_message=False)
        self.config_path = filepath
        self.config_info_var.set(f"設定ファイル: {self.config_path}")

        self._refresh_file_display_vars(skip_io=True)
        self._update_code_summary_label()
        self._configure_main_tree_columns()
        self._init_horizontal_sash()
        if any(bool(paths) for paths in self.file_lists.values()):
            self._refresh_file_mtime_async()
            self.status_var.set("設定を読み込みました。「集計・一覧更新」ボタンを押すと集計を開始します。")
        else:
            self.status_var.set("ファイルを選択して「集計・一覧更新」ボタンを押してください。")

    def _startup_async_init(self):
        """起動直後はファイル検出と更新日時確認のみ行い、ボタンを押すまで自動集計はしない"""
        self._startup_after_id = None
        if self._closing:
            return
        has_configured = any(bool(paths) for paths in self.file_lists.values())
        if not has_configured:
            self._auto_detect_default_files()

        if any(bool(paths) for paths in self.file_lists.values()):
            self._refresh_file_mtime_async()
            self.status_var.set("準備完了：「集計・一覧更新」ボタンを押すとファイルを読み込み集計します。")
        else:
            self.status_var.set("ファイルを選択して「集計・一覧更新」ボタンを押してください。")

    def _auto_detect_default_files(self):
        """未設定のファイルについて、アプリと同じフォルダの既定ファイルを自動検出する"""
        base = get_base_dir()
        defaults = {
            "daily": "daily_report.csv",
            "kounyu": "購入品.xlsm",
            "gaichu": "外注加工.xlsm",
            "juchu": "受注票頚城.xlsm",
        }
        for key, fname in defaults.items():
            if not self.file_lists[key]:
                candidate = os.path.join(base, fname)
                if os.path.exists(candidate):
                    self.file_lists[key] = [candidate]

    def _refresh_file_display_vars(self, daily_latest_in_data: str = "", skip_io: bool = False):
        """ファイル選択表示と更新日時ラベルを最新化する"""
        for key in ("daily", "kounyu", "gaichu", "juchu", "seizo"):
            paths = self.file_lists[key]
            self.path_summary_vars[key].set(summarize_file_paths_text(paths))
            if skip_io:
                if not paths:
                    self.mtime_vars[key].set("未選択")
                else:
                    self.mtime_vars[key].set(f"確認中... ({len(paths)}件)" if len(paths) > 1 else "確認中...")
            else:
                if key == "seizo":
                    self.mtime_vars[key].set(summarize_seizo_folder_mtime(paths))
                else:
                    extra = f"(データ内最新: {daily_latest_in_data})" if (key == "daily" and daily_latest_in_data) else ""
                    self.mtime_vars[key].set(summarize_file_list_mtime(paths, extra_info=extra))

    def _refresh_file_mtime_async(self):
        """ファイル更新日時だけを非同期で取得してラベル更新する"""
        snapshot = {k: list(v) for k, v in self.file_lists.items()}
        self._refresh_file_display_vars(skip_io=True)

        def worker():
            res = {}
            for k, paths in snapshot.items():
                if self._closing:
                    return
                res[k] = summarize_seizo_folder_mtime(paths) if k == "seizo" else summarize_file_list_mtime(paths)
            if not self._closing:
                self._ui_queue.put(lambda: self._apply_mtime_results(res))

        threading.Thread(target=worker, daemon=True).start()

    def _apply_mtime_results(self, mtime_dict: dict):
        if self._closing:
            return
        for k, txt in mtime_dict.items():
            if k in self.mtime_vars:
                self.mtime_vars[k].set(txt)

    def _update_code_summary_label(self):
        """作業コード範囲の現在の設定を画面上部に表示する"""
        parts = [f"{cat}: {self.code_ranges.get(cat, '')}" for cat in WORK_CATEGORIES]
        self.code_summary_var.set("作業コード設定  |  " + "   /   ".join(parts))

    def _on_search_key_release(self, event=None):
        """検索キーワードが変化したときデバウンス付きで一覧フィルタリングを実行して高速化する"""
        cur_kw = normalize_search_text(self.search_var.get())
        if cur_kw == self._last_search_kw:
            return
        if getattr(self, "_search_debounce_job", None) is not None:
            try:
                self.root.after_cancel(self._search_debounce_job)
            except Exception:
                pass
            self._search_debounce_job = None
        self._search_debounce_job = self.root.after(120, self._run_search_now)

    def _run_search_now(self, event=None):
        """全体検索を即時に反映する（Enter押下・検索ボタン押下・デバウンス完了時）"""
        if getattr(self, "_search_debounce_job", None) is not None:
            try:
                self.root.after_cancel(self._search_debounce_job)
            except Exception:
                pass
            self._search_debounce_job = None
        cur_kw = normalize_search_text(self.search_var.get())
        self._last_search_kw = cur_kw
        self.apply_filter_and_refresh()

    def _clear_search(self):
        if getattr(self, "_search_debounce_job", None) is not None:
            try:
                self.root.after_cancel(self._search_debounce_job)
            except Exception:
                pass
            self._search_debounce_job = None
        self.search_var.set("")
        self._last_search_kw = ""
        self.apply_filter_and_refresh()

    def select_files_replace(self, key: str, filetypes: list):
        """複数ファイル選択ダイアログ（既存選択を置き換え）"""
        cur = self.file_lists[key]
        init_dir = os.path.dirname(cur[-1]) if cur and os.path.exists(os.path.dirname(cur[-1])) else get_base_dir()
        selected = filedialog.askopenfilenames(
            parent=self.root,
            title="ファイルを選択してください（Ctrl/Shiftで複数選択可）",
            initialdir=init_dir,
            filetypes=filetypes,
        )
        if selected:
            self.file_lists[key] = self._normalize_file_list(list(selected))
            self._refresh_file_mtime_async()
            self.save_config()

    def select_files_append(self, key: str, filetypes: list):
        """複数ファイル選択ダイアログ（既存選択に追加）"""
        cur = self.file_lists[key]
        init_dir = os.path.dirname(cur[-1]) if cur and os.path.exists(os.path.dirname(cur[-1])) else get_base_dir()
        selected = filedialog.askopenfilenames(
            parent=self.root,
            title="追加するファイルを選択してください（Ctrl/Shiftで複数選択可）",
            initialdir=init_dir,
            filetypes=filetypes,
        )
        if selected:
            self.file_lists[key] = self._normalize_file_list(self.file_lists[key] + list(selected))
            self._refresh_file_mtime_async()
            self.save_config()

    def open_seizo_folder_dialog(self, initial_tab: int = 0):
        """制作伝票フォルダの複数選択・管理ダイアログを開く"""
        def on_apply(new_folders):
            self.file_lists["seizo"] = self._normalize_file_list(new_folders)
            self._refresh_file_mtime_async()
            self.save_config()

        SeizoFolderDialog(self.root, self.file_lists["seizo"], on_apply, initial_tab=initial_tab)

    def select_seizo_folder_replace(self):
        """制作伝票フォルダを複数選択・管理できるダイアログを開く"""
        self.open_seizo_folder_dialog(initial_tab=0)

    def select_seizo_folder_append(self):
        """制作伝票Excelが入った別フォルダを追加する（指定フォルダ直下のみ参照）"""
        cur = self.file_lists["seizo"]
        init_dir = cur[-1] if cur and os.path.exists(cur[-1]) else get_base_dir()
        folder = filedialog.askdirectory(
            parent=self.root,
            title="追加する制作伝票フォルダを選択してください（サブフォルダは含みません）",
            initialdir=init_dir,
        )
        if folder:
            self.file_lists["seizo"] = self._normalize_file_list(self.file_lists["seizo"] + [folder])
            self._refresh_file_mtime_async()
            self.save_config()

    def open_seizo_file_list_dialog(self):
        """制作伝票フォルダ管理・検出済みExcelファイル一覧ダイアログ（一覧タブ）を開く"""
        self.open_seizo_folder_dialog(initial_tab=1)

    def open_file_list_dialog(self, key: str, label_text: str, filetypes: list):
        """指定カテゴリの複数ファイル一覧編集ダイアログを開く"""
        def on_apply(new_list):
            self.file_lists[key] = self._normalize_file_list(new_list)
            self._refresh_file_mtime_async()
            self.save_config()

        FileListDialog(self.root, label_text, self.file_lists[key], filetypes, on_apply)

    def clear_files(self, key: str):
        """指定カテゴリの選択ファイルをクリアする"""
        self.file_lists[key] = []
        self._refresh_file_display_vars(skip_io=True)
        self.save_config()

    def _update_error_button_state(self):
        """直近のエラー件数に応じてエラー確認ボタンの表示と強調スタイルを更新する"""
        if not hasattr(self, "btn_error_report"):
            return
        cnt = len(self.last_errors)
        if cnt > 0:
            self.btn_error_report.configure(text=f"エラー確認 ({cnt}件)", style="Danger.TButton")
        else:
            self.btn_error_report.configure(text="エラー確認 (0件)", style="TButton")

    def open_error_report_dialog(self):
        """直近のフォーマット不一致・読み込みエラーの詳細確認ダイアログを開く"""
        if not self.last_errors:
            messagebox.showinfo(
                "エラーなし",
                "直近の集計でフォーマット不一致や読み込みエラーは発生していません。",
                parent=self.root,
            )
            return
        ErrorReportDialog(self.root, self.last_errors)

    @staticmethod
    def _make_error_entry(category: str, filepath: str, exc: Exception, default_kind: str = "読込エラー") -> dict:
        if isinstance(exc, FormatValidationError):
            return exc.to_dict()
        if isinstance(exc, FileNotFoundError):
            return {
                "category": category,
                "file": os.path.basename(filepath) if filepath else "-",
                "path": filepath or "",
                "kind": "ファイル不在",
                "reason": f"指定されたファイルが見つかりません: {filepath}",
                "expected": "有効なファイルパス",
                "detected": "ファイルが存在しません",
            }
        if isinstance(exc, PermissionError) or getattr(exc, "winerror", None) in (5, 32, 33):
            return {
                "category": category,
                "file": os.path.basename(filepath) if filepath else "-",
                "path": filepath or "",
                "kind": "他ユーザー使用中",
                "reason": "他のユーザーまたは別アプリがファイルを排他ロック（保存処理中など）しているため読み込めませんでした。数秒後に再度「集計・一覧更新」をお試しください。",
                "expected": "読み取り可能なファイル状態",
                "detected": str(exc),
            }
        return {
            "category": category,
            "file": os.path.basename(filepath) if filepath else "-",
            "path": filepath or "",
            "kind": default_kind,
            "reason": str(exc),
            "expected": "",
            "detected": "",
        }

    def cancel_aggregation(self):
        """実行中の集計・ファイル読み込みを安全に中断する"""
        if not self._is_loading:
            return
        self._active_job_id += 1
        self._stop_loading_indicator(completed=False)
        self.status_var.set("集計・一覧更新を中止しました。")

    def run_aggregation(self, silent_missing: bool = False):
        """選択された全ファイル（複数可・制作伝票は指定フォルダ直下のみ）をバックグラウンドで検証・読み込みし、管理Noごとに集計を行う"""
        self.save_config()

        has_any_path = any(bool(paths) for paths in self.file_lists.values())
        if not has_any_path:
            if not silent_missing:
                messagebox.showwarning(
                    "ファイル未選択",
                    "集計対象のファイルが選択されていません。\n各ファイルの「選択(複数可)」ボタンからファイルを選択してください。",
                    parent=self.root,
                )
            return

        self._active_job_id += 1
        job_id = self._active_job_id
        self._is_loading = True

        file_snapshot = {k: list(v) for k, v in self.file_lists.items()}
        code_ranges_snapshot = dict(self.code_ranges)
        work_rates_snapshot = dict(self.work_rates)

        self._refresh_file_display_vars(skip_io=True)
        self.status_var.set("ファイルを読み込んでいます（バックグラウンド処理中）...")
        self._start_loading_indicator("集計準備中：対象ファイルを確認しています...", 2)

        def is_cancelled():
            return self._closing or (job_id != self._active_job_id)

        last_post_t = [0.0]

        def post_progress(pct: float, msg: str, force: bool = False):
            if is_cancelled():
                return
            now = time.perf_counter()
            if force or (now - last_post_t[0] >= 0.03):
                last_post_t[0] = now
                p_clamped = min(max(float(pct), 0.0), 100.0)
                self._ui_queue.put(
                    lambda p=p_clamped, m=msg: (not is_cancelled()) and self._update_progress_ui(p, m)
                )

        def worker():
            errors = []
            daily_latest = ""
            raw_daily = []
            raw_kounyu = []
            raw_gaichu = []
            raw_juchu = []
            raw_seizo = []
            mtime_map = {}

            try:
                # 0. 全ファイルの存在・更新日時チェック (2% -> 6%) ※1パススキャンで重複statを完全排除
                seizo_folders = file_snapshot.get("seizo", [])
                seizo_entries = scan_seizo_excel_entries(seizo_folders, cancel_check=is_cancelled)
                for k, paths in file_snapshot.items():
                    if is_cancelled():
                        return
                    if k == "seizo":
                        mtime_map[k] = summarize_seizo_folder_mtime(paths, entries=seizo_entries)
                    else:
                        mtime_map[k] = summarize_file_list_mtime(paths)

                for sf in seizo_folders:
                    if sf and not os.path.exists(sf):
                        errors.append({
                            "category": "制作伝票フォルダ",
                            "file": os.path.basename(sf),
                            "path": sf,
                            "kind": "フォルダ不在",
                            "reason": f"指定された制作伝票フォルダが見つかりません: {sf}",
                            "expected": "存在するフォルダパス",
                            "detected": "フォルダ不在",
                        })

                any_exists = any(os.path.exists(p) for k, paths in file_snapshot.items() if k != "seizo" for p in paths) or bool(seizo_entries)
                if not any_exists and not errors:
                    if not is_cancelled():
                        self._ui_queue.put(
                            lambda: self._on_aggregation_no_files(job_id, mtime_map, silent_missing)
                        )
                    return

                post_progress(6, f"対象ファイルの確認完了（制作伝票 {len(seizo_entries):,}件）。読み込みを開始します...", force=True)

                # 1～4. 主要4ファイル群（日報・購入品・外注加工・受注表）のキャッシュ活用＋並列高速読み込み (6% -> 64%)
                main_cat_defs = (
                    ("daily", "日報ファイル", None, None),
                    ("kounyu", "購入品ファイル", ["テーブル"], KOUNYU_TARGET_COLS),
                    ("gaichu", "外注加工ファイル", ["テーブル"], KOUNYU_TARGET_COLS),
                    ("juchu", "受注表ファイル", ["受注表"], JUCHU_TARGET_COLS),
                )
                main_results = {}   # (cat, idx) -> payload
                main_errors = {}    # (cat, idx) -> error_dict
                main_todo = []      # list of (cat, idx, cat_label, pref_sheets, t_cols, filepath, sz, mns)
                main_total = sum(len(file_snapshot[c]) for c, _, _, _ in main_cat_defs)
                main_done = 0
                main_cache_hits = 0

                for cat, cat_label, pref_sheets, t_cols in main_cat_defs:
                    for idx, p in enumerate(file_snapshot[cat]):
                        if is_cancelled():
                            return
                        if not os.path.exists(p):
                            main_errors[(cat, idx)] = self._make_error_entry(cat_label, p, FileNotFoundError(p))
                            main_done += 1
                            continue
                        try:
                            st_m = os.stat(p)
                            sz_m = int(st_m.st_size)
                            mns_m = int(getattr(st_m, "st_mtime_ns", int(st_m.st_mtime * 1_000_000_000)))
                        except OSError:
                            sz_m, mns_m = -1, -1
                        cached_payload = self.main_file_cache.get(cat, p, sz_m, mns_m) if sz_m >= 0 else None
                        if cached_payload is not None:
                            main_results[(cat, idx)] = cached_payload
                            main_cache_hits += 1
                            main_done += 1
                        else:
                            main_todo.append((cat, idx, cat_label, pref_sheets, t_cols, p, sz_m, mns_m))

                if main_cache_hits > 0 and main_total > 0:
                    post_progress(
                        6.0 + 58.0 * (main_done / main_total),
                        f"[1～4/5] 未変更ファイル高速復元: {main_cache_hits}/{main_total}件（読込対象 {len(main_todo)}件）",
                        force=True,
                    )

                if main_todo:
                    def load_one_main(item):
                        cat, idx, cat_label, pref_sheets, t_cols, p, sz_m, mns_m = item
                        if is_cancelled():
                            return cat, idx, cat_label, p, sz_m, mns_m, None, InterruptedError()
                        try:
                            if cat == "daily":
                                recs, latest = read_daily_report_csv(p, cancel_check=is_cancelled)
                                return cat, idx, cat_label, p, sz_m, mns_m, (recs, latest), None
                            rows, m_sheet, all_sh = read_excel_sheet_rows(
                                p,
                                preferred_sheets=pref_sheets,
                                cancel_check=is_cancelled,
                                return_meta=True,
                                target_cols=t_cols,
                            )
                            validate_table_excel_rows(p, cat, rows, matched_sheet=m_sheet, all_sheets=all_sh)
                            data_rows = rows[1:]
                            if cat in ("kounyu", "gaichu"):
                                src_name = os.path.basename(p)
                                for r in data_rows:
                                    r["_src_file"] = src_name
                            return cat, idx, cat_label, p, sz_m, mns_m, data_rows, None
                        except InterruptedError as ie:
                            return cat, idx, cat_label, p, sz_m, mns_m, None, ie
                        except Exception as exc_m:
                            return cat, idx, cat_label, p, sz_m, mns_m, None, exc_m

                    post_progress(
                        6.0 + (58.0 * (main_done / max(1, main_total))),
                        f"[1～4/5] 日報・購入品・外注・受注表を並列読込中 ({len(main_todo)}ファイル)...",
                        force=True,
                    )
                    n_workers_main = min(4, len(main_todo))
                    if n_workers_main == 1:
                        results_iter = [load_one_main(main_todo[0])]
                    else:
                        pool_m = ThreadPoolExecutor(max_workers=n_workers_main)
                        try:
                            futs_m = [pool_m.submit(load_one_main, it) for it in main_todo]
                            results_iter = []
                            for fut_m in as_completed(futs_m):
                                if is_cancelled():
                                    for fm_rem in futs_m:
                                        fm_rem.cancel()
                                    return
                                res_tuple = fut_m.result()
                                cat, idx, cat_label, p, sz_m, mns_m, payload, err_m = res_tuple
                                if isinstance(err_m, InterruptedError) or is_cancelled():
                                    return
                                if err_m is not None:
                                    main_errors[(cat, idx)] = self._make_error_entry(cat_label, p, err_m)
                                else:
                                    main_results[(cat, idx)] = payload
                                    if sz_m >= 0:
                                        self.main_file_cache.put(cat, p, sz_m, mns_m, payload)
                                main_done += 1
                                pct_m = 6.0 + 58.0 * (main_done / max(1, main_total))
                                post_progress(
                                    pct_m,
                                    f"[1～4/5] {cat_label} 読込完了 ({main_done}/{main_total}): {os.path.basename(p)}",
                                    force=True,
                                )
                        finally:
                            pool_m.shutdown(wait=False, cancel_futures=True)
                    if n_workers_main == 1:
                        for cat, idx, cat_label, p, sz_m, mns_m, payload, err_m in results_iter:
                            if isinstance(err_m, InterruptedError) or is_cancelled():
                                return
                            if err_m is not None:
                                main_errors[(cat, idx)] = self._make_error_entry(cat_label, p, err_m)
                            else:
                                main_results[(cat, idx)] = payload
                                if sz_m >= 0:
                                    self.main_file_cache.put(cat, p, sz_m, mns_m, payload)
                            main_done += 1
                            pct_m = 6.0 + 58.0 * (main_done / max(1, main_total))
                            post_progress(
                                pct_m,
                                f"[1～4/5] {cat_label} 読込完了 ({main_done}/{main_total}): {os.path.basename(p)}",
                                force=True,
                            )

                # 元のカテゴリ・ファイル順で結果を結合
                for cat, _, _, _ in main_cat_defs:
                    for idx in range(len(file_snapshot[cat])):
                        key_ci = (cat, idx)
                        if key_ci in main_errors:
                            errors.append(main_errors[key_ci])
                        elif key_ci in main_results:
                            payload = main_results[key_ci]
                            if cat == "daily":
                                recs, latest = payload
                                raw_daily.extend(recs)
                                if latest and latest > daily_latest:
                                    daily_latest = latest
                            elif cat == "kounyu":
                                raw_kounyu.extend(payload)
                            elif cat == "gaichu":
                                raw_gaichu.extend(payload)
                            elif cat == "juchu":
                                raw_juchu.extend(payload)

                if daily_latest:
                    mtime_map["daily"] = summarize_file_list_mtime(
                        file_snapshot["daily"],
                        extra_info=f"(データ内最新: {daily_latest})",
                    )
                post_progress(64, "[1～4/5] 主要4ファイル群の読込完了", force=True)

                # 5. 制作伝票フォルダ（指定フォルダ直下の .xls / .xlsx / .xlsm）のキャッシュ活用＋高速並列読み込み (64% -> 90%)
                total_seizo = len(seizo_entries)
                if total_seizo > 0:
                    # (5-1) 未変更ファイルをインメモリ＆永続キャッシュから即時復元（O(1)）
                    self.seizo_cache.ensure_loaded()
                    cache_hits = 0
                    xml_todo = []  # .xlsx / .xlsm（並列部分ZIP読取対象）
                    xls_todo = []  # .xls（xlrd並列 or バッチExcel COM対象）
                    seizo_records_by_idx = {}

                    for seq_i, entry in enumerate(seizo_entries):
                        if is_cancelled():
                            return
                        fp_i, sz_i, mns_i, _msec_i = entry
                        cached_rec = self.seizo_cache.get(fp_i, sz_i, mns_i)
                        if cached_rec is not None:
                            seizo_records_by_idx[seq_i] = cached_rec
                            cache_hits += 1
                        else:
                            if fp_i.lower().endswith(".xls"):
                                xls_todo.append((seq_i, entry))
                            else:
                                xml_todo.append((seq_i, entry))

                    done_lock = threading.Lock()
                    done_cnt = [cache_hits]

                    def report_seizo_progress(cur_name: str, force_update: bool = False):
                        with done_lock:
                            cur_d = done_cnt[0]
                        pct_s = 64.0 + 26.0 * (cur_d / total_seizo)
                        hit_info = f"・キャッシュ {cache_hits:,}件" if cache_hits > 0 else ""
                        post_progress(
                            pct_s,
                            f"[5/5] 制作伝票読込中 ({cur_d:,}/{total_seizo:,}{hit_info}): {cur_name}",
                            force=force_update,
                        )

                    if cache_hits > 0:
                        post_progress(
                            64.0 + 26.0 * (cache_hits / total_seizo),
                            f"[5/5] 制作伝票キャッシュ高速復元: {cache_hits:,}/{total_seizo:,}件（新規・更新 {len(xml_todo) + len(xls_todo):,}件）",
                            force=True,
                        )

                    # (5-2) .xlsx / .xlsm 群（および xlrd 利用可能な場合の .xls 群）を高並列スレッドプール + as_completed で処理
                    xlrd_ok = _has_xlrd_available() if xls_todo else False
                    parallel_todo = list(xml_todo)
                    com_xls_todo = []
                    if xls_todo:
                        if xlrd_ok:
                            parallel_todo.extend(xls_todo)
                        else:
                            com_xls_todo = list(xls_todo)

                    if parallel_todo:
                        max_w = min(8, max(4, (os.cpu_count() or 2) * 2), len(parallel_todo))

                        def load_one_entry(seq_entry):
                            seq_i, (fp_i, sz_i, mns_i, msec_i) = seq_entry
                            if is_cancelled():
                                return seq_i, fp_i, sz_i, mns_i, None, InterruptedError()
                            try:
                                rec = read_seizo_denpyo_sheet(
                                    fp_i,
                                    cancel_check=is_cancelled,
                                    base_folders=seizo_folders,
                                    mtime_sec=msec_i,
                                )
                                return seq_i, fp_i, sz_i, mns_i, rec, None
                            except InterruptedError as ie:
                                return seq_i, fp_i, sz_i, mns_i, None, ie
                            except Exception as e:
                                return seq_i, fp_i, sz_i, mns_i, None, e

                        pool = ThreadPoolExecutor(max_workers=max_w)
                        try:
                            futures = [pool.submit(load_one_entry, item) for item in parallel_todo]
                            for fut in as_completed(futures):
                                if is_cancelled():
                                    for f_rem in futures:
                                        f_rem.cancel()
                                    return
                                seq_i, fp_i, sz_i, mns_i, rec, err = fut.result()
                                if isinstance(err, InterruptedError) or is_cancelled():
                                    return
                                if err is not None:
                                    errors.append(self._make_error_entry("制作伝票ファイル", fp_i, err))
                                elif rec:
                                    seizo_records_by_idx[seq_i] = rec
                                    self.seizo_cache.put(fp_i, sz_i, mns_i, rec)
                                with done_lock:
                                    done_cnt[0] += 1
                                    cur_d = done_cnt[0]
                                report_seizo_progress(os.path.basename(fp_i), force_update=(cur_d == total_seizo))
                        finally:
                            pool.shutdown(wait=False, cancel_futures=True)

                    # (5-3) xlrd 無し環境で .xls が残っている場合のみ、1つの Excel COM インスタンスを再利用して一括高速読込
                    if com_xls_todo and not is_cancelled():
                        post_progress(
                            64.0 + 26.0 * (done_cnt[0] / total_seizo),
                            f"[5/5] 旧形式(.xls)制作伝票を一括読込中 ({len(com_xls_todo):,}件)...",
                            force=True,
                        )
                        seq_by_path = {entry[0]: seq_i for seq_i, entry in com_xls_todo}

                        def on_com_item_done(fp_done: str):
                            with done_lock:
                                done_cnt[0] += 1
                                cur_d = done_cnt[0]
                            report_seizo_progress(os.path.basename(fp_done), force_update=(cur_d == total_seizo))

                        batch_res = _read_seizo_xls_batch_com(
                            [entry for _, entry in com_xls_todo],
                            cancel_check=is_cancelled,
                            base_folders=seizo_folders,
                            on_item_done=on_com_item_done,
                        )
                        for entry, rec, err in batch_res:
                            if is_cancelled():
                                return
                            fp_i, sz_i, mns_i, _msec_i = entry
                            if err is not None:
                                errors.append(self._make_error_entry("制作伝票ファイル", fp_i, err))
                            elif rec:
                                seq_i = seq_by_path.get(fp_i, len(seizo_records_by_idx))
                                seizo_records_by_idx[seq_i] = rec
                                self.seizo_cache.put(fp_i, sz_i, mns_i, rec)

                    # 元のソート順で raw_seizo へ格納
                    for seq_i in sorted(seizo_records_by_idx.keys()):
                        raw_seizo.append(seizo_records_by_idx[seq_i])

                    # 更新されたキャッシュをバックグラウンドで安全にディスク保存（参照元保護チェック付き）
                    prot_paths = self.get_protected_source_paths()
                    prot_dirs = self.get_protected_source_dirs()
                    threading.Thread(
                        target=lambda pp=prot_paths, pd=prot_dirs: self.seizo_cache.save_if_dirty(pp, pd),
                        daemon=True,
                    ).start()

                post_progress(90, "全ファイル読込完了。集計データを構築しています...", force=True)

                if is_cancelled():
                    return

                bundle = self._compute_aggregation_bundle(
                    raw_daily,
                    raw_kounyu,
                    raw_gaichu,
                    raw_juchu,
                    code_ranges_snapshot,
                    raw_seizo=raw_seizo,
                    work_rates=work_rates_snapshot,
                )
                post_progress(97, "一覧テーブルを描画しています...", force=True)
                if not is_cancelled():
                    self._ui_queue.put(
                        lambda: self._on_aggregation_finished(
                            job_id, mtime_map, errors, raw_daily, raw_kounyu, raw_gaichu, raw_juchu, raw_seizo, bundle
                        )
                    )
            except Exception as exc:
                if not is_cancelled():
                    self._ui_queue.put(lambda err=str(exc): self._on_aggregation_fatal(job_id, err))

        threading.Thread(target=worker, daemon=True).start()

    def _start_loading_indicator(self, msg: str, pct: float = 0.0):
        """集計・一覧更新の確定進捗バーを表示して開始状態にする"""
        if self._progress_hide_job is not None:
            try:
                self.root.after_cancel(self._progress_hide_job)
            except Exception:
                pass
            self._progress_hide_job = None
        try:
            self.btn_aggregate.configure(text="集計中...")
            if not self.progress_panel.winfo_ismapped():
                self.progress_panel.pack(fill=tk.X, pady=(4, 0))
            self._update_progress_ui(pct, msg)
        except Exception:
            pass

    def _update_progress_ui(self, pct: float, msg: str):
        """進捗バー（0～100%）と進捗テキスト・ステータスバーを同期更新する"""
        if self._closing:
            return
        p_int = int(round(min(max(pct, 0.0), 100.0)))
        self.progress_msg_var.set(msg)
        self.progress_pct_var.set(f"{p_int}%")
        self.status_var.set(f"{msg} ({p_int}%)")
        try:
            self.progress_bar["value"] = p_int
        except Exception:
            pass

    def _stop_loading_indicator(self, completed: bool = False):
        self._is_loading = False
        try:
            self.btn_aggregate.configure(text="集計・一覧更新")
            if completed:
                self.progress_bar["value"] = 100
                self.progress_pct_var.set("100%")
                self.progress_msg_var.set("集計・一覧更新が完了しました")
                if self._progress_hide_job is not None:
                    try:
                        self.root.after_cancel(self._progress_hide_job)
                    except Exception:
                        pass
                def hide_panel():
                    self._progress_hide_job = None
                    if not self._is_loading and not self._closing:
                        try:
                            self.progress_panel.pack_forget()
                        except Exception:
                            pass
                self._progress_hide_job = self.root.after(600, hide_panel)
            else:
                self.progress_panel.pack_forget()
        except Exception:
            pass

    def _on_aggregation_no_files(self, job_id: int, mtime_map: dict, silent_missing: bool):
        if self._closing or job_id != self._active_job_id:
            return
        self._stop_loading_indicator()
        self._apply_mtime_results(mtime_map)
        self.status_var.set("対象ファイルが見つかりませんでした。ファイルを選択してください。")
        if not silent_missing:
            messagebox.showwarning(
                "ファイル未選択",
                "集計対象のファイルが選択されていないか、指定されたパスにファイルが存在しません。\n各ファイルの「選択(複数可)」ボタンからファイルを選択してください。",
                parent=self.root,
            )

    def _on_aggregation_fatal(self, job_id: int, error_msg: str):
        if self._closing or job_id != self._active_job_id:
            return
        self._stop_loading_indicator()
        self.status_var.set(f"集計エラー: {error_msg}")
        messagebox.showerror("集計エラー", f"集計処理中にエラーが発生しました:\n{error_msg}", parent=self.root)

    def _on_aggregation_finished(
        self, job_id: int, mtime_map: dict, errors: list,
        raw_daily: list, raw_kounyu: list, raw_gaichu: list, raw_juchu: list, raw_seizo: list, bundle: dict
    ):
        if self._closing or job_id != self._active_job_id:
            return
        self._stop_loading_indicator(completed=True)
        self._apply_mtime_results(mtime_map)

        self.last_errors = list(errors)
        self._update_error_button_state()

        self.raw_daily = raw_daily
        self.raw_kounyu = raw_kounyu
        self.raw_gaichu = raw_gaichu
        self.raw_juchu = raw_juchu
        self.raw_seizo = raw_seizo

        self.daily_by_kno = bundle["daily_by_kno"]
        self.kounyu_by_kno = bundle["kounyu_by_kno"]
        self.gaichu_by_kno = bundle["gaichu_by_kno"]
        self.juchu_by_kno = bundle["juchu_by_kno"]
        self.seizo_by_kno = bundle["seizo_by_kno"]
        self.aggregated_rows = bundle["aggregated_rows"]

        self.apply_filter_and_refresh()
        self._refresh_current_detail_if_selected()
        err_suffix = f" / エラー: {len(self.last_errors)} 件（エラー確認ボタンから詳細確認可）" if self.last_errors else ""
        self.status_var.set(
            f"集計完了: 全管理No {len(self.aggregated_rows):,} 件 "
            f"(日報: {len(self.raw_daily):,} 行 / 購入品: {len(self.raw_kounyu):,} 行 / "
            f"外注: {len(self.raw_gaichu):,} 行 / 受注表: {len(self.juchu_by_kno):,} 件 / 制作伝票: {len(self.raw_seizo):,} 件){err_suffix}"
        )
        if self.last_errors:
            ErrorReportDialog(self.root, self.last_errors)

    def _refresh_current_detail_if_selected(self):
        """集計・再算出後に現在選択中の管理Noが存在すれば下部内訳明細も最新データへ同期する"""
        kno = self._current_detail_kno
        if not kno:
            return
        if (
            kno in self.daily_by_kno
            or kno in self.kounyu_by_kno
            or kno in self.gaichu_by_kno
            or kno in self.juchu_by_kno
            or kno in self.seizo_by_kno
            or any(r.get("管理No") == kno for r in self.aggregated_rows)
        ):
            self._show_detail_for_kno(kno)
        else:
            self._current_detail_kno = None
            for tv in (self.tree_daily, self.tree_kounyu, self.tree_gaichu, self.tree_juchu, self.tree_seizo):
                tv.delete(*tv.get_children())
            self.detail_notebook.tab(0, text="日報データ明細 (0件)")
            self.detail_notebook.tab(1, text="購入品明細 (0件)")
            self.detail_notebook.tab(2, text="外注加工明細 (0件)")
            self.detail_notebook.tab(4, text="制作伝票明細 (0件)")
            if hasattr(self, "detail_frame"):
                self.detail_frame.configure(text="選択した管理Noの内訳明細")
            if hasattr(self, "detail_bar_info_var"):
                self.detail_bar_info_var.set("一覧から管理Noを選択すると内訳明細が表示されます（Excel出力・印刷対応）")

    def _recalculate_from_loaded_data(self):
        """読み込み済みデータと現在の作業コード設定から全管理Noの集計データを再構築する"""
        bundle = self._compute_aggregation_bundle(
            self.raw_daily,
            self.raw_kounyu,
            self.raw_gaichu,
            self.raw_juchu,
            self.code_ranges,
            raw_seizo=self.raw_seizo,
            work_rates=self.work_rates,
        )
        self.daily_by_kno = bundle["daily_by_kno"]
        self.kounyu_by_kno = bundle["kounyu_by_kno"]
        self.gaichu_by_kno = bundle["gaichu_by_kno"]
        self.juchu_by_kno = bundle["juchu_by_kno"]
        self.seizo_by_kno = bundle["seizo_by_kno"]
        self.aggregated_rows = bundle["aggregated_rows"]
        self.apply_filter_and_refresh()
        self._refresh_current_detail_if_selected()
        self.status_var.set(
            f"集計完了: 全管理No {len(self.aggregated_rows):,} 件 "
            f"(日報: {len(self.raw_daily):,} 行 / 購入品: {len(self.raw_kounyu):,} 行 / "
            f"外注: {len(self.raw_gaichu):,} 行 / 受注表: {len(self.juchu_by_kno):,} 件 / 制作伝票: {len(self.raw_seizo):,} 件)"
        )

    @staticmethod
    def _populate_row_rate_cache(row_dict: dict, rates_tuple: tuple):
        """1行分の円換算キャッシュ（作業費5区分・作業小計・小計1・消耗品費・管理費・小計2・小計1+小計2・諸経費・合計・単価・調整額）を一括計算して格納する"""
        r0, r1, r2, r3, r4 = rates_tuple
        y0 = int(round(to_float(row_dict.get("作業時間", 0.0), 0.0) * r0))
        y1 = int(round(to_float(row_dict.get("移動時間", 0.0), 0.0) * r1))
        y2 = int(round(to_float(row_dict.get("修理改造時間", 0.0), 0.0) * r2))
        y3 = int(round(to_float(row_dict.get("設計関係", 0.0), 0.0) * r3))
        y4 = int(round(to_float(row_dict.get("その他作業", 0.0), 0.0) * r4))
        y_sub = y0 + y1 + y2 + y3 + y4
        y_adj = int(round(to_float(row_dict.get("調整用", 0.0), 0.0)))
        y_ksub = int(round(to_float(row_dict.get("購入品小計", 0.0), 0.0)))
        y_mat = int(round(to_float(row_dict.get("材料費計", 0.0), 0.0)))
        y_des = int(round(to_float(row_dict.get("設計費", 0.0), 0.0)))
        y_sub1 = y_sub + y_adj + y_ksub + y_mat + y_des
        y_shoumou = int(round((y_sub + y_mat + y_des) * 0.06))
        y_juki = int(round(to_float(row_dict.get("重機、運搬費", 0.0), 0.0)))
        y_sharyo = int(round(to_float(row_dict.get("車両、交通費", 0.0), 0.0)))
        y_shucho = int(round(to_float(row_dict.get("出張費", 0.0), 0.0)))
        y_kanri = int(round(y_sub1 * 0.02))
        y_sub2 = y_shoumou + y_juki + y_sharyo + y_shucho + y_kanri
        y_sub12 = y_sub1 + y_sub2
        y_shokeihi = int(round(y_sub12 * 0.08))
        y_total = y_sub12 + y_shokeihi
        seizo_qty = to_float(row_dict.get("_制作伝票数量", 0.0), 0.0)
        y_total_per_qty = int(round(y_total / seizo_qty)) if (seizo_qty > 0 and y_total != 0) else 0
        y_seikyu = int(round(to_float(row_dict.get("総請求金額", 0.0), 0.0)))
        y_choseigaku = y_seikyu - y_total
        y_seikyu_per_qty = int(round(y_seikyu / seizo_qty)) if (seizo_qty > 0 and y_seikyu != 0) else 0
        row_dict["_rates_key"] = rates_tuple
        row_dict["_yen_作業時間"] = float(y0)
        row_dict["_yen_移動時間"] = float(y1)
        row_dict["_yen_修理改造時間"] = float(y2)
        row_dict["_yen_設計関係"] = float(y3)
        row_dict["_yen_その他作業"] = float(y4)
        row_dict["_yen_作業小計"] = float(y_sub)
        row_dict["_yen_小計1"] = float(y_sub1)
        row_dict["_yen_消耗品費(6%)"] = float(y_shoumou)
        row_dict["_yen_管理費"] = float(y_kanri)
        row_dict["_yen_小計2"] = float(y_sub2)
        row_dict["_yen_小計1+小計2"] = float(y_sub12)
        row_dict["_yen_諸経費(8%)"] = float(y_shokeihi)
        row_dict["_yen_合計"] = float(y_total)
        row_dict["_yen_合計/数量"] = float(y_total_per_qty)
        row_dict["_yen_調整額"] = float(y_choseigaku)
        row_dict["_yen_総請求金額/数量"] = float(y_seikyu_per_qty)

    @staticmethod
    def _compute_aggregation_bundle(
        raw_daily: list,
        raw_kounyu: list,
        raw_gaichu: list,
        raw_juchu: list,
        code_ranges: dict,
        raw_seizo: list = None,
        work_rates: dict = None,
    ) -> dict:
        """ワーカースレッドからも安全に呼べる高速集計処理（検索用テキストと円換算キャッシュも事前構築）"""
        compiled_ranges = {
            cat: parse_code_range_spec(code_ranges.get(cat, ""))
            for cat in WORK_CATEGORIES
        }
        rates_src = work_rates or DEFAULT_WORK_RATES
        rates_tuple = tuple(
            to_float(rates_src.get(cat, DEFAULT_WORK_RATES[cat]), DEFAULT_WORK_RATES[cat])
            for cat in WORK_CATEGORIES
        )

        daily_by_kno = defaultdict(list)
        kounyu_by_kno = defaultdict(list)
        gaichu_by_kno = defaultdict(list)
        juchu_by_kno = {}
        seizo_by_kno = defaultdict(list)

        all_knos = set()
        fallback_meta = {}

        # (A) 受注表データ整理（複数ファイルがある場合は空でない値で補完・マージ）
        for r in raw_juchu:
            kno = normalize_kanri_no(r.get("B"))
            if not kno or kno == "管理No":
                continue
            all_knos.add(kno)
            if kno not in juchu_by_kno:
                juchu_by_kno[kno] = dict(r)
            else:
                merged = juchu_by_kno[kno]
                for col_k, col_v in r.items():
                    if col_v:
                        merged[col_k] = col_v

        # (B) 日報データ集計
        work_sums = defaultdict(lambda: {cat: 0.0 for cat in WORK_CATEGORIES})
        for rec in raw_daily:
            kno = rec["管理No"]
            if not kno:
                continue
            all_knos.add(kno)
            code_str = rec.get("作業コード", "")
            mins = rec.get("合計", 0.0)
            cat_matched = ""
            try:
                code_int = int(float(code_str))
                cat_matched = match_code_category(code_int, compiled_ranges)
            except ValueError:
                pass
            if cat_matched:
                work_sums[kno][cat_matched] += mins

            rec_with_cat = dict(rec)
            rec_with_cat["分類区分"] = cat_matched if cat_matched else "対象外"
            daily_by_kno[kno].append(rec_with_cat)

            if kno not in fallback_meta:
                fallback_meta[kno] = {
                    "出荷先": rec.get("出荷先", ""),
                    "プロジェクト名": rec.get("プロジェクト名", ""),
                    "品名、ライン名等": rec.get("品名、ライン名等", ""),
                    "製品詳細": rec.get("製品詳細", ""),
                    "数量": rec.get("数量", ""),
                    "単位": rec.get("単位", ""),
                }

        # (C) 購入品データ集計（Q列: 金額 単価X必要数、空欄時はM×N補完）
        kounyu_sums = defaultdict(float)
        for r in raw_kounyu:
            kno = normalize_kanri_no(r.get("A"))
            if not kno or kno == "管理No":
                continue
            all_knos.add(kno)
            amount = _calc_kounyu_row_amount(r)
            kounyu_sums[kno] += amount
            kounyu_by_kno[kno].append(r)

            if kno not in fallback_meta:
                fallback_meta[kno] = {
                    "出荷先": r.get("B", ""),
                    "プロジェクト名": "",
                    "品名、ライン名等": r.get("C", ""),
                    "製品詳細": r.get("D", ""),
                    "数量": r.get("E", ""),
                    "単位": "",
                }

        # (D) 外注加工データ集計（Q列: 金額 単価X必要数、空欄時はM×N補完）
        gaichu_sums = defaultdict(float)
        for r in raw_gaichu:
            kno = normalize_kanri_no(r.get("A"))
            if not kno or kno == "管理No":
                continue
            all_knos.add(kno)
            amount = _calc_kounyu_row_amount(r)
            gaichu_sums[kno] += amount
            gaichu_by_kno[kno].append(r)

            if kno not in fallback_meta:
                fallback_meta[kno] = {
                    "出荷先": r.get("B", ""),
                    "プロジェクト名": "",
                    "品名、ライン名等": r.get("C", ""),
                    "製品詳細": r.get("D", ""),
                    "数量": r.get("E", ""),
                    "単位": "",
                }

        # (E) 制作伝票（個別Excel）集計
        seizo_sums = defaultdict(lambda: {c: 0.0 for c in SEIZO_COST_COLUMNS})
        seizo_qty_sums = defaultdict(float)
        for srec in (raw_seizo or []):
            kno = srec.get("管理No", "")
            if not kno:
                continue
            all_knos.add(kno)
            seizo_by_kno[kno].append(srec)
            for c in SEIZO_COST_COLUMNS:
                seizo_sums[kno][c] += to_float(srec.get(c, 0.0), 0.0)
            seizo_qty_sums[kno] += to_float(srec.get("_制作伝票数量", 0.0), 0.0)
            s_meta = srec.get("_meta", {})
            if kno not in fallback_meta:
                fallback_meta[kno] = {
                    "出荷先": s_meta.get("出荷先", ""),
                    "プロジェクト名": s_meta.get("プロジェクト名", ""),
                    "品名、ライン名等": s_meta.get("品名、ライン名等", ""),
                    "製品詳細": s_meta.get("製品詳細", ""),
                    "数量": s_meta.get("数量", ""),
                    "単位": s_meta.get("単位", ""),
                    "売り上げ区分": s_meta.get("売り上げ区分", ""),
                    "受注先": s_meta.get("受注先", ""),
                    "受注日": s_meta.get("受注日", ""),
                    "出荷日": s_meta.get("出荷日", ""),
                    "請求完了日": s_meta.get("請求完了日", ""),
                    "備考": s_meta.get("備考", ""),
                }

        # (F) 統合レコード作成
        aggregated_rows = []
        populate_cache = DenpyoHubApp._populate_row_rate_cache
        for kno in sorted(all_knos):
            j = juchu_by_kno.get(kno, {})
            fb = fallback_meta.get(kno, {})
            w = work_sums[kno]
            sz = seizo_sums[kno]

            w_sagyo = w["作業時間"]
            w_idou = w["移動時間"]
            w_shuri = w["修理改造時間"]
            w_sekkei = w["設計関係"]
            w_sonota = w["その他作業"]
            w_subtotal = w_sagyo + w_idou + w_shuri + w_sekkei + w_sonota

            c_kounyu = kounyu_sums[kno]
            c_gaichu = gaichu_sums[kno]
            c_subtotal = c_kounyu + c_gaichu

            v_juchu_bi = format_excel_date(j.get("A", "")) or fb.get("受注日", "")
            v_juchu_saki = j.get("C", "") or fb.get("受注先", "")
            v_shukka_saki = j.get("E", "") or fb.get("出荷先", "")
            v_proj = j.get("G", "") or fb.get("プロジェクト名", "")
            v_hinmei = j.get("H", "") or fb.get("品名、ライン名等", "")
            v_detail = j.get("I", "") or fb.get("製品詳細", "")
            v_qty = j.get("J", "") or fb.get("数量", "")
            v_unit = j.get("K", "") or fb.get("単位", "")
            v_tehai = j.get("M", "")
            v_uriage = j.get("U", "") or fb.get("売り上げ区分", "")
            v_shukka_bi = format_excel_date(j.get("R", "")) or fb.get("出荷日", "")
            v_seikyu_bi = format_excel_date(j.get("T", "")) or fb.get("請求完了日", "")
            v_bikou = j.get("V", "") or fb.get("備考", "")

            has_seizo_flag = bool(kno in seizo_by_kno or any(abs(sz[c]) > 1e-9 for c in SEIZO_COST_COLUMNS))
            search_haystack = normalize_search_text(
                f"{kno} {v_juchu_saki} {v_shukka_saki} {v_proj} {v_hinmei} {v_detail} {v_tehai} {v_uriage} {v_bikou}"
            )

            row_dict = {
                "管理No": kno,
                "受注日": v_juchu_bi,
                "受注先": v_juchu_saki,
                "出荷先": v_shukka_saki,
                "プロジェクト名": v_proj,
                "品名、ライン名等": v_hinmei,
                "製品詳細": v_detail,
                "数量": v_qty,
                "単位": v_unit,
                "製造又は手配先": v_tehai,
                "売り上げ区分": v_uriage,
                "出荷日": v_shukka_bi,
                "請求完了日": v_seikyu_bi,
                "備考": v_bikou,
                "作業時間": w_sagyo,
                "移動時間": w_idou,
                "修理改造時間": w_shuri,
                "設計関係": w_sekkei,
                "その他作業": w_sonota,
                "作業小計": w_subtotal,
                "購入品": c_kounyu,
                "外注加工": c_gaichu,
                "購入品小計": c_subtotal,
                "調整用": sz["調整用"],
                "材料費１(型鋼)": sz["材料費１(型鋼)"],
                "材料費２(鉄板)": sz["材料費２(鉄板)"],
                "材料費３(コラム)": sz["材料費３(コラム)"],
                "その他材料費(31行)": sz["その他材料費(31行)"],
                "その他材料費(32行)": sz["その他材料費(32行)"],
                "その他材料費(33行)": sz["その他材料費(33行)"],
                "切断材料": sz["切断材料"],
                "材料費計": sz["材料費計"],
                "設計費": sz["設計費"],
                "小計1": 0.0,
                "消耗品費(6%)": 0.0,
                "重機、運搬費": sz["重機、運搬費"],
                "車両、交通費": sz["車両、交通費"],
                "出張費": sz["出張費"],
                "管理費": 0.0,
                "小計2": 0.0,
                "小計1+小計2": 0.0,
                "諸経費(8%)": 0.0,
                "合計": 0.0,
                "合計/数量": 0.0,
                "調整額": 0.0,
                "総請求金額": sz["総請求金額"],
                "総請求金額/数量": 0.0,
                "_制作伝票数量": seizo_qty_sums[kno],
                "_has_seizo": has_seizo_flag,
                "_search_text": search_haystack,
            }
            populate_cache(row_dict, rates_tuple)
            aggregated_rows.append(row_dict)

        return {
            "daily_by_kno": daily_by_kno,
            "kounyu_by_kno": kounyu_by_kno,
            "gaichu_by_kno": gaichu_by_kno,
            "juchu_by_kno": juchu_by_kno,
            "seizo_by_kno": seizo_by_kno,
            "aggregated_rows": aggregated_rows,
        }

    def _get_rates_tuple(self) -> tuple:
        r = self.work_rates
        v0 = r.get("作業時間", 62.5)
        v1 = r.get("移動時間", 50.0)
        v2 = r.get("修理改造時間", 70.0)
        v3 = r.get("設計関係", 62.5)
        v4 = r.get("その他作業", 62.5)
        return (
            float(v0) if isinstance(v0, (int, float)) and math.isfinite(v0) else to_float(v0, 62.5),
            float(v1) if isinstance(v1, (int, float)) and math.isfinite(v1) else to_float(v1, 50.0),
            float(v2) if isinstance(v2, (int, float)) and math.isfinite(v2) else to_float(v2, 70.0),
            float(v3) if isinstance(v3, (int, float)) and math.isfinite(v3) else to_float(v3, 62.5),
            float(v4) if isinstance(v4, (int, float)) and math.isfinite(v4) else to_float(v4, 62.5),
        )

    def _get_work_numeric_val(self, row: dict, col: str, rates_tuple: tuple = None, unit: str = None) -> float:
        """表示単位（分・時間・工数・金額）と単価レートを反映した作業列および派生円算出列の数値を高速に返す"""
        if rates_tuple is None:
            rates_tuple = self._get_rates_tuple()
        if row.get("_rates_key") != rates_tuple:
            self._populate_row_rate_cache(row, rates_tuple)

        if col in DERIVED_YEN_COLUMNS:
            return row.get(f"_yen_{col}", 0.0)
        if unit is None:
            unit = self.time_unit_var.get()
        if unit == "金額(円)":
            return row.get(f"_yen_{col}", 0.0)
        raw_v = row.get(col, 0.0)
        minutes_val = float(raw_v) if isinstance(raw_v, (int, float)) and math.isfinite(raw_v) else to_float(raw_v, 0.0)
        if unit == "時間(分÷60)":
            return round(minutes_val / 60.0, 2)
        if unit == "工数(8h=1)":
            return round(minutes_val / KOUSU_MINUTES, 2)
        return minutes_val

    def _format_converted_work_val(self, val: float) -> str:
        """表示単位に変換済みの作業数値（分・時間・工数・円）を文字列フォーマットする"""
        unit = self.time_unit_var.get()
        if unit == "金額(円)":
            return f"{int(round(val)):,}"
        if unit in ("時間(分÷60)", "工数(8h=1)"):
            return f"{val:,.2f}"
        if abs(val - round(val)) < 1e-6:
            return f"{int(round(val)):,}"
        return f"{val:,.1f}"

    def _get_cell_display_str(self, row: dict, col: str, unit: str = None, rates_tuple: tuple = None) -> str:
        """一覧表示上のセル文字列を取得する（視認性向上のため単位付き表示・0はシンプル表示）"""
        if unit is None:
            unit = self.time_unit_var.get()
        if col in DERIVED_YEN_COLUMNS:
            iv = int(round(self._get_work_numeric_val(row, col, rates_tuple=rates_tuple, unit=unit)))
            return f"{iv:,}円" if iv != 0 else "0"
        if col in WORK_ONLY_COLUMNS:
            val = self._get_work_numeric_val(row, col, rates_tuple=rates_tuple, unit=unit)
            if unit == "金額(円)":
                iv = int(round(val))
                return f"{iv:,}円" if iv != 0 else "0"
            if unit == "時間(分÷60)":
                return f"{val:,.2f}h" if abs(val) >= 1e-6 else "0.00"
            if unit == "工数(8h=1)":
                return f"{val:,.2f}工数" if abs(val) >= 1e-6 else "0.00"
            if abs(val) < 1e-6:
                return "0"
            if abs(val - round(val)) < 1e-6:
                return f"{int(round(val)):,}分"
            return f"{val:,.1f}分"
        elif col in COST_YEN_COLUMNS:
            iv = int(round(row.get(col, 0.0)))
            return f"{iv:,}円" if iv != 0 else "0"
        return str(row.get(col, ""))

    def _get_cell_filter_key(self, row: dict, col: str, unit: str = None, rates_tuple: tuple = None) -> str:
        """Excel風フィルター用の正規化値（数値列はカンマ無しの表示値、文字はそのまま）"""
        if col in DERIVED_YEN_COLUMNS:
            return str(int(round(self._get_work_numeric_val(row, col, rates_tuple=rates_tuple, unit=unit))))
        if col in WORK_ONLY_COLUMNS:
            if unit is None:
                unit = self.time_unit_var.get()
            val = self._get_work_numeric_val(row, col, rates_tuple=rates_tuple, unit=unit)
            if unit == "金額(円)":
                return str(int(round(val)))
            if unit in ("時間(分÷60)", "工数(8h=1)"):
                return f"{val:.2f}"
            if abs(val - round(val)) < 1e-6:
                return str(int(round(val)))
            return f"{val:.1f}"
        elif col in COST_YEN_COLUMNS:
            return str(int(round(row.get(col, 0.0))))
        return str(row.get(col, ""))

    def _row_matches_base_conditions(
        self,
        r: dict,
        exclude_col: str = None,
        mode: str = None,
        kw_tokens: list = None,
        unit: str = None,
        rates_tuple: tuple = None,
    ) -> bool:
        """全体検索・表示対象モード・および指定列以外の全列フィルターに一致するか高速判定"""
        w_sub = r["作業小計"]
        c_sub = r["購入品小計"]
        has_seizo = r.get("_has_seizo")
        if has_seizo is None:
            has_seizo = any(abs(r.get(c, 0.0)) > 1e-9 for c in SEIZO_COST_COLUMNS)
        if mode is None:
            mode = self.filter_mode_var.get()
        if mode == "実績ありのみ" and (w_sub <= 0 and c_sub <= 0 and not has_seizo):
            return False
        elif mode == "作業実績ありのみ" and w_sub <= 0:
            return False
        elif mode == "購入・外注実績ありのみ" and c_sub <= 0:
            return False
        elif mode == "作業＋購入・外注の両方あり" and (w_sub <= 0 or c_sub <= 0):
            return False

        if kw_tokens is None:
            kw = normalize_search_text(self.search_var.get())
            kw_tokens = [t for t in kw.split() if t] if kw else []
        if kw_tokens:
            haystack = r.get("_search_text")
            if haystack is None:
                haystack = normalize_search_text(" ".join([
                    str(r.get("管理No", "")),
                    str(r.get("受注先", "")),
                    str(r.get("出荷先", "")),
                    str(r.get("プロジェクト名", "")),
                    str(r.get("品名、ライン名等", "")),
                    str(r.get("製品詳細", "")),
                    str(r.get("製造又は手配先", "")),
                    str(r.get("売り上げ区分", "")),
                    str(r.get("備考", "")),
                ]))
                r["_search_text"] = haystack
            if not all(t in haystack for t in kw_tokens):
                return False

        for f_col, f_spec in self.column_filters.items():
            if f_col == exclude_col:
                continue
            if not self._row_matches_column_filter(r, f_col, f_spec, unit=unit, rates_tuple=rates_tuple):
                return False
        return True

    def _row_matches_column_filter(self, r: dict, col: str, spec: dict, unit: str = None, rates_tuple: tuple = None) -> bool:
        """単一列のExcel風フィルター条件判定"""
        sel = spec.get("selected")
        if sel is not None:
            key_str = self._get_cell_filter_key(r, col, unit=unit, rates_tuple=rates_tuple)
            if key_str not in sel:
                return False

        num_op = spec.get("num_op", "指定なし")
        if col in NUMERIC_COLUMNS and num_op != "指定なし":
            if col in RATE_DEPENDENT_COLUMNS:
                val = self._get_work_numeric_val(r, col, rates_tuple=rates_tuple, unit=unit)
            else:
                val = to_float(r.get(col, 0.0), 0.0)
            v1 = spec.get("num_v1")
            v2 = spec.get("num_v2")
            if num_op == "0より大きい (>0)" and not (val > 0):
                return False
            elif num_op == "0のみ (=0)" and not (abs(val) < 1e-9):
                return False
            elif num_op == "以上 (>=)" and v1 is not None and not (val >= v1):
                return False
            elif num_op == "以下 (<=)" and v1 is not None and not (val <= v1):
                return False
            elif num_op == "より大きい (>)" and v1 is not None and not (val > v1):
                return False
            elif num_op == "未満 (<)" and v1 is not None and not (val < v1):
                return False
            elif num_op == "範囲 (最小～最大)" and v1 is not None and v2 is not None and not (v1 <= val <= v2):
                return False
        return True

    def open_excel_filter_popup(self, col_name: str, x_root: int, y_root: int):
        """指定列のExcel風フィルターポップアップを開く"""
        if not self.aggregated_rows:
            messagebox.showinfo("データなし", "先に「集計・一覧更新」を実行してデータを読み込んでください。", parent=self.root)
            return

        # 他の列の条件を満たす行からユニーク値と件数を集計（Excelと同じカスケード動作・ループ外事前計算で高速化）
        mode_now = self.filter_mode_var.get()
        kw_now = normalize_search_text(self.search_var.get())
        kw_tokens = [t for t in kw_now.split() if t] if kw_now else []
        unit_now = self.time_unit_var.get()
        rates_now = self._get_rates_tuple()
        counts = Counter()
        for r in self.aggregated_rows:
            if self._row_matches_base_conditions(
                r,
                exclude_col=col_name,
                mode=mode_now,
                kw_tokens=kw_tokens,
                unit=unit_now,
                rates_tuple=rates_now,
            ):
                counts[self._get_cell_filter_key(r, col_name, unit=unit_now, rates_tuple=rates_now)] += 1

        # すでに選択中のフィルター値が件数0になっていても復元できるよう含める
        cur_f = self.column_filters.get(col_name)
        if cur_f and cur_f.get("selected"):
            for v in cur_f["selected"]:
                if v not in counts:
                    counts[v] = 0

        disp_col = get_display_column_label(col_name, unit_now)
        ExcelFilterPopup(
            parent=self.root,
            col_name=col_name,
            x_root=x_root,
            y_root=y_root,
            value_counts=counts,
            is_numeric=(col_name in NUMERIC_COLUMNS),
            current_filter=cur_f,
            on_sort=self.set_sort,
            on_apply_filter=self._apply_single_filter,
            on_clear_filter=self._clear_single_filter,
            display_col_name=disp_col,
        )

    def _apply_single_filter(self, col_name: str, spec: dict):
        self.column_filters[col_name] = spec
        self.apply_filter_and_refresh()

    def _clear_single_filter(self, col_name: str):
        if col_name in self.column_filters:
            del self.column_filters[col_name]
        self.apply_filter_and_refresh()

    def clear_all_column_filters(self):
        self.column_filters.clear()
        self.search_var.set("")
        self._last_search_kw = ""
        self.apply_filter_and_refresh()

    def _update_filter_status_label(self):
        if not self.column_filters:
            self.filter_status_var.set("列フィルター: なし")
            self.lbl_filter_status.configure(style="Badge.TLabel")
        else:
            unit_now = self.time_unit_var.get()
            desc = ", ".join(get_display_column_label(k, unit_now) for k in self.column_filters.keys())
            if len(desc) > 28:
                desc = desc[:25] + "..."
            self.filter_status_var.set(f"列絞込中({len(self.column_filters)}列): {desc}")
            self.lbl_filter_status.configure(style="FilterActive.TLabel")

    def apply_filter_and_refresh(self, save_cfg: bool = False):
        """検索キーワード・表示対象・Excel風列フィルターを適用し、左右分割ペインを高速更新する"""
        if save_cfg:
            self.save_config()
        self._update_filter_status_label()
        self._update_unit_button_styles()

        mode_now = self.filter_mode_var.get()
        kw_now = normalize_search_text(self.search_var.get())
        kw_tokens = [t for t in kw_now.split() if t] if kw_now else []
        unit = self.time_unit_var.get()
        rates_tuple = self._get_rates_tuple()

        filtered = [
            r for r in self.aggregated_rows
            if self._row_matches_base_conditions(r, mode=mode_now, kw_tokens=kw_tokens, unit=unit, rates_tuple=rates_tuple)
        ]
        self.displayed_rows = filtered

        is_yen_mode = (unit == "金額(円)")
        totals = {c: 0.0 for c in CALC_COLUMNS}
        populate_cache = self._populate_row_rate_cache

        for r in self.displayed_rows:
            if r.get("_rates_key") != rates_tuple:
                populate_cache(r, rates_tuple)
            if is_yen_mode:
                totals["作業時間"] += r["_yen_作業時間"]
                totals["移動時間"] += r["_yen_移動時間"]
                totals["修理改造時間"] += r["_yen_修理改造時間"]
                totals["設計関係"] += r["_yen_設計関係"]
                totals["その他作業"] += r["_yen_その他作業"]
                totals["作業小計"] += r["_yen_作業小計"]
            else:
                for c in WORK_CATEGORIES:
                    totals[c] += r[c]
                totals["作業小計"] += r["作業小計"]
            for c in COST_YEN_COLUMNS:
                totals[c] += r.get(c, 0.0)
            for c in DERIVED_YEN_COLUMNS:
                totals[c] += r.get(f"_yen_{c}", 0.0)

        if is_yen_mode:
            self.summary_var.set(
                f"表示件数: {len(self.displayed_rows):,} 件   |   "
                f"作業費: {int(round(totals['作業時間'])):,}円   "
                f"移動費: {int(round(totals['移動時間'])):,}円   "
                f"修理改造費: {int(round(totals['修理改造時間'])):,}円   "
                f"設計(時間分): {int(round(totals['設計関係'])):,}円   "
                f"その他作業費: {int(round(totals['その他作業'])):,}円   "
                f"【作業費小計: {int(round(totals['作業小計'])):,}円】   |   "
                f"【購入品小計: {int(round(totals['購入品小計'])):,}円】   |   "
                f"調整用: {int(round(totals['調整用'])):,}円   "
                f"材料費計: {int(round(totals['材料費計'])):,}円   "
                f"設計費: {int(round(totals['設計費'])):,}円   |   "
                f"【小計1: {int(round(totals['小計1'])):,}円】   "
                f"【小計2: {int(round(totals['小計2'])):,}円】   "
                f"【合計: {int(round(totals['合計'])):,}円】   "
                f"【総請求金額: {int(round(totals['総請求金額'])):,}円】"
            )
        else:
            if unit == "時間(分÷60)":
                unit_label = "h"
                fmt = lambda v: f"{v / 60.0:,.2f}"
                l_sagyo, l_idou, l_shuri, l_sekkei, l_sonota, l_wsub = (
                    "作業時間", "移動時間", "修理改造時間", "設計関係", "その他作業", "作業小計"
                )
            elif unit == "工数(8h=1)":
                unit_label = "工数"
                fmt = lambda v: f"{v / KOUSU_MINUTES:,.2f}"
                l_sagyo, l_idou, l_shuri, l_sekkei, l_sonota, l_wsub = (
                    "作業工数", "移動工数", "修理改造工数", "設計工数", "その他工数", "作業工数小計"
                )
            else:
                unit_label = "分"
                fmt = self._format_converted_work_val
                l_sagyo, l_idou, l_shuri, l_sekkei, l_sonota, l_wsub = (
                    "作業時間", "移動時間", "修理改造時間", "設計関係", "その他作業", "作業小計"
                )
            self.summary_var.set(
                f"表示件数: {len(self.displayed_rows):,} 件   |   "
                f"{l_sagyo}: {fmt(totals['作業時間'])}{unit_label}   "
                f"{l_idou}: {fmt(totals['移動時間'])}{unit_label}   "
                f"{l_shuri}: {fmt(totals['修理改造時間'])}{unit_label}   "
                f"{l_sekkei}: {fmt(totals['設計関係'])}{unit_label}   "
                f"{l_sonota}: {fmt(totals['その他作業'])}{unit_label}   "
                f"【{l_wsub}: {fmt(totals['作業小計'])}{unit_label}】   |   "
                f"【購入品小計: {int(round(totals['購入品小計'])):,}円】   |   "
                f"調整用: {int(round(totals['調整用'])):,}円   "
                f"材料費計: {int(round(totals['材料費計'])):,}円   "
                f"設計費: {int(round(totals['設計費'])):,}円   |   "
                f"【小計1: {int(round(totals['小計1'])):,}円】   "
                f"【小計2: {int(round(totals['小計2'])):,}円】   "
                f"【合計: {int(round(totals['合計'])):,}円】   "
                f"【総請求金額: {int(round(totals['総請求金額'])):,}円】"
            )

        self._refresh_both_panes_sorted()

    def _refresh_both_panes_sorted(self):
        """行リンクON時は1回だけソートして左右両ペインへ渡し、2重ソートを省いて高速化する"""
        if self.link_rows_var.get():
            col = self.sort_col
            desc = self.sort_desc
            if col in RATE_DEPENDENT_COLUMNS:
                get_w = self._get_work_numeric_val
                rates_tuple = self._get_rates_tuple()
                unit_now = self.time_unit_var.get()
                sort_key = lambda item: get_w(item, col, rates_tuple=rates_tuple, unit=unit_now)
            elif col in COST_YEN_COLUMNS:
                sort_key = lambda item: item.get(col, 0.0)
            elif col in NUMERIC_COLUMNS:
                sort_key = lambda item: to_float(item.get(col, 0.0), 0.0)
            else:
                sort_key = lambda item: str(item.get(col, ""))
            sorted_rows = list(self.displayed_rows)
            sorted_rows.sort(key=sort_key, reverse=desc)
            self.meta_pane.refresh_rows(sorted_rows, pre_sorted=True)
            self.calc_pane.refresh_rows(sorted_rows, pre_sorted=True)
        else:
            self.meta_pane.refresh_rows(self.displayed_rows, pre_sorted=False)
            self.calc_pane.refresh_rows(self.displayed_rows, pre_sorted=False)

    def set_sort(self, col: str, desc: bool):
        self.sort_col = col
        self.sort_desc = desc
        if self.link_rows_var.get():
            self.meta_pane.sort_col = col
            self.meta_pane.sort_desc = desc
            self.calc_pane.sort_col = col
            self.calc_pane.sort_desc = desc
            self._refresh_both_panes_sorted()
            return
        if col in CALC_COLUMNS:
            self.calc_pane.set_sort(col, desc, sync=False)
        elif col == "管理No":
            self._get_active_pane().set_sort(col, desc, sync=False)
        else:
            self.meta_pane.set_sort(col, desc, sync=False)

    def shift_column(self, col_name: str, delta: int, pane_id: str = None):
        """指定列を同一ペイン内で左右に1つ移動し、設定に保存する"""
        if pane_id == "calc" or col_name in CALC_COLUMNS:
            pane_cols = [c for c in self.column_order if c in self.visible_set and c in CALC_COLUMNS]
        else:
            pane_cols = [c for c in self.column_order if c in self.visible_set and c not in CALC_COLUMNS]
        if col_name not in pane_cols:
            return
        vis_idx = pane_cols.index(col_name)
        new_vis_idx = vis_idx + delta
        if not (0 <= new_vis_idx < len(pane_cols)):
            return
        target_col = pane_cols[new_vis_idx]
        i1 = self.column_order.index(col_name)
        i2 = self.column_order.index(target_col)
        self.column_order[i1], self.column_order[i2] = self.column_order[i2], self.column_order[i1]
        self.visible_columns = [c for c in self.column_order if c in self.visible_set]
        self._configure_main_tree_columns()
        self.apply_filter_and_refresh(save_cfg=True)

    def open_column_order_dialog(self):
        """列の順番・表示設定ダイアログを開く"""
        self._capture_current_column_widths()

        def on_apply(new_order, new_visible_set):
            self.column_order = new_order
            self.visible_set = new_visible_set
            self.visible_columns = [c for c in self.column_order if c in self.visible_set]
            self._configure_main_tree_columns()
            self.apply_filter_and_refresh(save_cfg=True)

        ColumnOrderDialog(
            self.root,
            self.column_order,
            self.visible_set,
            on_apply,
            display_unit=self.time_unit_var.get(),
        )

    def autofit_all_columns(self):
        """左右両ペインの全列幅を自動調整する"""
        for col in self.meta_pane.visible_columns:
            self.meta_pane.autofit_column(col, save_after=False)
        for col in self.calc_pane.visible_columns:
            self.calc_pane.autofit_column(col, save_after=False)
        self.save_config()
        self.meta_pane._schedule_overlay_update()
        self.calc_pane._schedule_overlay_update()
        self.status_var.set("左右すべての列幅を自動調整しました。")

    def _show_detail_for_kno(self, kno: str):
        """指定した管理Noの内訳明細を下部パネルの各タブに表示する"""
        self._current_detail_kno = kno
        # 1. 日報明細
        self.tree_daily.delete(*self.tree_daily.get_children())
        d_list = self.daily_by_kno.get(kno, [])
        d_total_mins = 0.0
        d_total_yen = 0
        for i, d in enumerate(d_list):
            mins = d.get("合計", 0.0)
            d_total_mins += mins
            upd = d.get("更新日", "")
            if "." in upd:
                upd = upd.split(".")[0]
            cat = d.get("分類区分", "")
            if cat in self.work_rates:
                rate = to_float(self.work_rates.get(cat, DEFAULT_WORK_RATES.get(cat, 0.0)), 0.0)
                yen_i = int(round(mins * rate))
                d_total_yen += yen_i
                yen_str = f"{yen_i:,}"
            else:
                yen_str = "-"
            vals = (
                upd,
                d.get("作業者", ""),
                d.get("作業コード", ""),
                d.get("作業名", ""),
                cat,
                f"{mins:,.0f}",
                f"{mins / 60.0:,.2f}",
                f"{mins / KOUSU_MINUTES:,.2f}",
                yen_str,
                d.get("出荷先", ""),
                d.get("製品詳細", ""),
                d.get("ファイル名", ""),
            )
            self.tree_daily.insert("", tk.END, values=vals, tags=("even" if i % 2 == 0 else "odd",))
        self.detail_notebook.tab(0, text=f"日報データ明細 ({len(d_list)}件)")

        # 2. 購入品明細
        self.tree_kounyu.delete(*self.tree_kounyu.get_children())
        k_list = self.kounyu_by_kno.get(kno, [])
        k_total_yen = 0
        for i, r in enumerate(k_list):
            q_num = _calc_kounyu_row_amount(r)
            k_total_yen += int(round(q_num))
            m_num = to_float(r.get("M"), 0.0)
            vals = (
                r.get("F", ""),
                r.get("G", ""),
                r.get("H", ""),
                r.get("I", ""),
                r.get("K", ""),
                r.get("L", ""),
                f"{int(round(m_num)):,}" if r.get("M") else "",
                r.get("N", ""),
                f"{int(round(q_num)):,}",
                format_excel_date(r.get("U", "")),
                format_excel_date(r.get("W", "")),
                r.get("T", ""),
                r.get("_src_file", ""),
            )
            self.tree_kounyu.insert("", tk.END, values=vals, tags=("even" if i % 2 == 0 else "odd",))
        self.detail_notebook.tab(1, text=f"購入品明細 ({len(k_list)}件)")

        # 3. 外注加工明細
        self.tree_gaichu.delete(*self.tree_gaichu.get_children())
        g_list = self.gaichu_by_kno.get(kno, [])
        g_total_yen = 0
        for i, r in enumerate(g_list):
            q_num = _calc_kounyu_row_amount(r)
            g_total_yen += int(round(q_num))
            m_num = to_float(r.get("M"), 0.0)
            vals = (
                r.get("F", ""),
                r.get("G", ""),
                r.get("H", ""),
                r.get("I", ""),
                r.get("K", ""),
                r.get("L", ""),
                f"{int(round(m_num)):,}" if r.get("M") else "",
                r.get("N", ""),
                f"{int(round(q_num)):,}",
                format_excel_date(r.get("U", "")),
                format_excel_date(r.get("W", "")),
                r.get("T", ""),
                r.get("_src_file", ""),
            )
            self.tree_gaichu.insert("", tk.END, values=vals, tags=("even" if i % 2 == 0 else "odd",))
        self.detail_notebook.tab(2, text=f"外注加工明細 ({len(g_list)}件)")

        # 4. 受注表詳細
        self.tree_juchu.delete(*self.tree_juchu.get_children())
        j = self.juchu_by_kno.get(kno, {})
        pairs = [
            ("管理No", kno, "受注日", format_excel_date(j.get("A", ""))),
            ("受注先", j.get("C", ""), "客先担当者", j.get("D", "")),
            ("出荷先", j.get("E", ""), "請求先", j.get("F", "")),
            ("プロジェクト名", j.get("G", ""), "品名、ライン名等", j.get("H", "")),
            ("製品詳細", j.get("I", ""), "数量 / 単位", f"{j.get('J', '')} {j.get('K', '')}".strip()),
            ("希望納期", format_excel_date(j.get("L", "")), "製造又は手配先", j.get("M", "")),
            ("段取済", format_excel_date(j.get("N", "")), "段取者", j.get("O", "")),
            ("切断済", format_excel_date(j.get("P", "")), "製作完", j.get("Q", "")),
            ("出荷日", format_excel_date(j.get("R", "")), "発送", j.get("S", "")),
            ("請求完了日", format_excel_date(j.get("T", "")), "売り上げ区分", j.get("U", "")),
            ("目標工数", j.get("Y", ""), "伝票", j.get("W", "")),
            ("出荷予定日", format_excel_date(j.get("AB", "")), "塗装色", j.get("AC", "")),
            ("備考", j.get("V", ""), "仕分け色", j.get("AD", "")),
        ]
        for i, p in enumerate(pairs):
            self.tree_juchu.insert("", tk.END, values=p, tags=("even" if i % 2 == 0 else "odd",))

        # 5. 制作伝票明細
        self.tree_seizo.delete(*self.tree_seizo.get_children())
        s_list = self.seizo_by_kno.get(kno, [])
        row_cnt = 0
        for srec in s_list:
            src_f = srec.get("_src_file", "")
            upd_t = srec.get("_mtime", "")
            specs = srec.get("_specs", {})
            units = srec.get("_units", {})
            item_defs = [
                (21, "調整用", "", "調整用"),
                (28, "材料費１（型鋼）", "", "材料費１(型鋼)"),
                (29, "材料費２（鉄板）", "", "材料費２(鉄板)"),
                (30, "材料費３（コラム）", "", "材料費３(コラム)"),
                (31, "その他材料費", specs.get(31, "ヘルテン"), "その他材料費(31行)"),
                (32, "その他材料費", specs.get(32, "ｴﾊﾞｰﾊｰﾄﾞ"), "その他材料費(32行)"),
                (33, "その他材料費", specs.get(33, "SUS"), "その他材料費(33行)"),
                (34, "切断材料", "", "切断材料"),
                (35, "材料費計", "", "材料費計"),
                (36, "設計費", "", "設計費"),
                ("AA17", "重機、運搬費", "", "重機、運搬費"),
                ("AA18", "車両、交通費", "", "車両、交通費"),
                ("AA19", "出張費", "", "出張費"),
                ("AA30", "総請求金額", "", "総請求金額"),
            ]
            for r_no, lbl, sp_txt, key_name in item_defs:
                u_val, q_val = units.get(r_no, ("", ""))
                amt = int(round(to_float(srec.get(key_name, 0.0), 0.0)))
                self.tree_seizo.insert(
                    "",
                    tk.END,
                    values=(r_no, lbl, sp_txt, u_val, q_val, f"{amt:,}", upd_t, src_f),
                    tags=("even" if row_cnt % 2 == 0 else "odd",),
                )
                row_cnt += 1
            sh_sub1 = int(round(to_float(srec.get("_シート小計1", 0.0), 0.0)))
            self.tree_seizo.insert(
                "",
                tk.END,
                values=(38, "小計１ (Excelシート記載値)", "", "", "", f"{sh_sub1:,}", upd_t, src_f),
                tags=("even" if row_cnt % 2 == 0 else "odd",),
            )
            row_cnt += 1
        self.detail_notebook.tab(4, text=f"制作伝票明細 ({len(s_list)}件)")
        if hasattr(self, "detail_frame") and self.detail_frame is not None:
            d_kousu = d_total_mins / KOUSU_MINUTES
            d_hours = d_total_mins / 60.0
            summary_line = (
                f"管理No【{kno}】  "
                f"日報: {len(d_list)}件 ({d_hours:,.2f}h / {d_kousu:,.2f}工数 / {d_total_yen:,}円)  |  "
                f"購入品: {len(k_list)}件 ({k_total_yen:,}円)  |  "
                f"外注: {len(g_list)}件 ({g_total_yen:,}円)  |  "
                f"制作伝票: {len(s_list)}件"
            )
            self.detail_frame.configure(
                text=f"選択した管理Noの内訳明細【{kno}】   ※明細行は Ctrl+C / 右クリックでコピー可"
            )
            if hasattr(self, "detail_bar_info_var"):
                self.detail_bar_info_var.set(summary_line)

    def open_code_settings_dialog(self):
        """作業コード範囲設定ダイアログを開く"""
        dlg = tk.Toplevel(self.root)
        dlg.title("作業コード範囲設定")
        dlg.geometry("520x370")
        dlg.minsize(480, 350)
        dlg.resizable(True, True)
        dlg.configure(bg=COLORS["bg_card"])
        dlg.transient(self.root)
        dlg.grab_set()

        frame = ttk.Frame(dlg, style="Card.TFrame", padding=16)
        frame.pack(fill=tk.BOTH, expand=True)

        # 最下部フッターボタンを先に side=BOTTOM で配置し、高DPI環境でも見切れを防止
        btn_f = ttk.Frame(frame, style="Flat.TFrame")
        btn_f.pack(side=tk.BOTTOM, fill=tk.X, pady=(12, 0))

        ttk.Label(
            frame,
            text="日報ファイルから時間を算出する作業コードの範囲を設定します。\n「20-49」や「20～49」、カンマ区切り「20-49, 100-105」の形式で指定できます。",
            style="Card.TLabel",
        ).pack(anchor=tk.W, pady=(0, 12))

        grid_f = ttk.Frame(frame, style="Flat.TFrame")
        grid_f.pack(fill=tk.X, pady=4)

        entry_vars = {}
        for idx, cat in enumerate(WORK_CATEGORIES):
            ttk.Label(grid_f, text=f"{cat} :", style="Card.TLabel", width=16).grid(
                row=idx, column=0, sticky=tk.W, padx=6, pady=6
            )
            v = tk.StringVar(value=self.code_ranges.get(cat, DEFAULT_CODE_RANGES[cat]))
            entry_vars[cat] = v
            ent = ttk.Entry(grid_f, textvariable=v, width=24)
            ent.grid(row=idx, column=1, sticky=tk.W, padx=6, pady=6)
            ttk.Label(
                grid_f,
                text=f"(デフォルト: {DEFAULT_CODE_RANGES[cat].replace('-', '～')})",
                style="Mtime.TLabel",
            ).grid(row=idx, column=2, sticky=tk.W, padx=6, pady=6)

        def restore_defaults():
            for cat, def_val in DEFAULT_CODE_RANGES.items():
                entry_vars[cat].set(def_val)

        def save_and_apply():
            for cat in WORK_CATEGORIES:
                spec = entry_vars[cat].get().strip()
                parsed = parse_code_range_spec(spec)
                if not parsed:
                    messagebox.showwarning(
                        "入力確認",
                        f"「{cat}」の作業コード範囲が正しくありません。\n例: 20-49",
                        parent=dlg,
                    )
                    return
                self.code_ranges[cat] = spec
            self._update_code_summary_label()
            self.save_config()
            dlg.destroy()
            if self.raw_daily or self.raw_kounyu or self.raw_gaichu or self.raw_juchu or self.raw_seizo:
                self._recalculate_from_loaded_data()

        ttk.Button(btn_f, text="デフォルトに戻す", command=restore_defaults).pack(side=tk.LEFT)
        ttk.Button(btn_f, text="キャンセル", command=dlg.destroy).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(btn_f, text="保存して再算出", style="Primary.TButton", command=save_and_apply).pack(side=tk.RIGHT)
        dlg.bind("<Escape>", lambda e: dlg.destroy())

    def open_rate_settings_dialog(self):
        """既定表示単位（金額/時間）および日報の金額換算レート（円/分）設定ダイアログを開く"""
        def on_apply(new_rates, new_default_unit, switch_now):
            self.work_rates = dict(new_rates)
            if new_default_unit in VALID_TIME_UNITS:
                self.default_time_unit = new_default_unit
            if switch_now:
                self.time_unit_var.set(self.default_time_unit)
            self._update_unit_button_styles()
            self.save_config()
            self.apply_filter_and_refresh()
            if self._current_detail_kno:
                self._show_detail_for_kno(self._current_detail_kno)
            self.status_var.set("表示単位・単価レート設定を保存し、一覧を再計算しました。")

        WorkRateSettingsDialog(
            self.root,
            self.work_rates,
            self.default_time_unit,
            self.time_unit_var.get(),
            on_apply,
        )

    def _prepare_export_data(self, include_col_groups: bool = False):
        """
        現在表示中の列順番・アクティブペインの並び順・フィルター結果からExcel出力データを生成する。
        小計1・小計2の内訳列はExcelの列グループ化（デフォルト折りたたみ）用として常に含め、
        グループ範囲（1始まりの列番号タプル一覧）も併せて算出する。
        """
        unit = self.time_unit_var.get()
        export_visible = set(self.visible_set) | set(SUBTOTAL1_DETAIL_COLUMNS) | set(SUBTOTAL2_DETAIL_COLUMNS)
        src_cols = [c for c in self.column_order if c in export_visible]
        headers = [get_display_column_label(c, unit) for c in src_cols]
        numeric_indices = {i for i, col in enumerate(src_cols) if col in NUMERIC_COLUMNS}

        # 小計1内訳・小計2内訳の連続列区間をそれぞれ1始まりのグループ範囲として抽出
        col_groups = []
        for group_col_set in (set(SUBTOTAL1_DETAIL_COLUMNS), set(SUBTOTAL2_DETAIL_COLUMNS)):
            run_start = None
            run_end = None
            for idx_1, col in enumerate(src_cols, 1):
                if col in group_col_set:
                    if run_start is None:
                        run_start = idx_1
                    run_end = idx_1
                else:
                    if run_start is not None:
                        col_groups.append((run_start, run_end))
                        run_start = None
                        run_end = None
            if run_start is not None:
                col_groups.append((run_start, run_end))

        pane = self._get_active_pane()
        export_source = pane.rows if pane.rows else self.displayed_rows
        rows = []
        is_yen_mode = (unit == "金額(円)")
        totals = {c: 0.0 for c in CALC_COLUMNS}
        rates_tuple = self._get_rates_tuple()
        get_w = self._get_work_numeric_val

        for r in export_source:
            if is_yen_mode:
                row_sub_yen = 0
                for c in WORK_CATEGORIES:
                    yen_v = int(round(get_w(r, c, rates_tuple=rates_tuple, unit=unit)))
                    totals[c] += yen_v
                    row_sub_yen += yen_v
                totals["作業小計"] += row_sub_yen
            else:
                for c in WORK_CATEGORIES:
                    totals[c] += r[c]
                totals["作業小計"] += r["作業小計"]
            for c in COST_YEN_COLUMNS:
                totals[c] += r.get(c, 0.0)
            for c in DERIVED_YEN_COLUMNS:
                totals[c] += get_w(r, c, rates_tuple=rates_tuple, unit=unit)

            row_vals = []
            for col in src_cols:
                if col in RATE_DEPENDENT_COLUMNS:
                    row_vals.append(get_w(r, col, rates_tuple=rates_tuple, unit=unit))
                elif col in COST_YEN_COLUMNS:
                    row_vals.append(int(round(r.get(col, 0.0))))
                else:
                    row_vals.append(r.get(col, ""))
            rows.append(row_vals)

        totals_row = []
        for idx, col in enumerate(src_cols):
            if idx == 0:
                totals_row.append(f"合計 ({len(self.displayed_rows)}件)")
            elif col in ("合計/数量", "総請求金額/数量"):
                totals_row.append("")
            elif col in DERIVED_YEN_COLUMNS:
                totals_row.append(int(round(totals[col])))
            elif col in WORK_CATEGORIES or col == "作業小計":
                if is_yen_mode:
                    totals_row.append(int(round(totals[col])))
                elif unit == "時間(分÷60)":
                    totals_row.append(round(totals[col] / 60.0, 2))
                elif unit == "工数(8h=1)":
                    totals_row.append(round(totals[col] / KOUSU_MINUTES, 2))
                else:
                    totals_row.append(totals[col])
            elif col in COST_YEN_COLUMNS:
                totals_row.append(int(round(totals[col])))
            else:
                totals_row.append("")

        if include_col_groups:
            return headers, rows, numeric_indices, totals_row, col_groups
        return headers, rows, numeric_indices, totals_row

    def open_excel_new_workbook_action(self):
        """表示中の一覧データを保存ファイルではなくExcelの新規ブック（未保存ブック）として開く"""
        if not self.displayed_rows:
            messagebox.showinfo("出力対象なし", "Excelへ出力するデータがありません。先に集計を実行してください。", parent=self.root)
            return

        headers, rows, numeric_indices, totals_row, col_groups = self._prepare_export_data(include_col_groups=True)
        protected_snapshot = self.get_protected_source_paths()
        protected_dirs_snapshot = self.get_protected_source_dirs()
        self.status_var.set("Excelを起動して新規ブックにデータを書き込んでいます...")
        try:
            self.btn_excel_open.configure(state=tk.DISABLED, text="Excel起動中...")
        except Exception:
            pass

        def worker():
            err_msg = None
            fallback_path = None
            try:
                fallback_path = open_in_excel_new_workbook(
                    headers,
                    rows,
                    numeric_indices,
                    totals_row=totals_row,
                    protected_paths=protected_snapshot,
                    col_groups=col_groups,
                    protected_dirs=protected_dirs_snapshot,
                )
            except Exception as e:
                err_msg = str(e)

            def finish():
                try:
                    self.btn_excel_open.configure(state=tk.NORMAL, text="Excel新規ブックで開く")
                except Exception:
                    pass
                if self._closing:
                    return
                if err_msg:
                    self.status_var.set("Excel新規ブックの起動に失敗しました。")
                    messagebox.showerror("Excel起動エラー", f"Excel新規ブックの作成中にエラーが発生しました:\n{err_msg}", parent=self.root)
                elif fallback_path:
                    self.status_var.set(f"Excel（一時ファイル）で表示しました（{len(rows):,}件）")
                else:
                    self.status_var.set(f"Excelの新規ブックを開きました（{len(rows):,}件・未保存ブック）")

            if not self._closing:
                self._ui_queue.put(finish)

        threading.Thread(target=worker, daemon=True).start()

    def _prepare_detail_sheets_spec(self, kno: str, only_active_tab: bool = False):
        """
        選択中の管理Noの内訳明細5タブ（日報・購入品・外注加工・受注表詳細・制作伝票）から
        印刷レイアウト対応のExcelシート仕様リストを構築する。
        """
        j = self.juchu_by_kno.get(kno, {})
        # 案件サブタイトル（受注先・出荷先・品名・製品詳細）
        agg_row = next((r for r in self.aggregated_rows if r.get("管理No") == kno), None)
        v_juchu = (agg_row.get("受注先", "") if agg_row else j.get("C", "")) or "-"
        v_shukka = (agg_row.get("出荷先", "") if agg_row else j.get("E", "")) or "-"
        v_hinmei = (agg_row.get("品名、ライン名等", "") if agg_row else j.get("H", "")) or "-"
        v_detail = (agg_row.get("製品詳細", "") if agg_row else j.get("I", "")) or "-"
        subtitle = f"受注先: {v_juchu}  /  出荷先: {v_shukka}  /  品名: {v_hinmei}  /  製品詳細: {v_detail}"

        # 1. 日報データ明細
        d_list = self.daily_by_kno.get(kno, [])
        d_rows = []
        d_sum_mins = 0.0
        d_sum_yen = 0
        for d in d_list:
            mins = to_float(d.get("合計", 0.0), 0.0)
            d_sum_mins += mins
            upd = d.get("更新日", "")
            if "." in upd:
                upd = upd.split(".")[0]
            cat = d.get("分類区分", "")
            if cat in self.work_rates:
                rate = to_float(self.work_rates.get(cat, DEFAULT_WORK_RATES.get(cat, 0.0)), 0.0)
                yen_v = int(round(mins * rate))
                d_sum_yen += yen_v
            else:
                yen_v = ""
            d_rows.append([
                upd,
                d.get("作業者", ""),
                d.get("作業コード", ""),
                d.get("作業名", ""),
                cat,
                int(round(mins)) if abs(mins - round(mins)) < 1e-6 else round(mins, 2),
                round(mins / 60.0, 2),
                round(mins / KOUSU_MINUTES, 2),
                yen_v,
                d.get("出荷先", ""),
                d.get("製品詳細", ""),
                d.get("ファイル名", ""),
            ])
        d_totals = (
            [
                f"合計 ({len(d_list)}件)", "", "", "", "",
                int(round(d_sum_mins)) if abs(d_sum_mins - round(d_sum_mins)) < 1e-6 else round(d_sum_mins, 2),
                round(d_sum_mins / 60.0, 2),
                round(d_sum_mins / KOUSU_MINUTES, 2),
                d_sum_yen, "", "", "",
            ]
            if d_rows
            else None
        )
        spec_daily = {
            "name": "日報データ明細",
            "kind": "daily",
            "title": f"管理No: {kno}  日報データ明細（{len(d_list)}件）",
            "subtitle": subtitle,
            "print_header": f"管理No: {kno} 日報データ明細（DenpyoHub_Miya）",
            "headers": ["更新日", "作業者", "作業コード", "作業名", "分類区分", "時間(分)", "時間(h)", "工数(h/8)", "金額(円)", "出荷先", "製品詳細", "元ファイル"],
            "rows": d_rows,
            "numeric_col_indices": {5, 6, 7, 8},
            "decimal_col_indices": {6, 7},
            "wrap_col_indices": {10, 11},
            "totals_row": d_totals,
        }

        # 2 & 3. 購入品明細 / 外注加工明細
        def build_kg_spec(tab_name: str, kind: str, src_list: list) -> dict:
            rows_kg = []
            sum_qty = 0.0
            sum_amt = 0
            for r in src_list:
                m_raw = r.get("M", "")
                n_raw = r.get("N", "")
                q_num = _calc_kounyu_row_amount(r)
                n_num = to_float(n_raw, 0.0) if str(n_raw).strip() != "" else 0.0
                sum_qty += n_num
                sum_amt += int(round(q_num))
                rows_kg.append([
                    r.get("F", ""),
                    r.get("G", ""),
                    r.get("H", ""),
                    r.get("I", ""),
                    r.get("K", ""),
                    r.get("L", ""),
                    int(round(to_float(m_raw, 0.0))) if str(m_raw).strip() != "" else "",
                    n_num if str(n_raw).strip() != "" else "",
                    int(round(q_num)),
                    format_excel_date(r.get("U", "")),
                    format_excel_date(r.get("W", "")),
                    r.get("T", ""),
                    r.get("_src_file", ""),
                ])
            tot_kg = (
                [
                    f"合計 ({len(src_list)}件)", "", "", "", "", "", "",
                    int(round(sum_qty)) if abs(sum_qty - round(sum_qty)) < 1e-6 else round(sum_qty, 2),
                    sum_amt, "", "", "", "",
                ]
                if rows_kg
                else None
            )
            return {
                "name": tab_name,
                "kind": kind,
                "title": f"管理No: {kno}  {tab_name}（{len(src_list)}件）",
                "subtitle": subtitle,
                "print_header": f"管理No: {kno} {tab_name}（DenpyoHub_Miya）",
                "headers": ["図面番号", "品番", "品名", "メーカー", "仕様", "仕入先", "単価", "必要数", "金額(単価X必要数)", "注文日", "請求日", "備考", "元ファイル"],
                "rows": rows_kg,
                "numeric_col_indices": {6, 7, 8},
                "decimal_col_indices": set(),
                "wrap_col_indices": {2, 4, 11},
                "totals_row": tot_kg,
            }

        spec_kounyu = build_kg_spec("購入品明細", "kounyu", self.kounyu_by_kno.get(kno, []))
        spec_gaichu = build_kg_spec("外注加工明細", "gaichu", self.gaichu_by_kno.get(kno, []))

        # 4. 受注表詳細情報
        juchu_pairs = [
            ["管理No", kno, "受注日", format_excel_date(j.get("A", ""))],
            ["受注先", j.get("C", ""), "客先担当者", j.get("D", "")],
            ["出荷先", j.get("E", ""), "請求先", j.get("F", "")],
            ["プロジェクト名", j.get("G", ""), "品名、ライン名等", j.get("H", "")],
            ["製品詳細", j.get("I", ""), "数量 / 単位", f"{j.get('J', '')} {j.get('K', '')}".strip()],
            ["希望納期", format_excel_date(j.get("L", "")), "製造又は手配先", j.get("M", "")],
            ["段取済", format_excel_date(j.get("N", "")), "段取者", j.get("O", "")],
            ["切断済", format_excel_date(j.get("P", "")), "製作完", j.get("Q", "")],
            ["出荷日", format_excel_date(j.get("R", "")), "発送", j.get("S", "")],
            ["請求完了日", format_excel_date(j.get("T", "")), "売り上げ区分", j.get("U", "")],
            ["目標工数", j.get("Y", ""), "伝票", j.get("W", "")],
            ["出荷予定日", format_excel_date(j.get("AB", "")), "塗装色", j.get("AC", "")],
            ["備考", j.get("V", ""), "仕分け色", j.get("AD", "")],
        ]
        spec_juchu = {
            "name": "受注表詳細情報",
            "kind": "juchu",
            "title": f"管理No: {kno}  受注表詳細情報",
            "subtitle": subtitle,
            "print_header": f"管理No: {kno} 受注表詳細情報（DenpyoHub_Miya）",
            "headers": ["項目名", "内容", "項目名2", "内容2"],
            "rows": juchu_pairs,
            "numeric_col_indices": set(),
            "decimal_col_indices": set(),
            "wrap_col_indices": {1, 3},
            "key_col_indices": {0, 2},
            "totals_row": None,
        }

        # 5. 制作伝票明細
        s_list = self.seizo_by_kno.get(kno, [])
        s_rows = []
        for srec in s_list:
            src_f = srec.get("_src_file", "")
            upd_t = srec.get("_mtime", "")
            specs = srec.get("_specs", {})
            units = srec.get("_units", {})
            item_defs = [
                (21, "調整用", "", "調整用"),
                (28, "材料費１（型鋼）", "", "材料費１(型鋼)"),
                (29, "材料費２（鉄板）", "", "材料費２(鉄板)"),
                (30, "材料費３（コラム）", "", "材料費３(コラム)"),
                (31, "その他材料費", specs.get(31, "ヘルテン"), "その他材料費(31行)"),
                (32, "その他材料費", specs.get(32, "ｴﾊﾞｰﾊｰﾄﾞ"), "その他材料費(32行)"),
                (33, "その他材料費", specs.get(33, "SUS"), "その他材料費(33行)"),
                (34, "切断材料", "", "切断材料"),
                (35, "材料費計", "", "材料費計"),
                (36, "設計費", "", "設計費"),
                ("AA17", "重機、運搬費", "", "重機、運搬費"),
                ("AA18", "車両、交通費", "", "車両、交通費"),
                ("AA19", "出張費", "", "出張費"),
                ("AA30", "総請求金額", "", "総請求金額"),
            ]
            for r_no, lbl, sp_txt, key_name in item_defs:
                u_val, q_val = units.get(r_no, ("", ""))
                amt = int(round(to_float(srec.get(key_name, 0.0), 0.0)))
                s_rows.append([str(r_no), lbl, sp_txt, str(u_val), str(q_val), amt, upd_t, src_f])
            sh_sub1 = int(round(to_float(srec.get("_シート小計1", 0.0), 0.0)))
            s_rows.append(["38", "小計１ (Excelシート記載値)", "", "", "", sh_sub1, upd_t, src_f])

        spec_seizo = {
            "name": "制作伝票明細",
            "kind": "seizo",
            "title": f"管理No: {kno}  制作伝票明細（{len(s_list)}件）",
            "subtitle": subtitle,
            "print_header": f"管理No: {kno} 制作伝票明細（DenpyoHub_Miya）",
            "headers": ["行", "項目名", "材質・仕様", "単価 / レート", "時間 / 重量", "金額(円)", "更新日時", "元ファイル"],
            "rows": s_rows,
            "numeric_col_indices": {5},
            "decimal_col_indices": set(),
            "wrap_col_indices": {2, 7},
            "totals_row": None,
        }

        all_specs = [spec_daily, spec_kounyu, spec_gaichu, spec_juchu, spec_seizo]
        try:
            cur_tab_idx = self.detail_notebook.index(self.detail_notebook.select())
        except Exception:
            cur_tab_idx = 0
        if not (0 <= cur_tab_idx < len(all_specs)):
            cur_tab_idx = 0
        active_name = all_specs[cur_tab_idx]["name"]

        if only_active_tab:
            return [all_specs[cur_tab_idx]], active_name
        return all_specs, active_name

    def open_detail_excel_action(self, only_active_tab: bool = False):
        """選択中の管理Noの内訳明細（表示中タブのみ、または全5タブ）を印刷設定済みのExcel新規ブックで開く"""
        kno = self._current_detail_kno
        if not kno:
            messagebox.showinfo(
                "管理No未選択",
                "内訳明細をExcelで開くには、まず一覧から対象の管理Noの行をクリックして選択してください。",
                parent=self.root,
            )
            return

        sheets_spec, active_name = self._prepare_detail_sheets_spec(kno, only_active_tab=only_active_tab)
        protected_snapshot = self.get_protected_source_paths()
        protected_dirs_snapshot = self.get_protected_source_dirs()
        mode_lbl = f"「{active_name}」" if only_active_tab else "全5タブ"
        self.status_var.set(f"管理No【{kno}】の内訳明細（{mode_lbl}）をExcel新規ブックで開いています...")
        try:
            self.btn_detail_excel_tab.configure(state=tk.DISABLED)
            self.btn_detail_excel_all.configure(state=tk.DISABLED, text="Excel起動中...")
        except Exception:
            pass

        def worker():
            err_msg = None
            fallback_path = None
            try:
                fallback_path = open_detail_in_excel_new_workbook(
                    kno,
                    sheets_spec,
                    active_sheet_name=active_name,
                    protected_paths=protected_snapshot,
                    protected_dirs=protected_dirs_snapshot,
                )
            except Exception as e:
                err_msg = str(e)

            def finish():
                try:
                    self.btn_detail_excel_tab.configure(state=tk.NORMAL)
                    self.btn_detail_excel_all.configure(state=tk.NORMAL, text="内訳明細をExcelで開く（全タブ）")
                except Exception:
                    pass
                if self._closing:
                    return
                if err_msg:
                    self.status_var.set("内訳明細のExcel起動に失敗しました。")
                    messagebox.showerror("Excel起動エラー", f"内訳明細のExcel作成中にエラーが発生しました:\n{err_msg}", parent=self.root)
                elif fallback_path:
                    self.status_var.set(f"管理No【{kno}】の内訳明細（{mode_lbl}）をExcel（一時ファイル）で開きました。")
                else:
                    self.status_var.set(f"管理No【{kno}】の内訳明細（{mode_lbl}）をExcelの新規ブックで開きました（印刷レイアウト設定済み）。")

            if not self._closing:
                self._ui_queue.put(finish)

        threading.Thread(target=worker, daemon=True).start()

    def open_help_dialog(self):
        """見やすいガイド形式のヘルプ画面を表示する"""
        HelpDialog(self.root)

    def _on_close(self):
        self._closing = True
        self._active_job_id += 1
        self.meta_pane.cancel_pending_jobs()
        self.calc_pane.cancel_pending_jobs()
        for attr_name in (
            "_progress_hide_job",
            "_poll_after_id",
            "_max_after_id",
            "_startup_after_id",
            "_sash_after_id",
            "_restore_y_after_id",
            "_search_debounce_job",
        ):
            self._cancel_after(attr_name)
        self.save_config()
        try:
            self.root.quit()
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass


# 互換エイリアス
StockMatchApp = DenpyoHubApp


def main():
    root = tk.Tk()
    app = DenpyoHubApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
