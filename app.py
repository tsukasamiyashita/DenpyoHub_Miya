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
import xml.etree.ElementTree as ET
from collections import defaultdict, Counter
from concurrent.futures import ThreadPoolExecutor
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

# デフォルトの作業コード範囲設定
DEFAULT_CODE_RANGES = {
    "作業時間": "20-49",
    "移動時間": "90-99",
    "修理改造時間": "50-69",
    "設計関係": "1-19",
    "その他作業": "70-89",
}

# 表示単位の定義（デフォルトは金額(円)）
VALID_TIME_UNITS = ("金額(円)", "分(そのまま)", "時間(分÷60)")
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

# 制作伝票（個別Excel）から追加抽出する項目列
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
]
SEIZO_COST_COLUMNS = tuple(c for c in SEIZO_COLUMNS if c != "小計1")
COST_YEN_COLUMNS = {"購入品", "外注加工", "購入品小計"} | set(SEIZO_COST_COLUMNS)

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

# 利用可能な全列（デフォルト表示順は BASE_COLUMNS + CALC_COLUMNS、オプションで「備考」も追加可能）
DEFAULT_COLUMN_ORDER = BASE_COLUMNS + CALC_COLUMNS
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
    s = str(val).strip()
    if not s:
        return ""
    zen = "０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ"
    han = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    s = s.translate(str.maketrans(zen, han))
    if s.endswith(".0"):
        prefix = s[:-2]
        if prefix.lstrip("-").isdigit():
            s = prefix
    return s


def to_float(val, default: float = 0.0) -> float:
    """安全にfloatへ変換する"""
    if val is None:
        return default
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(",", "")
    if not s or s.startswith("#"):
        return default
    try:
        return float(s)
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
    latest_ts = max(os.path.getmtime(p) for p in existing)
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
    s = str(spec_str).strip()
    zen = "０１２３４５６７８９～〜－ー，、"
    han = "0123456789~~--,,"
    s = s.translate(str.maketrans(zen, han))
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


def open_readonly_bytes(src_path: str, cancel_check=None) -> bytes:
    """
    参照元ファイルを一切変更・ロックせず、読み取り専用("rb")でメモリへ高速に読み込む。
    一時ファイルのディスク書き出しを発生させず、元のファイルを厳重に保護する。
    """
    if cancel_check and cancel_check():
        raise InterruptedError("処理がキャンセルされました")
    if not src_path or not os.path.exists(src_path):
        raise FileNotFoundError(f"ファイルが見つかりません: {src_path}")
    if os.path.isdir(src_path):
        raise IsADirectoryError(f"ファイルではなくフォルダが指定されています: {src_path}")

    chunks = []
    with open(src_path, "rb") as f_src:
        while True:
            if cancel_check and cancel_check():
                raise InterruptedError("処理がキャンセルされました")
            buf = f_src.read(1024 * 1024)
            if not buf:
                break
            chunks.append(buf)
    return b"".join(chunks)


def _copy_file_chunked(src_path: str, dst_path: str, cancel_check=None):
    """メタデータコピーによる遅延を避けつつ1MB単位で安全に一時ファイルへコピーする（参照元上書き厳重防止付き）"""
    norm_src = os.path.normcase(os.path.abspath(os.path.normpath(src_path)))
    norm_dst = os.path.normcase(os.path.abspath(os.path.normpath(dst_path)))
    if norm_src == norm_dst:
        raise PermissionError(f"参照元ファイルへの同一パス書き込みをブロックしました: {src_path}")
    raw = open_readonly_bytes(src_path, cancel_check=cancel_check)
    with open(dst_path, "wb") as f_dst:
        f_dst.write(raw)
    try:
        os.chmod(dst_path, stat.S_IREAD)
    except Exception:
        pass


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


def read_excel_sheet_rows(filepath: str, preferred_sheets=None, cancel_check=None, return_meta: bool = False):
    """
    XLSM/XLSXファイルから指定シート（無ければ先頭シート）の全行を辞書リストとしてインメモリで高速に読み込む。
    各行は {'A': val, 'B': val, ...} の辞書。
    参照元ファイルは読み取り専用バイト列としてメモリに読み込むため、元ファイルは一切変更・ロックされない。
    """
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
        z_ctx = zipfile.ZipFile(io.BytesIO(raw_bytes), "r")
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

    with z_ctx as z:
        namelist_set = set(z.namelist())
        if "xl/workbook.xml" not in namelist_set or "xl/_rels/workbook.xml.rels" not in namelist_set:
            raise FormatValidationError(
                "Excelファイル",
                filepath,
                "Excelブック構造（xl/workbook.xml）が見つかりません。",
                "標準のExcelブック (.xlsm / .xlsx)",
                "workbook.xml 不在のZIPファイル",
            )

        shared_strings = _load_shared_strings_from_zip(z, namelist_set, cancel_check=cancel_check)
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
            return ([], "", all_sheet_names) if return_meta else []

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
                    row_vals = {}
                    for c_el in elem.findall(c_tag):
                        ref = c_el.attrib.get("r", "")
                        col = ref.rstrip("0123456789")
                        if not col:
                            continue
                        t_attr = c_el.attrib.get("t", "")
                        v_el = c_el.find(v_tag)
                        val = v_el.text if v_el is not None else None
                        if t_attr == "s" and val is not None:
                            try:
                                idx = int(val)
                                if 0 <= idx < ss_len:
                                    val = shared_strings[idx]
                            except ValueError:
                                pass
                        elif t_attr == "inlineStr":
                            is_el = c_el.find(is_tag)
                            if is_el is not None:
                                val = "".join(x.text or "" for x in is_el.iter(t_tag))
                        elif t_attr == "e":
                            val = None
                        if val is not None:
                            s_val = str(val).strip()
                            if s_val != "":
                                row_vals[col] = s_val
                    if row_vals:
                        rows.append(row_vals)
                    elem.clear()
                    row_count += 1
                    if sheet_data_el is not None and (row_count & 511) == 0:
                        sheet_data_el.clear()
        if return_meta:
            return rows, matched_sheet_name, all_sheet_names
        return rows


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


def expand_seizo_excel_files(paths: list) -> list:
    """
    制作伝票用のフォルダ・ファイル指定リストから、サブフォルダを含むすべての
    .xls / .xlsx / .xlsm ファイル一覧を再帰的に展開する。
    一時ロックファイル（~$...）や隠しシステムフォルダは除外し、重複を排除する。
    """
    result = []
    seen = set()
    exts = (".xls", ".xlsx", ".xlsm")

    def add_if_excel(fpath: str):
        name = os.path.basename(fpath)
        if name.startswith("~$") or not os.path.isfile(fpath):
            return
        if name.lower().endswith(exts):
            norm = os.path.normpath(fpath)
            key = os.path.normcase(os.path.abspath(norm))
            if key not in seen:
                seen.add(key)
                result.append(norm)

    for p in paths:
        if not p:
            continue
        if os.path.isdir(p):
            try:
                for root_dir, dirs, files in os.walk(p):
                    dirs[:] = sorted(d for d in dirs if not d.startswith("."))
                    for fname in sorted(files):
                        add_if_excel(os.path.join(root_dir, fname))
            except Exception:
                pass
        elif os.path.isfile(p):
            add_if_excel(p)
    return result


def expand_seizo_protected_dirs(paths: list) -> set:
    """制作伝票フォルダおよびその全サブフォルダの正規化絶対パス集合（書き込み禁止対象）を返す"""
    protected_dirs = set()
    for p in (paths or []):
        if not p:
            continue
        norm_p = os.path.normcase(os.path.abspath(os.path.normpath(p)))
        if os.path.isdir(p):
            protected_dirs.add(norm_p)
            try:
                for root_dir, dirs, _ in os.walk(p):
                    protected_dirs.add(os.path.normcase(os.path.abspath(os.path.normpath(root_dir))))
                    for d in dirs:
                        protected_dirs.add(os.path.normcase(os.path.abspath(os.path.normpath(os.path.join(root_dir, d)))))
            except Exception:
                pass
    return protected_dirs


def summarize_seizo_folder_mtime(folder_paths: list) -> str:
    """制作伝票フォルダ（サブフォルダ含む）内の対象Excelファイル数と最新更新日時を返す"""
    if not folder_paths:
        return "未選択"
    missing_folders = [p for p in folder_paths if p and not os.path.exists(p)]
    excel_files = expand_seizo_excel_files(folder_paths)
    if not excel_files:
        if missing_folders:
            return "フォルダが見つかりません"
        return "フォルダ・サブフォルダ内に対象Excelなし"
    latest_ts = max(os.path.getmtime(p) for p in excel_files)
    latest_str = datetime.datetime.fromtimestamp(latest_ts).strftime("%Y/%m/%d %H:%M:%S")
    return f"最新更新: {latest_str} (Excel {len(excel_files):,}ファイル・サブフォルダ含)"


def _read_seizo_cells_xml(filepath: str, raw_bytes: bytes, cancel_check=None) -> tuple:
    """XLSX/XLSMのメモリバイト列から「製作伝票」シートの1～40行のセル {(row_num, col): val} とシート情報を読み込む"""
    cells = {}
    try:
        z_ctx = zipfile.ZipFile(io.BytesIO(raw_bytes), "r")
    except zipfile.BadZipFile:
        raise FormatValidationError(
            "制作伝票ファイル",
            filepath,
            "有効なExcel (.xlsx / .xlsm) ファイルとして開けません（ファイル破損または非対応形式）。",
            "制作伝票Excel (.xls / .xlsx / .xlsm)",
            "ZIP構造不正",
        )

    with z_ctx as z:
        namelist_set = set(z.namelist())
        if "xl/workbook.xml" not in namelist_set or "xl/_rels/workbook.xml.rels" not in namelist_set:
            raise FormatValidationError(
                "制作伝票ファイル",
                filepath,
                "Excelブック構造（xl/workbook.xml）が見つかりません。",
                "制作伝票Excel (.xls / .xlsx / .xlsm)",
                "workbook.xml 不在",
            )

        shared_strings = _load_shared_strings_from_zip(z, namelist_set, cancel_check=cancel_check)
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
        ss_len = len(shared_strings)

        with z.open(sheet_target) as wf:
            for event, elem in ET.iterparse(wf, events=("end",)):
                if event == "end" and elem.tag == row_tag:
                    r_attr = elem.attrib.get("r", "")
                    try:
                        r_num = int(r_attr)
                    except ValueError:
                        r_num = 0
                    if r_num > 42:
                        elem.clear()
                        break
                    if 1 <= r_num <= 40:
                        for c_el in elem.findall(c_tag):
                            ref = c_el.attrib.get("r", "")
                            col = ref.rstrip("0123456789")
                            if not col:
                                continue
                            t_attr = c_el.attrib.get("t", "")
                            v_el = c_el.find(v_tag)
                            val = v_el.text if v_el is not None else None
                            if t_attr == "s" and val is not None:
                                try:
                                    idx = int(val)
                                    if 0 <= idx < ss_len:
                                        val = shared_strings[idx]
                                except ValueError:
                                    pass
                            elif t_attr == "inlineStr":
                                is_el = c_el.find(is_tag)
                                if is_el is not None:
                                    val = "".join(x.text or "" for x in is_el.iter(t_tag))
                            elif t_attr == "e":
                                val = None
                            if val is not None:
                                s_val = str(val).strip()
                                if s_val != "":
                                    cells[(r_num, col)] = s_val
                    elem.clear()
    return cells, matched_sheet_name, all_sheet_names


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


def read_seizo_denpyo_sheet(filepath: str, cancel_check=None, base_folders: list = None) -> dict:
    """
    制作伝票（.xls / .xlsx / .xlsm）から管理Noおよび指定行の項目を抽出して返す。
    必ず読み取り専用バイト列としてメモリ上で解析し、参照元ファイルを一切変更しない。
    サンプル（22601013.xlsm / 22601003.xlsm）とフォーマットが異なる場合は FormatValidationError を送出する。
    """
    raw_bytes = open_readonly_bytes(filepath, cancel_check=cancel_check)
    if not raw_bytes:
        raise FormatValidationError(
            "制作伝票ファイル",
            filepath,
            "ファイルサイズが0バイト（空ファイル）です。",
            "シート「製作伝票」を含む制作伝票Excelファイル",
            "空ファイル (0 bytes)",
        )

    lower = filepath.lower()
    if lower.endswith(".xls"):
        cells, matched_sheet, _ = _read_seizo_cells_xls(filepath, raw_bytes)
    else:
        cells, matched_sheet, _ = _read_seizo_cells_xml(filepath, raw_bytes, cancel_check=cancel_check)

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
        "_mtime": format_file_mtime(filepath),
    }


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
    if "小計" in h_text or "材料費計" in h_text:
        return "subtotal"
    if h_text in WORK_CATEGORIES or any(
        k in h_text
        for k in ("作業費", "移動費", "修理改造費", "設計費(関係)", "その他作業費", "作業時間", "移動時間", "修理改造時間", "設計関係", "その他作業")
    ):
        return "work"
    if h_text in ("購入品", "外注加工") or any(
        k in h_text for k in ("購入品費", "外注加工費", "調整用", "材料費", "切断材料", "設計費")
    ):
        return "cost"
    return "base"


def export_to_xlsx(filepath: str, headers: list, rows: list, numeric_col_indices: set, totals_row: list = None, protected_paths: set = None):
    """
    集計結果をExcel(.xlsx)に出力する。
    openpyxlがインストールされていればスタイル・オートフィルタ・合計行付きで出力し、
    未インストール環境でも標準ライブラリのみで正しい.xlsxを生成する。
    参照元ファイルへの上書きは厳重にブロックする。
    """
    if is_same_or_protected_path(filepath, protected_paths or set()):
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
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 11), 44)

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

    sheet_lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
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
    sheet_lines.append("</sheetData></worksheet>")

    with zipfile.ZipFile(filepath, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.relationships+xml"/>'
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
            "</workbook>",
        )
        zf.writestr("xl/worksheets/sheet1.xml", "".join(sheet_lines))


def open_in_excel_new_workbook(
    headers: list,
    rows: list,
    numeric_col_indices: set,
    totals_row: list = None,
    protected_paths: set = None,
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

                # 2次元配列で一括高速書き込み
                ws.Range(ws.Cells(1, 1), ws.Cells(n_rows, n_cols)).Value = tuple(tuple(r) for r in all_rows)

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

                # 罫線・オートフィルタ・枠固定・列幅自動調整
                full_rng = ws.Range(ws.Cells(1, 1), ws.Cells(n_rows, n_cols))
                full_rng.Borders.LineStyle = 1  # xlContinuous
                full_rng.Borders.Color = rgb_to_bgr(203, 213, 225)
                last_data_r = n_rows - 1 if totals_row else n_rows
                if last_data_r >= 1:
                    ws.Range(ws.Cells(1, 1), ws.Cells(last_data_r, n_cols)).AutoFilter()

                try:
                    ws.Columns.AutoFit()
                    for c_idx in range(1, n_cols + 1):
                        cur_w = ws.Columns(c_idx).ColumnWidth
                        ws.Columns(c_idx).ColumnWidth = min(max(cur_w + 2, 10), 48)
                except Exception:
                    pass

                try:
                    excel.ActiveWindow.SplitRow = 1
                    excel.ActiveWindow.FreezePanes = True
                except Exception:
                    pass
                return
            finally:
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

        ttk.Label(
            frame,
            text=f"{title_label}：複数ファイルを追加・管理できます（集計時に全ファイルが統合・合算されます）",
            style="Title.TLabel",
        ).pack(anchor=tk.W, pady=(0, 8))

        tv_frame = ttk.Frame(frame, style="Card.TFrame")
        tv_frame.pack(fill=tk.BOTH, expand=True)

        cols = ("ファイル名", "更新日時", "状態", "フルパス")
        self.tv = ttk.Treeview(tv_frame, columns=cols, show="headings", selectmode="extended")
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

        btn_bar = ttk.Frame(frame, style="Card.TFrame")
        btn_bar.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(btn_bar, text="ファイルを追加(複数選択可)...", style="Primary.TButton", command=self._add_files).pack(side=tk.LEFT)
        ttk.Button(btn_bar, text="選択行を削除", command=self._remove_selected).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_bar, text="すべてクリア", command=self._clear_all).pack(side=tk.LEFT)

        ttk.Button(btn_bar, text="キャンセル", command=self.destroy).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(btn_bar, text="OK（反映）", style="Success.TButton", command=self._ok).pack(side=tk.RIGHT)

        self._refresh_list()

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


class ColumnOrderDialog(tk.Toplevel):
    """一覧の列の順番・表示設定ダイアログ"""

    def __init__(self, parent, current_order: list, visible_set: set, on_apply):
        super().__init__(parent)
        self.title("列の順番・表示設定")
        self.geometry("500x560")
        self.minsize(460, 480)
        self.configure(bg=COLORS["bg_card"])
        self.transient(parent)
        self.grab_set()

        self.on_apply = on_apply

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

        ttk.Label(frame, text="一覧に表示する列の順番と表示/非表示を設定します", style="Title.TLabel").pack(anchor=tk.W)
        ttk.Label(
            frame,
            text="列を選択して「上へ」「下へ」で並び替え、ダブルクリックまたは「表示切替」でON/OFFできます。",
            style="Mtime.TLabel",
        ).pack(anchor=tk.W, pady=(2, 10))

        body = ttk.Frame(frame, style="Card.TFrame")
        body.pack(fill=tk.BOTH, expand=True)

        tv_frame = ttk.Frame(body, style="Card.TFrame")
        tv_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.tv = ttk.Treeview(tv_frame, columns=("状態", "列名", "種別"), show="headings", selectmode="browse")
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

        footer = ttk.Frame(frame, style="Card.TFrame")
        footer.pack(fill=tk.X, pady=(12, 0))
        ttk.Button(footer, text="キャンセル", command=self.destroy).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(footer, text="保存して適用", style="Primary.TButton", command=self._ok).pack(side=tk.RIGHT)

        self._refresh(select_idx=0)

    def _refresh(self, select_idx: int = None):
        self.tv.delete(*self.tv.get_children())
        for i, (col, vis) in enumerate(self.col_items):
            kind = "集計・小計" if col in CALC_COLUMNS else "基本情報"
            mark = "☑ 表示" if vis else "☐ 非表示"
            self.tv.insert("", tk.END, iid=str(i), values=(mark, col, kind))
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
    Excelのオートフィルタのような直感的な列フィルター＆ソートポップアップ
    - 昇順 / 降順ソート
    - 数値条件フィルター（数値列のみ）
    - テキスト検索＋チェックボックスによる値選択（高速Treeview使用）
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
    ):
        super().__init__(parent)
        self.title(f"フィルター: {col_name}")
        self.configure(bg=COLORS["bg_card"])
        self.transient(parent)
        self.grab_set()

        self.col_name = col_name
        self.is_numeric = is_numeric
        self.value_counts = value_counts
        self.on_sort = on_sort
        self.on_apply_filter = on_apply_filter
        self.on_clear_filter = on_clear_filter

        # ウィンドウ位置とサイズ調整（画面外はみ出し防止）
        w = 340
        h = 510 if is_numeric else 450
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        px = min(max(10, x_root), max(10, sw - w - 20))
        py = min(max(10, y_root), max(10, sh - h - 50))
        self.geometry(f"{w}x{h}+{px}+{py}")
        self.resizable(True, True)

        main = ttk.Frame(self, style="Card.TFrame", padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        # ヘッダー＆ソートボタン
        ttk.Label(main, text=f"列「{col_name}」の並べ替え・絞り込み", style="Title.TLabel").pack(anchor=tk.W, pady=(0, 6))

        sort_f = ttk.Frame(main, style="Card.TFrame")
        sort_f.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(sort_f, text="↑ 昇順ソート", command=self._do_sort_asc).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 3))
        ttk.Button(sort_f, text="↓ 降順ソート", command=self._do_sort_desc).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(3, 0))

        ttk.Button(main, text="この列のフィルターを解除", command=self._do_clear).pack(fill=tk.X, pady=(0, 6))

        # 数値条件フィルター（数値列の場合）
        self.num_op_var = tk.StringVar(value=(current_filter or {}).get("num_op", "指定なし"))
        self.num_v1_var = tk.StringVar(value=str((current_filter or {}).get("num_v1", "")))
        self.num_v2_var = tk.StringVar(value=str((current_filter or {}).get("num_v2", "")))

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

            self.ent_v1 = ttk.Entry(row1, textvariable=self.num_v1_var, width=8)
            self.ent_v1.pack(side=tk.LEFT, padx=2)
            self.lbl_tilde = ttk.Label(row1, text="～", style="Card.TLabel")
            self.ent_v2 = ttk.Entry(row1, textvariable=self.num_v2_var, width=8)
            self._toggle_num_inputs()

        # 検索ボックス
        search_f = ttk.Frame(main, style="Card.TFrame")
        search_f.pack(fill=tk.X, pady=(2, 4))
        ttk.Label(search_f, text="検索:", style="Card.TLabel").pack(side=tk.LEFT, padx=(0, 4))
        self.search_var = tk.StringVar()
        ent_s = ttk.Entry(search_f, textvariable=self.search_var)
        ent_s.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ent_s.bind("<KeyRelease>", lambda e: self._rebuild_tree())
        ent_s.focus_set()

        # すべて選択トグル
        self.all_selected_var = tk.BooleanVar(value=True)
        chk_all = tk.Checkbutton(
            main,
            text="(すべて選択 / 検索結果を全選択)",
            variable=self.all_selected_var,
            command=self._toggle_select_all,
            bg=COLORS["bg_card"],
            fg=COLORS["fg_main"],
            activebackground=COLORS["bg_card"],
            selectcolor="#FFFFFF",
            font=("Yu Gothic UI", 9, "bold"),
            anchor=tk.W,
        )
        chk_all.pack(fill=tk.X, pady=(0, 2))

        # 値チェックリスト（高速Treeview）
        tv_frame = ttk.Frame(main, style="Card.TFrame")
        tv_frame.pack(fill=tk.BOTH, expand=True)

        self.tv = ttk.Treeview(tv_frame, columns=("chk", "val", "cnt"), show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(tv_frame, orient=tk.VERTICAL, command=self.tv.yview)
        self.tv.configure(yscrollcommand=vsb.set)
        self.tv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tv.heading("chk", text="選択")
        self.tv.heading("val", text="値")
        self.tv.heading("cnt", text="件数")
        self.tv.column("chk", width=44, anchor=tk.CENTER, stretch=False)
        self.tv.column("val", width=190, anchor=tk.W)
        self.tv.column("cnt", width=55, anchor=tk.E, stretch=False)

        self.tv.bind("<Button-1>", self._on_tv_click)
        self.tv.bind("<space>", self._on_tv_space)

        # 全ユニーク値のソート（数値列は数値順、それ以外は文字列順）
        if is_numeric:
            self.all_values = sorted(
                value_counts.keys(),
                key=lambda v: (v == "", to_float(v, 0.0)),
            )
        else:
            self.all_values = sorted(value_counts.keys(), key=lambda v: (v == "", str(v)))

        selected_set = (current_filter or {}).get("selected")
        if selected_set is None:
            self.checked_map = {v: True for v in self.all_values}
        else:
            self.checked_map = {v: (v in selected_set) for v in self.all_values}

        self.visible_values = []
        self._rebuild_tree()

        # フッターボタン
        footer = ttk.Frame(main, style="Card.TFrame")
        footer.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(footer, text="キャンセル", command=self.destroy).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(footer, text="OK（適用）", style="Primary.TButton", command=self._do_ok).pack(side=tk.RIGHT)

        self.bind("<Escape>", lambda e: self.destroy())
        self.bind("<Return>", lambda e: self._do_ok())

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

    def _display_val(self, raw_v: str) -> str:
        return "(空白)" if raw_v == "" else raw_v

    def _rebuild_tree(self):
        kw = self.search_var.get().strip().lower()
        self.tv.delete(*self.tv.get_children())
        self.visible_values = []
        for v in self.all_values:
            disp = self._display_val(v)
            if kw and kw not in disp.lower():
                continue
            self.visible_values.append(v)

        all_on = bool(self.visible_values)
        for idx, v in enumerate(self.visible_values):
            checked = self.checked_map.get(v, True)
            if not checked:
                all_on = False
            mark = "☑" if checked else "☐"
            self.tv.insert("", tk.END, iid=str(idx), values=(mark, self._display_val(v), f"{self.value_counts[v]:,}"))
        self.all_selected_var.set(all_on)

    def _toggle_select_all(self):
        target = self.all_selected_var.get()
        for v in self.visible_values:
            self.checked_map[v] = target
        self._rebuild_tree()

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
            new_state = not self.checked_map.get(val, True)
            self.checked_map[val] = new_state
            mark = "☑" if new_state else "☐"
            self.tv.item(row_id, values=(mark, self._display_val(val), f"{self.value_counts[val]:,}"))
            self.all_selected_var.set(all(self.checked_map.get(v, True) for v in self.visible_values))

    def _on_tv_space(self, event):
        sel = self.tv.selection()
        if not sel:
            return
        row_id = sel[0]
        idx = int(row_id)
        if 0 <= idx < len(self.visible_values):
            val = self.visible_values[idx]
            new_state = not self.checked_map.get(val, True)
            self.checked_map[val] = new_state
            mark = "☑" if new_state else "☐"
            self.tv.item(row_id, values=(mark, self._display_val(val), f"{self.value_counts[val]:,}"))
            self.all_selected_var.set(all(self.checked_map.get(v, True) for v in self.visible_values))

    def _do_sort_asc(self):
        self.on_sort(self.col_name, False)
        self.destroy()

    def _do_sort_desc(self):
        self.on_sort(self.col_name, True)
        self.destroy()

    def _do_clear(self):
        self.on_clear_filter(self.col_name)
        self.destroy()

    def _do_ok(self):
        kw = self.search_var.get().strip()
        if kw:
            # 検索中にOKした場合は、検索結果の中でチェックされている値のみに絞る（Excelと同じ挙動）
            selected_set = {v for v in self.visible_values if self.checked_map.get(v, True)}
        else:
            if all(self.checked_map.get(v, True) for v in self.all_values):
                selected_set = None
            else:
                selected_set = {v for v in self.all_values if self.checked_map.get(v, True)}

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
        self.destroy()


class ColumnCopyDialog(tk.Toplevel):
    """指定した列（単一または複数）をExcel貼り付け用形式でクリップボードにコピーするダイアログ"""

    def __init__(self, parent, visible_columns: list, displayed_count: int, selected_count: int, on_copy):
        super().__init__(parent)
        self.title("列データのコピー（Excel貼付用）")
        self.geometry("480x560")
        self.minsize(420, 460)
        self.configure(bg=COLORS["bg_card"])
        self.transient(parent)
        self.grab_set()

        self.visible_columns = list(visible_columns)
        self.on_copy = on_copy
        self.checked_map = {c: (c in CALC_COLUMNS or c == "管理No") for c in self.visible_columns}

        frame = ttk.Frame(self, style="Card.TFrame", padding=14)
        frame.pack(fill=tk.BOTH, expand=True)

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
        tv_frame.pack(fill=tk.BOTH, expand=True)

        self.tv = ttk.Treeview(tv_frame, columns=("chk", "col", "kind"), show="headings", selectmode="browse")
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

        opt_box = ttk.LabelFrame(frame, text="コピー形式オプション", padding=(8, 6))
        opt_box.pack(fill=tk.X, pady=(8, 6))

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

        footer = ttk.Frame(frame, style="Flat.TFrame")
        footer.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(footer, text="キャンセル", command=self.destroy).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(footer, text="クリップボードにコピー", style="Primary.TButton", command=self._do_copy).pack(side=tk.RIGHT)

        self._refresh()

    def _refresh(self):
        self.tv.delete(*self.tv.get_children())
        for idx, col in enumerate(self.visible_columns):
            mark = "☑" if self.checked_map.get(col, False) else "☐"
            kind = "集計・小計" if col in CALC_COLUMNS else "基本情報"
            self.tv.insert("", tk.END, iid=str(idx), values=(mark, col, kind))

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
            self.tv.item(row_id, values=(mark, col, kind))

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
    """起動時のデフォルト表示単位（金額/時間）と日報単価レート（円/分）設定ダイアログ"""

    def __init__(self, parent, current_rates: dict, default_unit: str, current_unit: str, on_apply):
        super().__init__(parent)
        self.title("作業単価レート・既定表示設定")
        self.geometry("540x485")
        self.resizable(False, False)
        self.configure(bg=COLORS["bg_card"])
        self.transient(parent)
        self.grab_set()

        self.on_apply = on_apply

        frame = ttk.Frame(self, style="Card.TFrame", padding=16)
        frame.pack(fill=tk.BOTH, expand=True)

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
        unit_row = ttk.Frame(unit_box, style="Flat.TFrame")
        unit_row.pack(fill=tk.X)
        for u_val, u_lbl in (
            ("金額(円)", "金額(円)（デフォルト）"),
            ("分(そのまま)", "時間：分(そのまま)"),
            ("時間(分÷60)", "時間：時間(分÷60)"),
        ):
            tk.Radiobutton(
                unit_row,
                text=u_lbl,
                variable=self.default_unit_var,
                value=u_val,
                bg=COLORS["bg_card"],
                fg=COLORS["fg_main"],
                activebackground=COLORS["bg_card"],
                selectcolor="#FFFFFF",
                font=("Yu Gothic UI", 9),
            ).pack(side=tk.LEFT, padx=(0, 12))

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

        btn_f = ttk.Frame(frame, style="Flat.TFrame")
        btn_f.pack(fill=tk.X, pady=(14, 0))

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
                    if val < 0:
                        raise ValueError("negative")
                except ValueError:
                    messagebox.showwarning(
                        "入力確認",
                        f"「{cat}」のレートには0以上の数値を入力してください。\n例: 62.5",
                        parent=self,
                    )
                    return
                new_rates[cat] = val
            self.on_apply(new_rates, self.default_unit_var.get(), self.switch_now_var.get())
            self.destroy()

        ttk.Button(btn_f, text="デフォルトに戻す", command=restore_defaults).pack(side=tk.LEFT)
        ttk.Button(btn_f, text="キャンセル", command=self.destroy).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(btn_f, text="保存して適用", style="Primary.TButton", command=save_and_apply).pack(side=tk.RIGHT)


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
    """見やすく整理された目次付き・カード型ヘルプダイアログ"""

    def __init__(self, parent):
        super().__init__(parent)
        self.title(f"使い方ガイド - {APP_NAME}")
        self.geometry("920x660")
        self.minsize(780, 520)
        self.configure(bg=COLORS["bg_main"])
        self.transient(parent)

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
        ttk.Button(ftr, text="閉じる", style="Primary.TButton", command=self.destroy).pack(side=tk.RIGHT)

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
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self.canvas_win, width=e.width))
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
        self.bind("<Destroy>", lambda e: self._cleanup_bind())

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
                width=22,
                cursor="hand2",
                command=lambda k=key: self._scroll_to(k),
            )
            btn.pack(fill=tk.X, pady=2)
            self.nav_buttons[key] = btn

        self._build_sections()

    def _cleanup_bind(self):
        try:
            self.canvas.unbind_all("<MouseWheel>")
        except Exception:
            pass

    def _on_mousewheel(self, event):
        try:
            if not self.winfo_exists():
                return
            x, y = self.winfo_pointerxy()
            w = self.winfo_containing(x, y)
            if w and str(w).startswith(str(self)):
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
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
        tk.Label(
            self.inner,
            text=text,
            bg=COLORS["bg_card"],
            fg=COLORS["fg_main"],
            font=("Yu Gothic UI", 10),
            justify=tk.LEFT,
            anchor=tk.W,
            wraplength=580,
        ).pack(fill=tk.X, pady=(0, 6))

    def _add_callout(self, title: str, text: str, bg_color: str, border_color: str):
        card = tk.Frame(self.inner, bg=bg_color, highlightbackground=border_color, highlightthickness=1, padx=12, pady=8)
        card.pack(fill=tk.X, pady=(4, 10))
        tk.Label(card, text=title, bg=bg_color, fg=COLORS["fg_main"], font=("Yu Gothic UI", 10, "bold"), anchor=tk.W).pack(fill=tk.X)
        tk.Label(
            card,
            text=text,
            bg=bg_color,
            fg=COLORS["fg_sub"],
            font=("Yu Gothic UI", 9),
            justify=tk.LEFT,
            anchor=tk.W,
            wraplength=560,
        ).pack(fill=tk.X, pady=(2, 0))

    def _add_table(self, headers: list, rows: list):
        tbl = tk.Frame(self.inner, bg=COLORS["border"], padx=1, pady=1)
        tbl.pack(fill=tk.X, pady=(4, 10))
        grid = tk.Frame(tbl, bg=COLORS["bg_card"])
        grid.pack(fill=tk.X)

        for c_i, h in enumerate(headers):
            lbl = tk.Label(
                grid,
                text=h,
                bg=COLORS["bg_header"],
                fg=COLORS["accent"],
                font=("Yu Gothic UI", 9, "bold"),
                padx=8,
                pady=5,
                anchor=tk.W,
            )
            lbl.grid(row=0, column=c_i, sticky=tk.NSEW, padx=1, pady=1)
            grid.columnconfigure(c_i, weight=1)

        for r_i, r_vals in enumerate(rows, 1):
            bg = "#FFFFFF" if r_i % 2 == 1 else COLORS["bg_sub"]
            for c_i, val in enumerate(r_vals):
                tk.Label(
                    grid,
                    text=val,
                    bg=bg,
                    fg=COLORS["fg_main"],
                    font=("Yu Gothic UI", 9),
                    padx=8,
                    pady=4,
                    anchor=tk.W,
                ).grid(row=r_i, column=c_i, sticky=tk.NSEW, padx=1, pady=1)

    def _build_sections(self):
        # 1. 概要
        self._add_section_header("overview", "STEP 1", "アプリの概要と基本の流れ")
        self._add_paragraph(
            "DenpyoHub_Miya は、日報ファイル(CSV)・購入品(XLSM/XLSX)・外注加工(XLSM/XLSX)・受注表(XLSM/XLSX)および"
            "制作伝票フォルダ（サブフォルダ含む .xls / .xlsx / .xlsm）を読み込み、管理Noごとに作業時間・金額と購入・外注・材料費用を自動集計するデスクトップアプリです。"
        )
        self._add_table(
            ["順番", "操作", "内容"],
            [
                ["1", "参照ファイルを選択", "起動時は自動集計しません。各ファイル・制作伝票フォルダを選択します（複数選択・サブフォルダ対応）。"],
                ["2", "「集計・一覧更新」を押す", "ボタンを押すとファイルを検証・読み込み、管理Noごとに一覧テーブルとサマリーを作成します。"],
                ["3", "絞り込み・確認", "Excel風フィルターや行リンク表示で、明細（日報・購入・外注・受注・制作伝票）を確認します。"],
                ["4", "Excel新規ブックで開く", "現在の一覧（列順番・フィルター・表示単位・見出し反映済み）をExcelの新規ブックとして直接開きます。"],
            ],
        )

        # 2. ファイル複数選択と読み取り専用保護・フォーマット検証
        self._add_section_header("files", "STEP 2", "参照ファイル＆制作伝票フォルダ（読み取り専用保護・フォーマット検証）")
        self._add_paragraph(
            "・日報／購入品／外注加工／受注表は各行の「選択(複数可)...」や「追加...」で複数ファイルを選択できます。\n"
            "・制作伝票フォルダの行では「フォルダ選択...」でフォルダを指定すると、サブフォルダを含むすべての .xls / .xlsx / .xlsm ファイルを自動検出し、調整用(21行)・材料費各種(28～35行)・設計費(36行)・小計1を統合します。\n"
            "・各サンプルファイルとフォーマットが異なるファイルが含まれる場合はエラー通知が表示され、「エラー確認」ボタンから原因・期待形式・検出内容を詳しく確認できます。"
        )
        self._add_callout(
            "参照元ファイル・フォルダは変更されないよう厳重に保護されています",
            "選択した日報・購入品・外注加工・受注表・制作伝票フォルダ（サブフォルダ含む）内の全ファイルは、常にメモリ上の読み取り専用データとして参照され、アプリから元のファイルやフォルダを書き換えることは一切ありません。\n"
            "また、設定保存や出力時に参照元ファイルや制作伝票フォルダ内への書き込みを指定した場合も自動で厳重にブロックします。",
            COLORS["bg_header"],
            COLORS["accent"],
        )

        # 3. 作業コード・レート設定
        self._add_section_header("codes", "STEP 3", "表示切替（金額 / 時間）と単価レート・作業コード設定")
        self._add_paragraph(
            "・初期設定では一覧の作業データは「金額(円)」で表示され、見出しも「作業費(円)」「移動費(円)」「修理改造費(円)」等に分かりやすく切り替わります。\n"
            "・一覧上部の「一括表示」ボタン（金額(円) / 時間(分) / 時間(h)）からワンクリックで一括切り替えできます。\n"
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
                ["制作伝票（フォルダ読込）", "21行(調整用), 28～35行(材料費), 36行(設計費)", "-", "作業小計＋調整用＋購入品小計＋材料費計＋設計費 ＝ 小計1"],
            ],
        )

        # 4. 一覧とフィルター
        self._add_section_header("filter", "STEP 4", "左右分割一覧（行リンク連動）とExcel風フィルター")
        self._add_paragraph(
            "・画面中央の一覧は左（その他見出し・受注情報）と右（日報/購入/外注/制作伝票の算出データ・各小計）に分かれています。\n"
            "・「行リンク連動」がONの間は、縦スクロール・並び替え・選択行のハイライトが左右で自動的に揃うため、横並びで比較しやすくなります（横スクロールやセル選択・コピーは左右別々に操作できます）。\n"
            "・セルや行をクリックすると、左右の該当行が青く連動強調され、画面下部のタブにその管理Noの明細が表示されます。\n"
            "・ファイル設定の折りたたみ：「▲ ファイル設定を畳む」で一覧をより広く表示できます"
        )
        self._add_callout(
            "Excel風オートフィルタの使い方",
            "各列の見出しの右端「▽」をクリック（または見出し上で右クリック）すると、Excelと同様の絞り込みポップアップが開きます。\n"
            "値のチェックボックス選択・検索・数値条件（0より大きい、範囲指定など）・昇順/降順ソートが利用できます。",
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
            "・上部の「列を選んでコピー...」ボタン：左右をまたいで複数列をまとめて選択コピー"
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
        self._add_section_header("export", "STEP 8", "Excelの新規ブックとして開く")
        self._add_paragraph(
            "「Excel新規ブックで開く」ボタンを押すと、ファイルを上書き・一時保存することなく、\n"
            "現在表示されているフィルター条件・表示単位・列の順番のまま、Excelの新規ブック（未保存ブック）として直接開きます。"
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

        # 上部：テーブル本体（残り領域をすべて使用）
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
        if self._overlay_after_id is not None:
            try:
                self.app.root.after_cancel(self._overlay_after_id)
            except Exception:
                pass
            self._overlay_after_id = None

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
        """このペイン内の選択セル数・数値合計・平均をリアルタイム表示する"""
        row_total = len(self.rows)
        if not self.selected_cells:
            self.cell_info_var.set(f"件数: {row_total:,}件  |  選択: なし（セルクリック/ドラッグで選択してボタンコピー可）")
            return
        cnt = len(self.selected_cells)
        rows_cnt = len({r for r, _ in self.selected_cells})
        cols_cnt = len({c for _, c in self.selected_cells})
        num_vals = []
        for r_idx, c_name in self.selected_cells:
            if 0 <= r_idx < len(self.rows) and c_name in NUMERIC_COLUMNS:
                r_dict = self.rows[r_idx]
                if c_name in WORK_CATEGORIES or c_name in ("作業小計", "小計1"):
                    num_vals.append(self.app._get_work_numeric_val(r_dict, c_name))
                else:
                    num_vals.append(to_float(r_dict.get(c_name, 0.0), 0.0))

        unit = self.app.time_unit_var.get()
        unit_lbl = "円" if unit == "金額(円)" else ("h" if unit == "時間(分÷60)" else "分")
        if num_vals:
            s_val = sum(num_vals)
            avg_val = s_val / len(num_vals)
            s_str = f"{int(round(s_val)):,}" if abs(s_val - round(s_val)) < 1e-6 else f"{s_val:,.2f}"
            a_str = f"{int(round(avg_val)):,}" if abs(avg_val - round(avg_val)) < 1e-6 else f"{avg_val:,.2f}"
            self.cell_info_var.set(
                f"件数: {row_total:,}件  |  選択: {rows_cnt:,}行×{cols_cnt}列 ({cnt:,}セル)  |  合計: {s_str}  平均: {a_str} ({unit_lbl})"
            )
        elif cnt == 1:
            r_idx, c_name = next(iter(self.selected_cells))
            val_txt = self.app._get_cell_display_str(self.rows[r_idx], c_name) if 0 <= r_idx < len(self.rows) else ""
            if len(val_txt) > 30:
                val_txt = val_txt[:27] + "..."
            self.cell_info_var.set(f"件数: {row_total:,}件  |  選択 [{c_name}]: {val_txt}")
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
            if event.x >= col_end_x - 28 and col_end_x > 0:
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
            self.header_menu.delete(0, tk.END)
            self.header_menu.add_command(
                label=f"「{col_name}」列をまるごと選択",
                command=lambda c=col_name: self.select_entire_column(c, add_mode=False),
            )
            self.header_menu.add_command(
                label=f"「{col_name}」列をまるごと追加選択 (Ctrl相当)",
                command=lambda c=col_name: self.select_entire_column(c, add_mode=True),
            )
            self.header_menu.add_separator()
            self.header_menu.add_command(
                label=f"「{col_name}」列の全データをコピー（見出しなし・Excel貼付用）",
                command=lambda c=col_name: self.copy_columns_to_clipboard([c], include_header=False, raw_numbers=True, selected_only=False),
            )
            self.header_menu.add_command(
                label=f"「{col_name}」列の全データをコピー（見出し付き・Excel貼付用）",
                command=lambda c=col_name: self.copy_columns_to_clipboard([c], include_header=True, raw_numbers=True, selected_only=False),
            )
            self.header_menu.add_command(
                label="複数の列を選んでコピー（Excel貼付用）...",
                command=self.app.open_column_copy_dialog,
            )
            self.header_menu.add_separator()
            self.header_menu.add_command(
                label=f"「{col_name}」でExcel風フィルターを開く...",
                command=lambda c=col_name: self.app.open_excel_filter_popup(c, event.x_root, self.tree.winfo_rooty() + 30),
            )
            if col_name in self.app.column_filters:
                self.header_menu.add_command(
                    label=f"「{col_name}」のフィルターを解除",
                    command=lambda c=col_name: self.app._clear_single_filter(c),
                )
            self.header_menu.add_separator()
            self.header_menu.add_command(label=f"この一覧を「{col_name}」昇順に並べ替え", command=lambda c=col_name: self.set_sort(c, False))
            self.header_menu.add_command(label=f"この一覧を「{col_name}」降順に並べ替え", command=lambda c=col_name: self.set_sort(c, True))
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
                    label=f"このセル（{col_name}）の値をコピー",
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
                    label=f"「{col_name}」列をまるごと選択",
                    command=lambda c=col_name: self.select_entire_column(c, add_mode=False),
                )
                self.row_menu.add_command(
                    label=f"「{col_name}」列の全データをコピー（見出しなし）",
                    command=lambda c=col_name: self.copy_columns_to_clipboard([c], include_header=False, raw_numbers=True, selected_only=False),
                )
                self.row_menu.add_command(
                    label=f"「{col_name}」列の全データをコピー（見出し付き）",
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
                    label=f"「{col_name}」列でフィルター...",
                    command=lambda c=col_name: self.app.open_excel_filter_popup(c, event.x_root, event.y_root),
                )
                self.row_menu.add_command(label=f"「{col_name}」列の幅を自動調整", command=lambda c=col_name: self.autofit_column(c))
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
            if col in WORK_CATEGORIES or col in ("作業小計", "小計1"):
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
        if include_header:
            unit = self.app.time_unit_var.get()
            lines.append("\t".join(escape_tsv_cell(get_display_column_label(c, unit)) for c in cols_sorted))
        for r_idx in rows_sorted:
            r_dict = self.rows[r_idx]
            row_cells = []
            for col in cols_sorted:
                if (r_idx, col) not in self.selected_cells:
                    row_cells.append("")
                    continue
                if raw_numbers:
                    if col in WORK_CATEGORIES or col in ("作業小計", "小計1"):
                        val = self.app._get_work_numeric_val(r_dict, col)
                        cell_str = str(int(round(val))) if abs(val - round(val)) < 1e-6 else f"{val:.2f}"
                    elif col in COST_YEN_COLUMNS:
                        cell_str = str(int(round(r_dict.get(col, 0.0))))
                    else:
                        cell_str = str(r_dict.get(col, ""))
                else:
                    cell_str = self.app._get_cell_display_str(r_dict, col)
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
        if include_header:
            unit = self.app.time_unit_var.get()
            lines.append("\t".join(escape_tsv_cell(get_display_column_label(c, unit)) for c in columns))
        for r in target_rows:
            row_cells = []
            for col in columns:
                if raw_numbers:
                    if col in WORK_CATEGORIES or col in ("作業小計", "小計1"):
                        val = self.app._get_work_numeric_val(r, col)
                        cell_str = str(int(round(val))) if abs(val - round(val)) < 1e-6 else f"{val:.2f}"
                    elif col in COST_YEN_COLUMNS:
                        cell_str = str(int(round(r.get(col, 0.0))))
                    else:
                        cell_str = str(r.get(col, ""))
                else:
                    cell_str = self.app._get_cell_display_str(r, col)
                row_cells.append(escape_tsv_cell(cell_str))
            lines.append("\t".join(row_cells))

        self.app.root.clipboard_clear()
        self.app.root.clipboard_append("\r\n".join(lines))
        col_label = columns[0] if len(columns) == 1 else f"{len(columns)}列"
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

            if col in WORK_CATEGORIES or col in ("作業小計", "小計1"):
                get_w = self.app._get_work_numeric_val
                sort_key = lambda item: get_w(item, col)
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
        render_id = self._render_job_id
        self.tree.delete(*self.tree.get_children())

        total_rows = len(self.rows)
        if total_rows <= 600:
            self._insert_chunk(0, total_rows, render_id)
        else:
            # 最初の画面表示分（150行）を即座に描画し、残りは1200行単位で高速バッチ挿入
            self._insert_chunk(0, 150, render_id)

    def _insert_chunk(self, start_idx: int, chunk_size: int, render_id: int):
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
        tree_insert = self.tree.insert
        end_pos = tk.END

        for idx in range(start_idx, end_idx):
            r = rows_ref[idx]
            values = [get_disp(r, c_name) for c_name in vis_cols]
            tag = "active_both" if (r["作業小計"] > 0 and r["購入品小計"] > 0) else ("even" if (idx & 1) == 0 else "odd")
            tree_insert("", end_pos, iid=r["管理No"], values=values, tags=(tag,))

        if end_idx < total_len:
            self.app.root.after(1, lambda: self._insert_chunk(end_idx, 1200, render_id))
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

        # 設定ファイルの管理（デフォルトはアプリと同じ場所・旧ファイル名からの移行読込に対応）
        self.default_config_path = os.path.join(get_base_dir(), CONFIG_FILENAME)
        self.config_path = self.default_config_path
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
                self.work_rates[cat] = val if val >= 0 else def_rate
            except (TypeError, ValueError):
                self.work_rates[cat] = def_rate

        # 列の順番・表示・幅設定（新しく追加された列は既存設定引き継ぎ時に自動表示）
        raw_saved_order = self.config.get("column_order", list(DEFAULT_COLUMN_ORDER))
        saved_order_set = set(raw_saved_order)
        self.column_order = [c for c in raw_saved_order if c in ALL_AVAILABLE_COLUMNS]
        for c in ALL_AVAILABLE_COLUMNS:
            if c not in self.column_order:
                self.column_order.append(c)

        saved_visible = self.config.get("visible_columns", list(DEFAULT_COLUMN_ORDER))
        self.visible_set = {c for c in saved_visible if c in ALL_AVAILABLE_COLUMNS}
        if not self.visible_set:
            self.visible_set = set(DEFAULT_COLUMN_ORDER)
        for c in ALL_AVAILABLE_COLUMNS:
            if c not in saved_order_set and c != "備考":
                self.visible_set.add(c)
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

        # ファイル設定エリア開閉状態
        self.file_panel_expanded = bool(self.config.get("file_panel_expanded", True))

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
        self.root.after(20, self._maximize_window)
        self.root.after(120, self._startup_async_init)

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
        ttk.Button(header_frame, text="フォルダから一括自動選択", command=self.select_folder_auto).pack(side=tk.RIGHT, padx=3)

        self.btn_toggle_files = ttk.Button(header_frame, text="▲ ファイル設定を畳む", command=self.toggle_file_panel)
        self.btn_toggle_files.pack(side=tk.RIGHT, padx=3)

        # 2. ファイル選択＆更新日表示エリア（折りたたみ可能）
        self.file_group = ttk.LabelFrame(
            self.root,
            text="参照ファイル設定（複数選択可・制作伝票フォルダはサブフォルダ内も自動参照・すべて読み取り専用保護）",
            padding=(10, 4),
        )
        self.file_group.pack(side=tk.TOP, fill=tk.X, padx=10, pady=2)

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

        # 5行目: 制作伝票フォルダ (.xls / .xlsx / .xlsm・サブフォルダ含む)
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
            text="フォルダ選択",
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
        self.search_entry.pack(side=tk.LEFT, padx=(0, 4))
        self.search_entry.bind("<KeyRelease>", self._on_search_key_release)

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
        for u_val, u_text in (("金額(円)", "金額(円)"), ("分(そのまま)", "時間(分)"), ("時間(分÷60)", "時間(h)")):
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
        self.progress_bar = ttk.Progressbar(status_bar, mode="indeterminate", length=140)
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

        self._configure_main_tree_columns()
        self._update_unit_button_styles()
        self.root.after(100, self._init_horizontal_sash)

        # 下部：選択した管理Noの内訳明細ビュー
        detail_frame = ttk.LabelFrame(self.v_paned, text="選択した管理Noの内訳明細", padding=(6, 3))
        self.v_paned.add(detail_frame, minsize=145, stretch="never")

        self.detail_notebook = ttk.Notebook(detail_frame)
        self.detail_notebook.pack(fill=tk.BOTH, expand=True)

        # タブ1: 日報明細
        tab_daily = ttk.Frame(self.detail_notebook, style="Card.TFrame")
        self.detail_notebook.add(tab_daily, text="日報データ明細 (0件)")
        daily_cols = ("更新日", "作業者", "作業コード", "作業名", "分類区分", "時間(分)", "時間(h)", "金額(円)", "出荷先", "製品詳細", "元ファイル")
        self.tree_daily = self._create_detail_tree(
            tab_daily,
            daily_cols,
            [135, 115, 72, 110, 95, 78, 78, 90, 95, 200, 140],
            right_cols={"作業コード", "時間(分)", "時間(h)", "金額(円)"},
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
        """明細表示用Treeviewを作成する"""
        frame = ttk.Frame(parent, style="Card.TFrame")
        frame.pack(fill=tk.BOTH, expand=True)
        tv = ttk.Treeview(frame, columns=columns, show="headings", height=4)
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
        return tv

    def _bind_global_mousewheel(self):
        """カーソル下のTreeviewに対して縦・Shift+横スクロールを滑らかに効かせる"""
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
            }
            self.calc_pane.configure(text=pane_titles.get(cur, "算出データ・各小計"))

    def set_time_unit(self, unit: str):
        """一覧全体の表示単位（金額／時間）をボタンから一括切り替えする"""
        if unit not in VALID_TIME_UNITS:
            return
        self.time_unit_var.set(unit)
        self._update_unit_button_styles()
        self.apply_filter_and_refresh()
        if self._current_detail_kno:
            self._show_detail_for_kno(self._current_detail_kno)
        label_map = {"金額(円)": "金額(円)", "分(そのまま)": "時間(分)", "時間(分÷60)": "時間(h)"}
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

    def _update_heading_texts(self):
        self.meta_pane.update_heading_texts()
        self.calc_pane.update_heading_texts()

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
        if include_header:
            unit = self.time_unit_var.get()
            lines.append("\t".join(escape_tsv_cell(get_display_column_label(c, unit)) for c in columns))

        for r in target_rows:
            row_cells = []
            for col in columns:
                if raw_numbers:
                    if col in WORK_CATEGORIES or col in ("作業小計", "小計1"):
                        val = self._get_work_numeric_val(r, col)
                        cell_str = str(int(round(val))) if abs(val - round(val)) < 1e-6 else f"{val:.2f}"
                    elif col in COST_YEN_COLUMNS:
                        cell_str = str(int(round(r.get(col, 0.0))))
                    else:
                        cell_str = str(r.get(col, ""))
                else:
                    cell_str = self._get_cell_display_str(r, col)
                row_cells.append(escape_tsv_cell(cell_str))
            lines.append("\t".join(row_cells))

        self.root.clipboard_clear()
        self.root.clipboard_append("\r\n".join(lines))
        col_label = columns[0] if len(columns) == 1 else f"{len(columns)}列"
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
        )

    def get_protected_source_paths(self) -> set:
        """参照ファイル（絶対変更禁止）の正規化パス集合を返す"""
        protected = set()
        for k, paths in self.file_lists.items():
            expanded = expand_seizo_excel_files(paths) if k == "seizo" else paths
            for p in expanded:
                if p:
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
            with open(tmp_save, "w", encoding="utf-8") as f:
                f.write(json_str)
                f.flush()
                try:
                    os.fsync(f.fileno())
                except Exception:
                    pass
            os.replace(tmp_save, save_to)

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
        self.visible_set = set(DEFAULT_COLUMN_ORDER)
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

        saved_visible = cfg.get("visible_columns", list(DEFAULT_COLUMN_ORDER))
        self.visible_set = {c for c in saved_visible if c in ALL_AVAILABLE_COLUMNS} or set(DEFAULT_COLUMN_ORDER)
        for c in ALL_AVAILABLE_COLUMNS:
            if c not in saved_order_set and c != "備考":
                self.visible_set.add(c)
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
        """検索キーワードが実際に変化したときだけ一覧フィルタリングを実行して高速化する"""
        cur_kw = self.search_var.get().strip().lower()
        if cur_kw == self._last_search_kw:
            return
        self._last_search_kw = cur_kw
        self.apply_filter_and_refresh()

    def _clear_search(self):
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

    def select_seizo_folder_replace(self):
        """制作伝票Excelが入ったフォルダを選択して置き換える"""
        cur = self.file_lists["seizo"]
        init_dir = cur[-1] if cur and os.path.exists(cur[-1]) else get_base_dir()
        folder = filedialog.askdirectory(
            parent=self.root,
            title="制作伝票Excel (.xls / .xlsx / .xlsm) が入っているフォルダを選択してください（サブフォルダも自動参照）",
            initialdir=init_dir,
        )
        if folder:
            self.file_lists["seizo"] = self._normalize_file_list([folder])
            self._refresh_file_mtime_async()
            self.save_config()

    def select_seizo_folder_append(self):
        """制作伝票Excelが入った別フォルダを追加する"""
        cur = self.file_lists["seizo"]
        init_dir = cur[-1] if cur and os.path.exists(cur[-1]) else get_base_dir()
        folder = filedialog.askdirectory(
            parent=self.root,
            title="追加する制作伝票フォルダを選択してください（サブフォルダも自動参照）",
            initialdir=init_dir,
        )
        if folder:
            self.file_lists["seizo"] = self._normalize_file_list(self.file_lists["seizo"] + [folder])
            self._refresh_file_mtime_async()
            self.save_config()

    def open_seizo_file_list_dialog(self):
        """制作伝票フォルダ（サブフォルダ含む）内の検出済みExcel一覧を表示するダイアログ"""
        dlg = tk.Toplevel(self.root)
        dlg.title("制作伝票フォルダ（サブフォルダ含む）・検出済みExcelファイル一覧")
        dlg.geometry("860x480")
        dlg.minsize(680, 360)
        dlg.configure(bg=COLORS["bg_card"])
        dlg.transient(self.root)
        dlg.grab_set()

        frame = ttk.Frame(dlg, style="Card.TFrame", padding=14)
        frame.pack(fill=tk.BOTH, expand=True)

        ftr = ttk.Frame(frame, style="Flat.TFrame")
        ftr.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))

        ttk.Label(
            frame,
            text="指定中の制作伝票フォルダ（サブフォルダ含む）内の対象Excel (.xls / .xlsx / .xlsm)",
            style="Title.TLabel",
        ).pack(anchor=tk.W, pady=(0, 6))

        base_folders = self.file_lists["seizo"]
        folders_str = " / ".join(base_folders) if base_folders else "(フォルダ未指定)"
        ttk.Label(frame, text=f"指定フォルダ: {folders_str}", style="Mtime.TLabel", wraplength=810).pack(anchor=tk.W, pady=(0, 8))

        tv_frame = ttk.Frame(frame, style="Card.TFrame")
        tv_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        cols = ("ファイル名", "相対パス（サブフォルダ）", "更新日時", "フルパス")
        tv = ttk.Treeview(tv_frame, columns=cols, show="headings")
        vsb = ttk.Scrollbar(tv_frame, orient=tk.VERTICAL, command=tv.yview)
        hsb = ttk.Scrollbar(tv_frame, orient=tk.HORIZONTAL, command=tv.xview)
        tv.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tv.grid(row=0, column=0, sticky=tk.NSEW)
        vsb.grid(row=0, column=1, sticky=tk.NS)
        hsb.grid(row=1, column=0, sticky=tk.EW)
        tv_frame.rowconfigure(0, weight=1)
        tv_frame.columnconfigure(0, weight=1)

        tv.heading("ファイル名", text="ファイル名")
        tv.heading("相対パス（サブフォルダ）", text="相対パス（サブフォルダ）")
        tv.heading("更新日時", text="更新日時")
        tv.heading("フルパス", text="フルパス")
        tv.column("ファイル名", width=165, anchor=tk.W)
        tv.column("相対パス（サブフォルダ）", width=190, anchor=tk.W)
        tv.column("更新日時", width=150, anchor=tk.CENTER, stretch=False)
        tv.column("フルパス", width=320, anchor=tk.W)

        files = expand_seizo_excel_files(base_folders)
        for fp in files:
            rel_disp = os.path.basename(fp)
            for bf in base_folders:
                if bf and os.path.isdir(bf):
                    try:
                        norm_fp = os.path.abspath(os.path.normpath(fp))
                        norm_bf = os.path.abspath(os.path.normpath(bf))
                        if os.path.commonpath([norm_fp, norm_bf]) == norm_bf:
                            rel_disp = os.path.relpath(norm_fp, norm_bf)
                            break
                    except Exception:
                        pass
            tv.insert("", tk.END, values=(os.path.basename(fp), rel_disp, format_file_mtime(fp), fp))

        ttk.Label(ftr, text=f"合計 {len(files):,} ファイル検出（サブフォルダ含む・すべて読み取り専用で参照）", style="Mtime.TLabel").pack(side=tk.LEFT)
        ttk.Button(ftr, text="閉じる", style="Primary.TButton", command=dlg.destroy).pack(side=tk.RIGHT)

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

    def select_folder_auto(self):
        """指定フォルダ内から4種別のファイルおよび制作伝票サブフォルダを自動検出してセットする"""
        init_dir = get_base_dir()
        for paths in self.file_lists.values():
            if paths and os.path.exists(os.path.dirname(paths[0])):
                init_dir = os.path.dirname(paths[0])
                break

        folder = filedialog.askdirectory(
            parent=self.root,
            title="データファイルが入っているフォルダを選択してください",
            initialdir=init_dir,
        )
        if not folder:
            return

        detected = {"daily": [], "kounyu": [], "gaichu": [], "juchu": [], "seizo": []}
        for fname in sorted(os.listdir(folder)):
            lower = fname.lower()
            full = os.path.normpath(os.path.join(folder, fname))
            if os.path.isdir(full) and fname in ("制作伝票", "製作伝票"):
                detected["seizo"].append(full)
                continue
            if not os.path.isfile(full) or fname.startswith("~$"):
                continue
            if lower == "daily_report.csv" or ("daily" in lower and lower.endswith(".csv")) or ("日報" in fname and lower.endswith(".csv")):
                detected["daily"].append(full)
            elif "購入品" in fname and (lower.endswith(".xlsm") or lower.endswith(".xlsx")):
                detected["kounyu"].append(full)
            elif "外注加工" in fname and (lower.endswith(".xlsm") or lower.endswith(".xlsx")):
                detected["gaichu"].append(full)
            elif ("受注" in fname) and (lower.endswith(".xlsm") or lower.endswith(".xlsx")):
                detected["juchu"].append(full)

        found_labels = []
        for k, label in [("daily", "日報"), ("kounyu", "購入品"), ("gaichu", "外注加工"), ("juchu", "受注表"), ("seizo", "制作伝票フォルダ")]:
            if detected[k]:
                self.file_lists[k] = self._normalize_file_list(detected[k])
                found_labels.append(f"{label}({len(detected[k])}件)")

        self._refresh_file_mtime_async()
        self.save_config()
        if found_labels:
            self.run_aggregation()
        else:
            messagebox.showwarning("自動検出", "選択したフォルダ内に対象ファイルが見つかりませんでした。", parent=self.root)

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
        return {
            "category": category,
            "file": os.path.basename(filepath) if filepath else "-",
            "path": filepath or "",
            "kind": default_kind,
            "reason": str(exc),
            "expected": "",
            "detected": "",
        }

    def run_aggregation(self, silent_missing: bool = False):
        """選択された全ファイル（複数可・サブフォルダ含む）をバックグラウンドで検証・読み込みし、管理Noごとに集計を行う"""
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
        try:
            self.btn_aggregate.configure(text="読み込み中...")
            self.progress_bar.pack(side=tk.RIGHT, padx=(8, 12))
            self.progress_bar.start(12)
        except Exception:
            pass

        def is_cancelled():
            return self._closing or (job_id != self._active_job_id)

        last_post_t = [0.0]

        def post_status(msg: str, force: bool = False):
            if is_cancelled():
                return
            now = time.perf_counter()
            if force or (now - last_post_t[0] >= 0.04):
                last_post_t[0] = now
                self._ui_queue.put(lambda m=msg: (not is_cancelled()) and self.status_var.set(m))

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
                # 0. 全ファイルの存在・更新日時チェック
                seizo_folders = file_snapshot.get("seizo", [])
                seizo_files = expand_seizo_excel_files(seizo_folders)
                for k, paths in file_snapshot.items():
                    if is_cancelled():
                        return
                    mtime_map[k] = summarize_seizo_folder_mtime(paths) if k == "seizo" else summarize_file_list_mtime(paths)

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

                any_exists = any(os.path.exists(p) for k, paths in file_snapshot.items() if k != "seizo" for p in paths) or bool(seizo_files)
                if not any_exists and not errors:
                    if not is_cancelled():
                        self._ui_queue.put(
                            lambda: self._on_aggregation_no_files(job_id, mtime_map, silent_missing)
                        )
                    return

                # 1. 日報ファイル（複数可）の読み込み＆フォーマット検証
                for idx, p in enumerate(file_snapshot["daily"], 1):
                    if is_cancelled():
                        return
                    post_status(f"[1/5] 日報ファイル読込中 ({idx}/{len(file_snapshot['daily'])}): {os.path.basename(p)}", force=(idx == 1))
                    if not os.path.exists(p):
                        errors.append(self._make_error_entry("日報ファイル", p, FileNotFoundError(p)))
                        continue
                    try:
                        recs, latest = read_daily_report_csv(p, cancel_check=is_cancelled)
                        raw_daily.extend(recs)
                        if latest and latest > daily_latest:
                            daily_latest = latest
                    except InterruptedError:
                        return
                    except Exception as e:
                        errors.append(self._make_error_entry("日報ファイル", p, e))

                if daily_latest:
                    mtime_map["daily"] = summarize_file_list_mtime(
                        file_snapshot["daily"],
                        extra_info=f"(データ内最新: {daily_latest})",
                    )

                # 2. 購入品ファイル（複数可）の読み込み＆フォーマット検証
                for idx, p in enumerate(file_snapshot["kounyu"], 1):
                    if is_cancelled():
                        return
                    post_status(f"[2/5] 購入品ファイル読込中 ({idx}/{len(file_snapshot['kounyu'])}): {os.path.basename(p)}", force=(idx == 1))
                    if not os.path.exists(p):
                        errors.append(self._make_error_entry("購入品ファイル", p, FileNotFoundError(p)))
                        continue
                    try:
                        rows, m_sheet, all_sh = read_excel_sheet_rows(
                            p, preferred_sheets=["テーブル"], cancel_check=is_cancelled, return_meta=True
                        )
                        validate_table_excel_rows(p, "kounyu", rows, matched_sheet=m_sheet, all_sheets=all_sh)
                        src_name = os.path.basename(p)
                        for r in rows[1:]:
                            r["_src_file"] = src_name
                            raw_kounyu.append(r)
                    except InterruptedError:
                        return
                    except Exception as e:
                        errors.append(self._make_error_entry("購入品ファイル", p, e))

                # 3. 外注加工ファイル（複数可）の読み込み＆フォーマット検証
                for idx, p in enumerate(file_snapshot["gaichu"], 1):
                    if is_cancelled():
                        return
                    post_status(f"[3/5] 外注加工ファイル読込中 ({idx}/{len(file_snapshot['gaichu'])}): {os.path.basename(p)}", force=(idx == 1))
                    if not os.path.exists(p):
                        errors.append(self._make_error_entry("外注加工ファイル", p, FileNotFoundError(p)))
                        continue
                    try:
                        rows, m_sheet, all_sh = read_excel_sheet_rows(
                            p, preferred_sheets=["テーブル"], cancel_check=is_cancelled, return_meta=True
                        )
                        validate_table_excel_rows(p, "gaichu", rows, matched_sheet=m_sheet, all_sheets=all_sh)
                        src_name = os.path.basename(p)
                        for r in rows[1:]:
                            r["_src_file"] = src_name
                            raw_gaichu.append(r)
                    except InterruptedError:
                        return
                    except Exception as e:
                        errors.append(self._make_error_entry("外注加工ファイル", p, e))

                # 4. 受注表ファイル（複数可）の読み込み＆フォーマット検証
                for idx, p in enumerate(file_snapshot["juchu"], 1):
                    if is_cancelled():
                        return
                    post_status(f"[4/5] 受注表ファイル読込中 ({idx}/{len(file_snapshot['juchu'])}): {os.path.basename(p)}", force=(idx == 1))
                    if not os.path.exists(p):
                        errors.append(self._make_error_entry("受注表ファイル", p, FileNotFoundError(p)))
                        continue
                    try:
                        rows, m_sheet, all_sh = read_excel_sheet_rows(
                            p, preferred_sheets=["受注表"], cancel_check=is_cancelled, return_meta=True
                        )
                        validate_table_excel_rows(p, "juchu", rows, matched_sheet=m_sheet, all_sheets=all_sh)
                        for r in rows[1:]:
                            raw_juchu.append(r)
                    except InterruptedError:
                        return
                    except Exception as e:
                        errors.append(self._make_error_entry("受注表ファイル", p, e))

                # 5. 制作伝票フォルダ（サブフォルダ含む .xls / .xlsx / .xlsm）の高速読み込み＆フォーマット検証
                total_seizo = len(seizo_files)
                if total_seizo > 0:
                    can_parallel = total_seizo >= 3 and all(not fp.lower().endswith(".xls") for fp in seizo_files)
                    if can_parallel:
                        post_status(f"[5/5] 制作伝票を一括並列読込中 (全{total_seizo}ファイル)...", force=True)
                        max_w = min(4, os.cpu_count() or 2)

                        def load_one_seizo(idx_fp):
                            idx_i, fp_i = idx_fp
                            if is_cancelled():
                                return idx_i, fp_i, None, None
                            post_status(f"[5/5] 制作伝票読込中 ({idx_i}/{total_seizo}): {os.path.basename(fp_i)}")
                            try:
                                rec = read_seizo_denpyo_sheet(fp_i, cancel_check=is_cancelled, base_folders=seizo_folders)
                                return idx_i, fp_i, rec, None
                            except InterruptedError:
                                return idx_i, fp_i, None, InterruptedError()
                            except Exception as e:
                                return idx_i, fp_i, None, e

                        with ThreadPoolExecutor(max_workers=max_w) as pool:
                            results = list(pool.map(load_one_seizo, enumerate(seizo_files, 1)))
                        for _, fp_i, rec, err in results:
                            if is_cancelled() or isinstance(err, InterruptedError):
                                return
                            if err is not None:
                                errors.append(self._make_error_entry("制作伝票ファイル", fp_i, err))
                            elif rec:
                                raw_seizo.append(rec)
                    else:
                        for idx, p in enumerate(seizo_files, 1):
                            if is_cancelled():
                                return
                            post_status(f"[5/5] 制作伝票読込中 ({idx}/{total_seizo}): {os.path.basename(p)}", force=(idx == 1))
                            try:
                                rec = read_seizo_denpyo_sheet(p, cancel_check=is_cancelled, base_folders=seizo_folders)
                                if rec:
                                    raw_seizo.append(rec)
                            except InterruptedError:
                                return
                            except Exception as e:
                                errors.append(self._make_error_entry("制作伝票ファイル", p, e))

                if is_cancelled():
                    return

                post_status("集計データを構築しています...", force=True)
                bundle = self._compute_aggregation_bundle(
                    raw_daily,
                    raw_kounyu,
                    raw_gaichu,
                    raw_juchu,
                    code_ranges_snapshot,
                    raw_seizo=raw_seizo,
                    work_rates=work_rates_snapshot,
                )
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

    def _stop_loading_indicator(self):
        self._is_loading = False
        try:
            self.btn_aggregate.configure(text="集計・一覧更新")
            self.progress_bar.stop()
            self.progress_bar.pack_forget()
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
        self._stop_loading_indicator()
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
        err_suffix = f" / エラー: {len(self.last_errors)} 件（エラー確認ボタンから詳細確認可）" if self.last_errors else ""
        self.status_var.set(
            f"集計完了: 全管理No {len(self.aggregated_rows):,} 件 "
            f"(日報: {len(self.raw_daily):,} 行 / 購入品: {len(self.raw_kounyu):,} 行 / "
            f"外注: {len(self.raw_gaichu):,} 行 / 受注表: {len(self.juchu_by_kno):,} 件 / 制作伝票: {len(self.raw_seizo):,} 件){err_suffix}"
        )
        if self.last_errors:
            ErrorReportDialog(self.root, self.last_errors)

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
        self.status_var.set(
            f"集計完了: 全管理No {len(self.aggregated_rows):,} 件 "
            f"(日報: {len(self.raw_daily):,} 行 / 購入品: {len(self.raw_kounyu):,} 行 / "
            f"外注: {len(self.raw_gaichu):,} 行 / 受注表: {len(self.juchu_by_kno):,} 件 / 制作伝票: {len(self.raw_seizo):,} 件)"
        )

    @staticmethod
    def _populate_row_rate_cache(row_dict: dict, rates_tuple: tuple):
        """1行分の円換算キャッシュ（作業費5区分・作業小計・小計1）を一括計算して格納する"""
        r0, r1, r2, r3, r4 = rates_tuple
        y0 = int(round(to_float(row_dict.get("作業時間", 0.0), 0.0) * r0))
        y1 = int(round(to_float(row_dict.get("移動時間", 0.0), 0.0) * r1))
        y2 = int(round(to_float(row_dict.get("修理改造時間", 0.0), 0.0) * r2))
        y3 = int(round(to_float(row_dict.get("設計関係", 0.0), 0.0) * r3))
        y4 = int(round(to_float(row_dict.get("その他作業", 0.0), 0.0) * r4))
        y_sub = y0 + y1 + y2 + y3 + y4
        y_sub1 = (
            y_sub
            + int(round(to_float(row_dict.get("調整用", 0.0), 0.0)))
            + int(round(to_float(row_dict.get("購入品小計", 0.0), 0.0)))
            + int(round(to_float(row_dict.get("材料費計", 0.0), 0.0)))
            + int(round(to_float(row_dict.get("設計費", 0.0), 0.0)))
        )
        row_dict["_rates_key"] = rates_tuple
        row_dict["_yen_作業時間"] = float(y0)
        row_dict["_yen_移動時間"] = float(y1)
        row_dict["_yen_修理改造時間"] = float(y2)
        row_dict["_yen_設計関係"] = float(y3)
        row_dict["_yen_その他作業"] = float(y4)
        row_dict["_yen_作業小計"] = float(y_sub)
        row_dict["_yen_小計1"] = float(y_sub1)

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

        # (C) 購入品データ集計（Q列: 金額 単価X必要数）
        kounyu_sums = defaultdict(float)
        for r in raw_kounyu:
            kno = normalize_kanri_no(r.get("A"))
            if not kno or kno == "管理No":
                continue
            all_knos.add(kno)
            q_val = r.get("Q")
            if q_val is not None and str(q_val).strip() != "":
                amount = to_float(q_val, 0.0)
            else:
                amount = float(int(to_float(r.get("M"), 0.0) * to_float(r.get("N"), 0.0)))
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

        # (D) 外注加工データ集計（Q列: 金額 単価X必要数）
        gaichu_sums = defaultdict(float)
        for r in raw_gaichu:
            kno = normalize_kanri_no(r.get("A"))
            if not kno or kno == "管理No":
                continue
            all_knos.add(kno)
            q_val = r.get("Q")
            if q_val is not None and str(q_val).strip() != "":
                amount = to_float(q_val, 0.0)
            else:
                amount = float(int(to_float(r.get("M"), 0.0) * to_float(r.get("N"), 0.0)))
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
        for srec in (raw_seizo or []):
            kno = srec.get("管理No", "")
            if not kno:
                continue
            all_knos.add(kno)
            seizo_by_kno[kno].append(srec)
            for c in SEIZO_COST_COLUMNS:
                seizo_sums[kno][c] += to_float(srec.get(c, 0.0), 0.0)
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
            search_haystack = f"{kno} {v_juchu_saki} {v_shukka_saki} {v_proj} {v_hinmei} {v_detail} {v_tehai} {v_uriage} {v_bikou}".lower()

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
        return (
            to_float(r.get("作業時間", 62.5), 62.5),
            to_float(r.get("移動時間", 50.0), 50.0),
            to_float(r.get("修理改造時間", 70.0), 70.0),
            to_float(r.get("設計関係", 62.5), 62.5),
            to_float(r.get("その他作業", 62.5), 62.5),
        )

    def _get_work_numeric_val(self, row: dict, col: str) -> float:
        """表示単位（分・時間・金額）と単価レートを反映した作業列および小計1の数値を高速に返す"""
        rates_tuple = self._get_rates_tuple()
        if row.get("_rates_key") != rates_tuple:
            self._populate_row_rate_cache(row, rates_tuple)

        if col == "小計1":
            return row["_yen_小計1"]
        unit = self.time_unit_var.get()
        if unit == "金額(円)":
            return row.get(f"_yen_{col}", 0.0)
        minutes_val = to_float(row.get(col, 0.0), 0.0)
        if unit == "時間(分÷60)":
            return round(minutes_val / 60.0, 2)
        return minutes_val

    def _format_converted_work_val(self, val: float) -> str:
        """表示単位に変換済みの作業数値（分・時間・円）を文字列フォーマットする"""
        unit = self.time_unit_var.get()
        if unit == "金額(円)":
            return f"{int(round(val)):,}"
        if unit == "時間(分÷60)":
            return f"{val:,.2f}"
        if abs(val - round(val)) < 1e-6:
            return f"{int(round(val)):,}"
        return f"{val:,.1f}"

    def _format_work_val(self, minutes_val: float, col: str = None) -> str:
        """時間単位の設定に応じて作業時間または金額をフォーマットする"""
        unit = self.time_unit_var.get()
        if unit == "金額(円)":
            if col in WORK_CATEGORIES:
                rate = to_float(self.work_rates.get(col, DEFAULT_WORK_RATES[col]), DEFAULT_WORK_RATES[col])
                return f"{int(round(minutes_val * rate)):,}"
            return f"{int(round(minutes_val)):,}"
        if unit == "時間(分÷60)":
            return f"{minutes_val / 60.0:,.2f}"
        if abs(minutes_val - round(minutes_val)) < 1e-6:
            return f"{int(round(minutes_val)):,}"
        return f"{minutes_val:,.1f}"

    def _get_work_export_val(self, minutes_val: float, row_dict: dict = None, col: str = None) -> float:
        """エクスポート時の作業数値（分・時間・円）を取得する"""
        if row_dict is not None and col is not None:
            return self._get_work_numeric_val(row_dict, col)
        if self.time_unit_var.get() == "金額(円)":
            if col in WORK_CATEGORIES:
                rate = to_float(self.work_rates.get(col, DEFAULT_WORK_RATES[col]), DEFAULT_WORK_RATES[col])
                return float(int(round(minutes_val * rate)))
            return float(int(round(minutes_val)))
        if self.time_unit_var.get() == "時間(分÷60)":
            return round(minutes_val / 60.0, 2)
        return minutes_val

    def _get_cell_display_str(self, row: dict, col: str) -> str:
        """一覧表示上のセル文字列を取得する（視認性向上のため単位付き表示・0はシンプル表示）"""
        unit = self.time_unit_var.get()
        if col == "小計1":
            iv = int(round(self._get_work_numeric_val(row, "小計1")))
            return f"{iv:,}円" if iv != 0 else "0"
        if col in WORK_CATEGORIES or col == "作業小計":
            val = self._get_work_numeric_val(row, col)
            if unit == "金額(円)":
                iv = int(round(val))
                return f"{iv:,}円" if iv != 0 else "0"
            if unit == "時間(分÷60)":
                return f"{val:,.2f}h" if abs(val) >= 1e-6 else "0.00"
            if abs(val) < 1e-6:
                return "0"
            if abs(val - round(val)) < 1e-6:
                return f"{int(round(val)):,}分"
            return f"{val:,.1f}分"
        elif col in COST_YEN_COLUMNS:
            iv = int(round(row.get(col, 0.0)))
            return f"{iv:,}円" if iv != 0 else "0"
        return str(row.get(col, ""))

    def _get_cell_filter_key(self, row: dict, col: str) -> str:
        """Excel風フィルター用の正規化値（数値列はカンマ無しの表示値、文字はそのまま）"""
        if col == "小計1":
            return str(int(round(self._get_work_numeric_val(row, "小計1"))))
        if col in WORK_CATEGORIES or col == "作業小計":
            val = self._get_work_numeric_val(row, col)
            unit = self.time_unit_var.get()
            if unit == "金額(円)":
                return str(int(round(val)))
            if unit == "時間(分÷60)":
                return f"{val:.2f}"
            if abs(val - round(val)) < 1e-6:
                return str(int(round(val)))
            return f"{val:.1f}"
        elif col in COST_YEN_COLUMNS:
            return str(int(round(row.get(col, 0.0))))
        return str(row.get(col, ""))

    def _row_matches_base_conditions(self, r: dict, exclude_col: str = None) -> bool:
        """全体検索・表示対象モード・および指定列以外の全列フィルターに一致するか高速判定"""
        w_sub = r["作業小計"]
        c_sub = r["購入品小計"]
        has_seizo = r.get("_has_seizo")
        if has_seizo is None:
            has_seizo = any(abs(r.get(c, 0.0)) > 1e-9 for c in SEIZO_COST_COLUMNS)
        mode = self.filter_mode_var.get()
        if mode == "実績ありのみ" and (w_sub <= 0 and c_sub <= 0 and not has_seizo):
            return False
        elif mode == "作業実績ありのみ" and w_sub <= 0:
            return False
        elif mode == "購入・外注実績ありのみ" and c_sub <= 0:
            return False
        elif mode == "作業＋購入・外注の両方あり" and (w_sub <= 0 or c_sub <= 0):
            return False

        kw = self.search_var.get().strip().lower()
        if kw:
            haystack = r.get("_search_text")
            if haystack is None:
                haystack = " ".join([
                    str(r.get("管理No", "")),
                    str(r.get("受注先", "")),
                    str(r.get("出荷先", "")),
                    str(r.get("プロジェクト名", "")),
                    str(r.get("品名、ライン名等", "")),
                    str(r.get("製品詳細", "")),
                    str(r.get("製造又は手配先", "")),
                    str(r.get("売り上げ区分", "")),
                    str(r.get("備考", "")),
                ]).lower()
                r["_search_text"] = haystack
            if kw not in haystack:
                return False

        for f_col, f_spec in self.column_filters.items():
            if f_col == exclude_col:
                continue
            if not self._row_matches_column_filter(r, f_col, f_spec):
                return False
        return True

    def _row_matches_column_filter(self, r: dict, col: str, spec: dict) -> bool:
        """単一列のExcel風フィルター条件判定"""
        sel = spec.get("selected")
        if sel is not None:
            key_str = self._get_cell_filter_key(r, col)
            if key_str not in sel:
                return False

        num_op = spec.get("num_op", "指定なし")
        if col in NUMERIC_COLUMNS and num_op != "指定なし":
            if col in WORK_CATEGORIES or col in ("作業小計", "小計1"):
                val = self._get_work_numeric_val(r, col)
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

        # 他の列の条件を満たす行からユニーク値と件数を集計（Excelと同じカスケード動作）
        counts = Counter()
        for r in self.aggregated_rows:
            if self._row_matches_base_conditions(r, exclude_col=col_name):
                counts[self._get_cell_filter_key(r, col_name)] += 1

        # すでに選択中のフィルター値が件数0になっていても復元できるよう含める
        cur_f = self.column_filters.get(col_name)
        if cur_f and cur_f.get("selected"):
            for v in cur_f["selected"]:
                if v not in counts:
                    counts[v] = 0

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
            desc = ", ".join(self.column_filters.keys())
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

        filtered = [r for r in self.aggregated_rows if self._row_matches_base_conditions(r)]
        self.displayed_rows = filtered

        unit = self.time_unit_var.get()
        is_yen_mode = (unit == "金額(円)")
        totals = {c: 0.0 for c in CALC_COLUMNS}
        rates_tuple = self._get_rates_tuple()
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
            totals["小計1"] += r["_yen_小計1"]

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
                f"【小計1: {int(round(totals['小計1'])):,}円】"
            )
        else:
            unit_label = "h" if unit == "時間(分÷60)" else "分"
            fmt = (lambda v: f"{v / 60.0:,.2f}") if unit == "時間(分÷60)" else self._format_converted_work_val
            self.summary_var.set(
                f"表示件数: {len(self.displayed_rows):,} 件   |   "
                f"作業時間: {fmt(totals['作業時間'])}{unit_label}   "
                f"移動時間: {fmt(totals['移動時間'])}{unit_label}   "
                f"修理改造時間: {fmt(totals['修理改造時間'])}{unit_label}   "
                f"設計関係: {fmt(totals['設計関係'])}{unit_label}   "
                f"その他作業: {fmt(totals['その他作業'])}{unit_label}   "
                f"【作業小計: {fmt(totals['作業小計'])}{unit_label}】   |   "
                f"【購入品小計: {int(round(totals['購入品小計'])):,}円】   |   "
                f"材料費計: {int(round(totals['材料費計'])):,}円   "
                f"設計費: {int(round(totals['設計費'])):,}円   |   "
                f"【小計1: {int(round(totals['小計1'])):,}円】"
            )

        self._refresh_both_panes_sorted()

    def _refresh_both_panes_sorted(self):
        """行リンクON時は1回だけソートして左右両ペインへ渡し、2重ソートを省いて高速化する"""
        if self.link_rows_var.get():
            col = self.sort_col
            desc = self.sort_desc
            if col in WORK_CATEGORIES or col in ("作業小計", "小計1"):
                get_w = self._get_work_numeric_val
                sort_key = lambda item: get_w(item, col)
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

        ColumnOrderDialog(self.root, self.column_order, self.visible_set, on_apply)

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
        for i, d in enumerate(d_list):
            mins = d.get("合計", 0.0)
            upd = d.get("更新日", "")
            if "." in upd:
                upd = upd.split(".")[0]
            cat = d.get("分類区分", "")
            if cat in self.work_rates:
                rate = to_float(self.work_rates.get(cat, DEFAULT_WORK_RATES.get(cat, 0.0)), 0.0)
                yen_str = f"{int(round(mins * rate)):,}"
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
        for i, r in enumerate(k_list):
            q_num = to_float(r.get("Q"), 0.0)
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
        for i, r in enumerate(g_list):
            q_num = to_float(r.get("Q"), 0.0)
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

    def open_code_settings_dialog(self):
        """作業コード範囲設定ダイアログを開く"""
        dlg = tk.Toplevel(self.root)
        dlg.title("作業コード範囲設定")
        dlg.geometry("500x360")
        dlg.resizable(False, False)
        dlg.configure(bg=COLORS["bg_card"])
        dlg.transient(self.root)
        dlg.grab_set()

        frame = ttk.Frame(dlg, style="Card.TFrame", padding=16)
        frame.pack(fill=tk.BOTH, expand=True)

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

        btn_f = ttk.Frame(frame, style="Flat.TFrame")
        btn_f.pack(side=tk.BOTTOM, fill=tk.X, pady=(14, 0))

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

    def _prepare_export_data(self):
        """現在表示中の列順番・アクティブペインの並び順・フィルター結果から出力データを生成する"""
        unit = self.time_unit_var.get()
        src_cols = list(self.visible_columns)
        headers = [get_display_column_label(c, unit) for c in src_cols]
        numeric_indices = {i for i, col in enumerate(src_cols) if col in NUMERIC_COLUMNS}
        pane = self._get_active_pane()
        export_source = pane.rows if pane.rows else self.displayed_rows
        rows = []
        is_yen_mode = (unit == "金額(円)")
        totals = {c: 0.0 for c in CALC_COLUMNS}

        for r in export_source:
            if is_yen_mode:
                row_sub_yen = 0
                for c in WORK_CATEGORIES:
                    yen_v = int(round(self._get_work_numeric_val(r, c)))
                    totals[c] += yen_v
                    row_sub_yen += yen_v
                totals["作業小計"] += row_sub_yen
            else:
                for c in WORK_CATEGORIES:
                    totals[c] += r[c]
                totals["作業小計"] += r["作業小計"]
            for c in COST_YEN_COLUMNS:
                totals[c] += r.get(c, 0.0)
            totals["小計1"] += self._get_work_numeric_val(r, "小計1")

            row_vals = []
            for col in src_cols:
                if col in WORK_CATEGORIES or col in ("作業小計", "小計1"):
                    row_vals.append(self._get_work_numeric_val(r, col))
                elif col in COST_YEN_COLUMNS:
                    row_vals.append(int(round(r.get(col, 0.0))))
                else:
                    row_vals.append(r.get(col, ""))
            rows.append(row_vals)

        totals_row = []
        for idx, col in enumerate(src_cols):
            if idx == 0:
                totals_row.append(f"合計 ({len(self.displayed_rows)}件)")
            elif col == "小計1":
                totals_row.append(int(round(totals["小計1"])))
            elif col in WORK_CATEGORIES or col == "作業小計":
                if is_yen_mode:
                    totals_row.append(int(round(totals[col])))
                elif unit == "時間(分÷60)":
                    totals_row.append(round(totals[col] / 60.0, 2))
                else:
                    totals_row.append(totals[col])
            elif col in COST_YEN_COLUMNS:
                totals_row.append(int(round(totals[col])))
            else:
                totals_row.append("")

        return headers, rows, numeric_indices, totals_row

    def open_excel_new_workbook_action(self):
        """表示中の一覧データを保存ファイルではなくExcelの新規ブック（未保存ブック）として開く"""
        if not self.displayed_rows:
            messagebox.showinfo("出力対象なし", "Excelへ出力するデータがありません。先に集計を実行してください。", parent=self.root)
            return

        headers, rows, numeric_indices, totals_row = self._prepare_export_data()
        protected_snapshot = self.get_protected_source_paths()
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

    def open_help_dialog(self):
        """見やすいガイド形式のヘルプ画面を表示する"""
        HelpDialog(self.root)

    def _on_close(self):
        self._closing = True
        self._active_job_id += 1
        self.meta_pane.cancel_pending_jobs()
        self.calc_pane.cancel_pending_jobs()
        if getattr(self, "_poll_after_id", None) is not None:
            try:
                self.root.after_cancel(self._poll_after_id)
            except Exception:
                pass
            self._poll_after_id = None
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
