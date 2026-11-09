# -*- coding: utf-8 -*-
"""
STag SKU 速查

两种触发方式：
  · 鼠标悬停：光标在编码上停留片刻，自动弹出完整 SKU 名称和款式
  · 选中 + 热键：选中编码后按 F8

支持把数据源放在共享目录，程序会定时检查更新，并保留本机 SQLite 查询库。
"""

import asyncio
import concurrent.futures
import ctypes
import gc
import io
import json
import logging
from logging.handlers import RotatingFileHandler
import math
import os
import queue
import re
import shutil
import sqlite3
import sys
import tempfile
import threading
import time
import unicodedata
import winreg

import tkinter as tk
from tkinter import filedialog, messagebox

import win32api
import win32clipboard
import win32con
import win32gui
import win32process

try:
    import pythoncom
    import win32com.client
except Exception:  # pragma: no cover - pywin32 通常自带
    pythoncom = None
    win32com = None

# ---------------------------------------------------------------- 路径与日志


def _app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR_NAME = "STag"
DATA_DIR_NAME = "data"


BASE_DIR = _app_dir()
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
USER_CONFIG_DIR = os.path.join(
    os.environ.get("APPDATA", BASE_DIR), APP_DIR_NAME
)
USER_SETTINGS_PATH = os.path.join(USER_CONFIG_DIR, "settings.json")

CACHE_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", tempfile.gettempdir()), APP_DIR_NAME
)
SKU_DB_PATH = os.path.join(CACHE_DIR, "sku.db")
OPS_DB_PATH = os.path.join(CACHE_DIR, "ops.db")
DB_VERSION = 1
# 两个数据源各自导入；串行构建，避免同时读两份大 Excel 把内存峰值抬高
DB_BUILD_LOCK = threading.Lock()
SHOW_SIGNAL_PATH = os.path.join(CACHE_DIR, "show_window.signal")

LOG_PATH = os.path.join(BASE_DIR, "sku_lookup.log")


def resolve_source_path(path):
    """优先使用配置路径；找不到时，回退到 exe 旁边的 data 目录。"""
    path = str(path or "").strip()
    if not path:
        return path
    if os.path.exists(path):
        return os.path.abspath(path)
    basename = os.path.basename(path)
    candidates = []
    if not os.path.isabs(path):
        candidates.append(os.path.join(BASE_DIR, path))
    candidates.extend(
        os.path.join(base, DATA_DIR_NAME, basename)
        for base in (BASE_DIR, os.path.dirname(BASE_DIR))
    )
    candidates.append(os.path.join(os.path.dirname(BASE_DIR), basename))
    for candidate in candidates:
        if os.path.exists(candidate):
            return os.path.abspath(candidate)

    # 旧配置里常见“负责人对照表 (2).xlsx”这类副本名；如果同目录下有
    # 去掉 (2) 后缀的同名文件，自动接上，避免用户被旧路径卡死。
    configured_stem = os.path.splitext(basename)[0]
    configured_norm = re.sub(r"\s*\(\d+\)$", "", configured_stem).strip().lower()
    for source_dir in (
        os.path.join(base, DATA_DIR_NAME)
        for base in (BASE_DIR, os.path.dirname(BASE_DIR))
    ):
        if not os.path.isdir(source_dir):
            continue
        for name in os.listdir(source_dir):
            if not name.lower().endswith((".xlsx", ".xlsm")):
                continue
            stem = os.path.splitext(name)[0]
            stem_norm = re.sub(r"\s*\(\d+\)$", "", stem).strip().lower()
            if stem_norm == configured_norm:
                return os.path.abspath(os.path.join(source_dir, name))
    return path


def _cleanup_legacy_cache():
    """删除旧版 pickle 缓存和上次中断留下的临时库。"""
    for name in ("sku_cache.pkl", "ops_cache.pkl"):
        try:
            os.remove(os.path.join(CACHE_DIR, name))
        except FileNotFoundError:
            pass
        except OSError:
            pass
    try:
        now = time.time()
        for name in os.listdir(CACHE_DIR):
            if not (name.startswith("sku_build_") and name.endswith(".db")):
                continue
            path = os.path.join(CACHE_DIR, name)
            try:
                if now - os.path.getmtime(path) > 600:
                    os.remove(path)
            except OSError:
                pass
    except OSError:
        pass


def _setup_logging():
    """日志最多 1MB × 2 份，悬停未命中只写 DEBUG，正常使用不会越滚越大。"""
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    for path in (
        LOG_PATH,
        os.path.join(CACHE_DIR, "sku_lookup.log"),
    ):
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            handler = RotatingFileHandler(
                path,
                maxBytes=1_000_000,
                backupCount=2,
                encoding="utf-8",
            )
            handler.setFormatter(formatter)
            logging.basicConfig(
                level=logging.INFO, handlers=[handler], force=True
            )
            return
        except OSError:
            continue
    logging.basicConfig(level=logging.INFO, force=True)


# ---------------------------------------------------------------- 视觉主题

COLOR_BG = "#05070D"
COLOR_PANEL = "#0A101C"
COLOR_PANEL_2 = "#0E1626"
COLOR_INPUT = "#070C16"
COLOR_LINE = "#1A2740"
COLOR_LINE_BRIGHT = "#27405E"
COLOR_TEXT = "#E8F3FF"
COLOR_MUTED = "#7C8DA6"
COLOR_CYAN = "#35E7FF"
COLOR_MAGENTA = "#FF3D9A"
COLOR_AMBER = "#FFB547"
COLOR_DIM_CYAN = "#2A7F93"
COLOR_BADGE_INK = "#041018"

FONT_UI = "Microsoft YaHei UI"
FONT_MONO = "Consolas"
FONT_DISPLAY = "Bahnschrift"


def resource_path(relative):
    base = getattr(sys, "_MEIPASS", BASE_DIR)
    return os.path.join(base, relative)


# ---------------------------------------------------------------- 配置


def default_config():
    return {
        "data_source": {
            "path": "data/inventory.xlsx",
            "sheet": 0,
            "code_column": "sku编码",
            "name_column": "sku名称",
            "stock_columns": ["库存A", "库存B"],
            "check_interval_seconds": 30,
            "use_local_cache": True,
        },
        "operations": {
            "path": "data/owners.xlsx",
            "sheet": 0,
            "code_column": "编码",
            "owner_column": "负责人",
            "assistant_column": "助理",
            "check_interval_seconds": 30,
        },
        "trigger": {
            # mode 仅用于兼容旧配置；实际开关看下面两个布尔值
            "mode": "both",
            "hover_enabled": True,
            "hotkey_enabled": True,
            "hover_dwell_ms": 600,
            "hotkey": "F8",
            "hover_region_width": 520,
            "hover_region_height": 48,
            "ocr_scale": 3,
        },
        "display": {
            "auto_close_seconds": 15,
            "font_size": 13,
            "opacity": 0.98,
            "show_stock": True,
            "show_module_name": False,
        },
        "system": {
            "start_with_windows": False,
        },
    }


def deep_merge(base, override):
    result = dict(base)
    for key, value in (override or {}).items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def trigger_mode(trigger):
    """把两个开关映射成旧版认识的 mode 字符串。"""
    hover = bool(trigger.get("hover_enabled", True))
    hotkey = bool(trigger.get("hotkey_enabled", True))
    if hover and hotkey:
        return "both"
    if hover:
        return "hover"
    if hotkey:
        return "hotkey"
    return "off"


def normalize_trigger(cfg):
    """兼容旧配置：没有 hover_enabled/hotkey_enabled 时从 mode 推导。"""
    trigger = cfg.setdefault("trigger", {})
    if "hover_enabled" not in trigger or "hotkey_enabled" not in trigger:
        mode = str(trigger.get("mode", "hover")).lower()
        if mode == "hotkey":
            trigger["hover_enabled"] = False
            trigger["hotkey_enabled"] = True
        elif mode == "off":
            trigger["hover_enabled"] = False
            trigger["hotkey_enabled"] = False
        elif mode == "both":
            trigger["hover_enabled"] = True
            trigger["hotkey_enabled"] = True
        else:  # hover，以及没有 mode 的老配置
            trigger["hover_enabled"] = True
            trigger["hotkey_enabled"] = False
    trigger["hover_enabled"] = bool(trigger["hover_enabled"])
    trigger["hotkey_enabled"] = bool(trigger["hotkey_enabled"])
    # 2 倍放大时 OCR 容易把 1.01.08.3072 的点和 0 丢掉，统一抬到 3 倍
    try:
        trigger["ocr_scale"] = max(3, int(trigger.get("ocr_scale", 3)))
    except (TypeError, ValueError):
        trigger["ocr_scale"] = 3
    trigger["mode"] = trigger_mode(trigger)
    return cfg


def migrate_config(raw):
    """兼容上一版扁平配置。"""
    if not isinstance(raw, dict) or "data_source" in raw:
        return raw or {}
    cfg = default_config()
    if "excel_path" in raw:
        cfg["data_source"]["path"] = raw["excel_path"]
    if "sheet" in raw:
        cfg["data_source"]["sheet"] = raw["sheet"]
    if "code_column" in raw:
        cfg["data_source"]["code_column"] = raw["code_column"]
    if "name_column" in raw:
        cfg["data_source"]["name_column"] = raw["name_column"]
    if "stock_columns" in raw:
        cfg["data_source"]["stock_columns"] = raw["stock_columns"]
    trigger = raw.get("hotkey_lookup")
    if trigger:
        cfg["trigger"]["hotkey"] = trigger
    if "auto_close_seconds" in raw:
        cfg["display"]["auto_close_seconds"] = raw["auto_close_seconds"]
    if "font_size" in raw:
        cfg["display"]["font_size"] = raw["font_size"]
    if "show_stock" in raw:
        cfg["display"]["show_stock"] = raw["show_stock"]
    return cfg


def _write_json(path, payload):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as handle:
        handle.write(payload)
    os.replace(tmp, path)


def load_config():
    """共享 config.json 管数据源；本机 settings.json 管个人偏好。"""
    cfg = default_config()

    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as handle:
                cfg = deep_merge(cfg, migrate_config(json.load(handle)))
        except Exception:
            logging.exception("读取共享配置失败")

    if os.path.exists(USER_SETTINGS_PATH):
        try:
            with open(USER_SETTINGS_PATH, "r", encoding="utf-8") as handle:
                user = json.load(handle)
            for section in ("trigger", "display", "system"):
                if isinstance(user.get(section), dict):
                    cfg[section] = deep_merge(cfg[section], user[section])
            # 共享盘只读时，数据源路径会退回到本机设置
            if isinstance(user.get("data_source"), dict):
                cfg["data_source"] = deep_merge(
                    cfg["data_source"], user["data_source"]
                )
            if isinstance(user.get("operations"), dict):
                cfg["operations"] = deep_merge(
                    cfg["operations"], user["operations"]
                )
        except Exception:
            logging.exception("读取本机设置失败")

    cfg["data_source"]["path"] = resolve_source_path(
        cfg["data_source"].get("path", "")
    )
    cfg["operations"]["path"] = resolve_source_path(
        cfg["operations"].get("path", "")
    )
    return normalize_trigger(cfg)


def save_config(cfg):
    """
    数据源写入程序目录的 config.json（可以在多台设备间共享）；
    触发方式、显示、启动等个人偏好写入本机 AppData。
    程序目录不可写时，数据源也一并写进本机设置。
    """
    shared_ok = False
    try:
        shared = {}
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH, "r", encoding="utf-8") as handle:
                shared = json.load(handle)
            if not isinstance(shared, dict):
                shared = {}
        shared["data_source"] = cfg["data_source"]
        shared["operations"] = cfg["operations"]
        _write_json(
            CONFIG_PATH, json.dumps(shared, ensure_ascii=False, indent=2)
        )
        shared_ok = True
    except (OSError, ValueError):
        logging.exception("共享配置不可写，数据源将保存在本机")

    user_payload = {key: cfg[key] for key in ("trigger", "display", "system")}
    if not shared_ok:
        user_payload["data_source"] = cfg["data_source"]
        user_payload["operations"] = cfg["operations"]

    try:
        _write_json(
            USER_SETTINGS_PATH,
            json.dumps(user_payload, ensure_ascii=False, indent=2),
        )
    except OSError:
        if not shared_ok:
            raise OSError("配置无法保存：程序目录和 AppData 都没有写入权限")

    return CONFIG_PATH if shared_ok else USER_SETTINGS_PATH


# ---------------------------------------------------------------- 编码与查表

SEP_CHARS = set(".,，．·、。;；:：|｜丨/\\-－—_∙⋅•")
RE_NINE = re.compile(r"(?<!\d)\d{9}(?!\d)")
RE_DOTTED = re.compile(r"(?<!\d)\d{1,3}(?:\.\d{1,4}){1,3}(?!\d)")
RE_CODE_IN_TEXT = re.compile(r"\d{1,3}(?:\.\d{1,4}){2,3}|\d{9}")


_INVISIBLE = dict.fromkeys(
    map(
        ord,
        "​‌‍‎‏﻿⁠­"
        "‪‫‬‭‮",
    )
)


def _clean_text(value):
    """NFKC + 去掉复制时常见的零宽/方向控制字符。"""
    text = unicodedata.normalize("NFKC", str(value or ""))
    return text.translate(_INVISIBLE)


def normalize_code(value):
    if value is None:
        return ""
    text = _clean_text(value).strip().lstrip("'")
    text = re.sub(r"\s+", "", text)
    # 把数字之间的各种分隔符统一成点：1,01、08-3072 / 1。01。08。3072
    text = re.sub(
        r"(?<=\d)[.,，．·、。;；:：|｜丨/\\-－—_∙⋅•](?=\d)", ".", text
    )
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".")[0]
    return text


def reconstruct_dotted(digits):
    """把 9 位数字还原成 1.01.01.0033 这种固定格式。"""
    if len(digits) == 9:
        return f"{digits[0]}.{digits[1:3]}.{digits[3:5]}.{digits[5:]}"
    return None


def lookup_code(token, mapping):
    """返回 (规范编码, 匹配惩罚分) 或 None。"""
    token = normalize_code(token)
    if not token:
        return None
    if token in mapping:
        return token, 0.0
    if "." in token:
        rebuilt = reconstruct_dotted(re.sub(r"\D", "", token))
        if rebuilt and rebuilt in mapping:
            return rebuilt, 12.0
    return None


def lookup_text(text, mapping):
    """从一段普通文字里提取编码并查表。

    钉钉/RPA 机器人复制出来的内容经常带 `^l`、`|`、项目符号、零宽字符，
    编码分隔符也可能是 `。`、`·`、`,`、`、`、`-` 等。这里统一先清洗，
    再交给和 OCR 相同的匹配器处理。
    """
    raw = _clean_text(text).strip()
    if not raw:
        return None

    direct = lookup_code(raw, mapping)
    if direct:
        code, _ = direct
        entry = mapping.get(code)
        if entry is not None:
            return code, entry

    # 按行处理，避免把上下两行的编码粘成一个长数字串
    for line in raw.splitlines():
        compact = re.sub(r"[ \t\r\f\v]+", "", line)
        if not compact:
            continue
        normalized = _normalize_compact(compact)
        for _start, _end, code, _penalty in _iter_code_matches(
            normalized, mapping
        ):
            entry = mapping.get(code)
            if entry is not None:
                return code, entry
        # 复制出来只剩空格分隔时，也还原成带点编码
        for match in re.finditer(
            r"(?<![\d.])\d{1,3}(?:\s+\d{1,4}){2,3}(?![\d.])", line
        ):
            hit = lookup_code(".".join(match.group(0).split()), mapping)
            if hit:
                code = hit[0]
                entry = mapping.get(code)
                if entry is not None:
                    return code, entry
    return None


def detect_styles(name):
    # 赛博朋克配色映射：无边 = 青色，丝印 = 品红，未知 = 灰
    compact = re.sub(r"\s+", "", name or "")
    # 业务规则：无边优先；没有“无边”的，只要出现“热弯”或“丝印”都按丝印
    if "无边" in compact:
        return [("无边款", COLOR_CYAN)]
    if "丝印" in compact or "热弯" in compact:
        return [("丝印款", COLOR_MAGENTA)]
    return [("未标注款式", COLOR_MUTED)]


def fmt_num(value):
    if value is None:
        return ""
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return ""
    try:
        num = float(text)
        return f"{int(num):,}" if num == int(num) else f"{num:,.2f}"
    except ValueError:
        return text


# ---------------------------------------------------------------- 数据加载


def _path_key(path):
    """统一路径写法，避免 H:/x 和 H:\\x 被当成两个数据源。"""
    try:
        return os.path.normcase(os.path.abspath(os.path.normpath(str(path))))
    except Exception:
        return str(path)


def _copy_remote_source(path, filename):
    """共享盘先复制到本地临时文件，避免 Excel 被占用时读失败。"""
    stat = None
    try:
        info = os.stat(path)
        stat = (info.st_mtime, info.st_size)
    except OSError:
        stat = None

    read_path = path
    if path.startswith("\\\\"):
        local_copy = os.path.join(tempfile.gettempdir(), filename)
        shutil.copy2(path, local_copy)
        read_path = local_copy
    return read_path, stat


def _select_sheet(workbook, sheet):
    if isinstance(sheet, int):
        return workbook.worksheets[sheet]
    text = str(sheet)
    if text.isdigit():
        return workbook.worksheets[int(text)]
    return workbook[text]


def _read_header(sheet):
    """读第一行非空表头，同时返回剩余行的迭代器。"""
    # read_only 模式下 openpyxl 有时把工作表维度记成 A1，导致只读到第一列。
    # 重置维度后按实际内容继续流式读取。
    reset = getattr(sheet, "reset_dimensions", None)
    if reset is not None:
        try:
            reset()
        except Exception:
            pass
    rows = sheet.iter_rows(values_only=True)
    for row in rows:
        header = ["" if cell is None else str(cell).strip() for cell in row]
        if any(header):
            return header, rows
    return [], rows


def _column_index(header, configured, keyword):
    configured = str(configured or "")
    if configured in header:
        return header.index(configured)
    for index, name in enumerate(header):
        if keyword and keyword in name:
            return index
    return None


def _cell(row, index):
    if index is None or index >= len(row):
        return ""
    value = row[index]
    return "" if value is None else value


def _import_openpyxl():
    """导入 openpyxl 时跳过 numpy。

    openpyxl 只在兼容层里尝试 import numpy；我们只读单元格值，不需要
    numpy 的数值类型。跳过它可以少占几百 MB 的提交内存。
    """
    saved = sys.modules.get("numpy")
    sys.modules["numpy"] = None
    try:
        from openpyxl import load_workbook

        return load_workbook
    finally:
        if saved is None:
            sys.modules.pop("numpy", None)
        else:
            sys.modules["numpy"] = saved


def _new_temp_db():
    os.makedirs(CACHE_DIR, exist_ok=True)
    handle, path = tempfile.mkstemp(
        prefix="sku_build_", suffix=".db", dir=CACHE_DIR
    )
    os.close(handle)
    return path


def _write_meta(conn, meta):
    conn.execute(
        "CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT NOT NULL)"
    )
    conn.execute(
        "INSERT OR REPLACE INTO meta(key, value) VALUES('meta', ?)",
        (json.dumps(meta, ensure_ascii=False),),
    )


def build_sku_db(path, cfg):
    """把 SKU 表流式导入临时 SQLite，返回 (临时库路径, meta)。

    openpyxl 的 read_only 模式逐行读取，不再把整张表转成 DataFrame /
    list[dict]；应用运行时只保留数据库连接，查询时按编码取一行。
    """
    load_workbook = _import_openpyxl()

    data_cfg = cfg["data_source"]
    read_path, stat = _copy_remote_source(path, "stag_sku_source.xlsx")
    temp_path = _new_temp_db()
    conn = None

    try:
        with DB_BUILD_LOCK:
            workbook = load_workbook(
                read_path, read_only=True, data_only=True, keep_links=False
            )
            try:
                sheet = _select_sheet(workbook, data_cfg.get("sheet", 0))
                header, rows = _read_header(sheet)

                code_idx = _column_index(
                    header, data_cfg.get("code_column", ""), "编码"
                )
                name_idx = _column_index(
                    header, data_cfg.get("name_column", ""), "名称"
                )
                if code_idx is None or name_idx is None:
                    raise ValueError(
                        f"Excel 里找不到列「{data_cfg.get('code_column')}」"
                        f"或「{data_cfg.get('name_column')}」"
                    )

                stock_cols = []
                for name in data_cfg.get("stock_columns", []):
                    index = _column_index(header, name, "")
                    if index is not None:
                        stock_cols.append((name, index))

                date_idx = None
                for index, name in enumerate(header):
                    if "统计时间" in name or "日期" in name:
                        date_idx = index
                        break

                conn = sqlite3.connect(temp_path)
                conn.execute("PRAGMA journal_mode = OFF")
                conn.execute("PRAGMA synchronous = OFF")
                conn.execute("PRAGMA temp_store = MEMORY")
                conn.execute(
                    "CREATE TABLE sku ("
                    "code TEXT PRIMARY KEY, "
                    "name TEXT NOT NULL, "
                    "stock_json TEXT NOT NULL"
                    ")"
                )

                batch = []
                data_date = ""
                for row in rows:
                    code = normalize_code(_cell(row, code_idx))
                    name = str(_cell(row, name_idx)).strip()
                    if not code or not name or name.lower() == "nan":
                        continue

                    stock = {}
                    for column_name, index in stock_cols:
                        formatted = fmt_num(_cell(row, index))
                        if formatted:
                            stock[column_name] = formatted

                    if date_idx is not None:
                        value = str(_cell(row, date_idx)).strip()
                        if value and value.lower() != "nan" and value > data_date:
                            data_date = value

                    batch.append(
                        (
                            code,
                            name,
                            json.dumps(stock, ensure_ascii=False),
                        )
                    )
                    if len(batch) >= 1000:
                        conn.executemany(
                            "INSERT OR IGNORE INTO sku(code, name, stock_json) "
                            "VALUES (?, ?, ?)",
                            batch,
                        )
                        batch.clear()

                if batch:
                    conn.executemany(
                        "INSERT OR IGNORE INTO sku(code, name, stock_json) "
                        "VALUES (?, ?, ?)",
                        batch,
                    )

                count = conn.execute("SELECT COUNT(*) FROM sku").fetchone()[0]
                if not count:
                    raise ValueError(
                        "Excel 读取成功，但没有得到有效的 编码 → 名称 数据"
                    )

                meta = {
                    "db_version": DB_VERSION,
                    "source": path,
                    "stat": stat,
                    "count": count,
                    "data_date": data_date,
                    "loaded_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                }
                _write_meta(conn, meta)
                conn.commit()
            finally:
                workbook.close()
                if conn is not None:
                    conn.close()
                    conn = None

        return temp_path, meta
    except Exception:
        if conn is not None:
            conn.close()
        try:
            os.remove(temp_path)
        except OSError:
            pass
        raise


class _SqliteStore:
    """数据库只读句柄；查询只取需要的行，不在进程里保留全量数据。"""

    def __init__(self, path):
        self.path = path
        self._lock = threading.RLock()
        self._conn = None
        self._meta = {}
        self._count = 0
        self._open()

    def _open(self):
        self._close_locked()
        if not os.path.exists(self.path):
            return
        conn = None
        try:
            conn = sqlite3.connect(
                self.path, timeout=2.0, check_same_thread=False
            )
            conn.row_factory = sqlite3.Row
            conn.execute("PRAGMA query_only = ON")
            # 限制页缓存，避免数据库引擎又把整库缓存进内存
            conn.execute("PRAGMA cache_size = -4000")
            conn.execute("PRAGMA busy_timeout = 1000")
            row = conn.execute(
                "SELECT value FROM meta WHERE key = 'meta'"
            ).fetchone()
            meta = json.loads(row["value"]) if row else {}
            if isinstance(meta.get("stat"), list):
                meta["stat"] = tuple(meta["stat"])
            if int(meta.get("db_version") or 0) != DB_VERSION:
                conn.close()
                return
            self._conn = conn
            self._meta = meta
            self._count = int(meta.get("count") or 0)
        except Exception:
            logging.exception("打开本地数据库失败：%s", self.path)
            if conn is not None:
                try:
                    conn.close()
                except Exception:
                    pass

    def _close_locked(self):
        if self._conn is not None:
            try:
                self._conn.close()
            except Exception:
                pass
        self._conn = None
        self._meta = {}
        self._count = 0

    def close(self):
        with self._lock:
            self._close_locked()

    @property
    def meta(self):
        with self._lock:
            return dict(self._meta)

    @property
    def source(self):
        return self.meta.get("source")

    @property
    def count(self):
        with self._lock:
            return self._count

    def __len__(self):
        return self.count

    def __bool__(self):
        with self._lock:
            return self._count > 0 and self._conn is not None

    def matches_source(self, path):
        return _path_key(self.source) == _path_key(path)

    def _fetchone(self, sql, params):
        with self._lock:
            if self._conn is None:
                return None
            try:
                return self._conn.execute(sql, params).fetchone()
            except Exception:
                logging.debug("数据库查询失败", exc_info=True)
                return None

    def replace(self, temp_path):
        """用构建好的临时库原子替换当前库；替换期间锁住查询。"""
        with self._lock:
            self._close_locked()
            last_error = None
            for _attempt in range(3):
                try:
                    os.replace(temp_path, self.path)
                    last_error = None
                    break
                except OSError as exc:
                    last_error = exc
                    time.sleep(0.15)
            self._open()
            if last_error is not None:
                raise last_error

    def clear(self):
        with self._lock:
            self._close_locked()
            try:
                os.remove(self.path)
            except FileNotFoundError:
                pass
            except OSError:
                logging.exception("删除本地数据库失败：%s", self.path)


class SkuDatabase(_SqliteStore):
    def __init__(self, path=SKU_DB_PATH):
        super().__init__(path)

    def lookup(self, code):
        code = normalize_code(code)
        if not code:
            return None
        row = self._fetchone(
            "SELECT name, stock_json FROM sku WHERE code = ? LIMIT 1", (code,)
        )
        if row is None:
            return None
        try:
            stock = json.loads(row["stock_json"] or "{}")
        except ValueError:
            stock = {}
        return {"name": row["name"], "stock": stock}

    def __contains__(self, code):
        code = normalize_code(code)
        if not code:
            return False
        return (
            self._fetchone("SELECT 1 FROM sku WHERE code = ? LIMIT 1", (code,))
            is not None
        )

    def __getitem__(self, code):
        entry = self.lookup(code)
        if entry is None:
            raise KeyError(code)
        return entry

    def get(self, code, default=None):
        entry = self.lookup(code)
        return default if entry is None else entry

    def first_code(self):
        row = self._fetchone("SELECT code FROM sku LIMIT 1", ())
        return row["code"] if row else None


def _clean_person(value):
    text = str(value or "").strip()
    if not text or text.lower() == "nan" or text == "无":
        return ""
    return text


def build_ops_db(path, ops_cfg):
    """把负责人关系表流式导入临时 SQLite，返回 (临时库路径, meta)。"""
    load_workbook = _import_openpyxl()

    read_path, stat = _copy_remote_source(path, "stag_ops_source.xlsx")
    temp_path = _new_temp_db()
    conn = None

    try:
        with DB_BUILD_LOCK:
            workbook = load_workbook(
                read_path, read_only=True, data_only=True, keep_links=False
            )
            try:
                sheet = _select_sheet(workbook, ops_cfg.get("sheet", 0))
                header, rows = _read_header(sheet)

                code_idx = _column_index(
                    header, ops_cfg.get("code_column", ""), "编码"
                )
                owner_idx = _column_index(
                    header, ops_cfg.get("owner_column", ""), "负责人"
                )
                assistant_idx = _column_index(
                    header, ops_cfg.get("assistant_column", ""), "助理"
                )
                if code_idx is None or owner_idx is None:
                    raise ValueError("负责人关系表里找不到「编码」或「负责人」列")

                conn = sqlite3.connect(temp_path)
                conn.execute("PRAGMA journal_mode = OFF")
                conn.execute("PRAGMA synchronous = OFF")
                conn.execute("PRAGMA temp_store = MEMORY")
                conn.execute(
                    "CREATE TABLE ops ("
                    "code TEXT NOT NULL, "
                    "role TEXT NOT NULL, "
                    "name TEXT NOT NULL, "
                    "PRIMARY KEY (code, role, name)"
                    ")"
                )

                batch = []
                for row in rows:
                    code = normalize_code(_cell(row, code_idx))
                    if not code:
                        continue
                    owner = _clean_person(_cell(row, owner_idx))
                    assistant = (
                        _clean_person(_cell(row, assistant_idx))
                        if assistant_idx is not None
                        else ""
                    )
                    if owner:
                        batch.append((code, "owner", owner))
                    if assistant:
                        batch.append((code, "assistant", assistant))
                    if len(batch) >= 1000:
                        conn.executemany(
                            "INSERT OR IGNORE INTO ops(code, role, name) "
                            "VALUES (?, ?, ?)",
                            batch,
                        )
                        batch.clear()

                if batch:
                    conn.executemany(
                        "INSERT OR IGNORE INTO ops(code, role, name) "
                        "VALUES (?, ?, ?)",
                        batch,
                    )

                count = conn.execute(
                    "SELECT COUNT(DISTINCT code) FROM ops"
                ).fetchone()[0]
                meta = {
                    "db_version": DB_VERSION,
                    "source": path,
                    "stat": stat,
                    "count": count,
                    "loaded_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                }
                _write_meta(conn, meta)
                conn.commit()
            finally:
                workbook.close()
                if conn is not None:
                    conn.close()
                    conn = None

        return temp_path, meta
    except Exception:
        if conn is not None:
            conn.close()
        try:
            os.remove(temp_path)
        except OSError:
            pass
        raise


class OpsDatabase(_SqliteStore):
    def __init__(self, path=OPS_DB_PATH):
        super().__init__(path)

    def lookup(self, code):
        code = normalize_code(code)
        if not code:
            return None
        with self._lock:
            if self._conn is None:
                return None
            try:
                rows = self._conn.execute(
                    "SELECT role, name FROM ops WHERE code = ? ORDER BY rowid",
                    (code,),
                ).fetchall()
            except Exception:
                logging.debug("负责人关系查询失败", exc_info=True)
                return None
        owner = []
        assistant = []
        for row in rows:
            if row["role"] == "owner":
                owner.append(row["name"])
            else:
                assistant.append(row["name"])
        if not owner and not assistant:
            return None
        return {"owner": owner, "assistant": assistant}


class OperationsStore:
    def __init__(self, cfg, status_callback=None):
        self.cfg = cfg or {}
        self.status_callback = status_callback
        self._ops = OpsDatabase()
        self._lock = threading.Lock()
        self._stop = threading.Event()
        self._loading = False

    @property
    def ops(self):
        return self._ops

    def lookup(self, code):
        return self._ops.lookup(code)

    def _status(self, text):
        try:
            if self.status_callback:
                self.status_callback(text)
        except Exception:
            pass

    def _source_stat(self):
        try:
            info = os.stat(self.cfg["path"])
            return (info.st_mtime, info.st_size)
        except OSError:
            return None

    def start(self):
        threading.Thread(target=self._boot, daemon=True).start()
        threading.Thread(target=self._watch, daemon=True).start()

    def stop(self):
        self._stop.set()

    def _boot(self):
        path = self.cfg.get("path", "")
        db = self._ops
        if db and not db.matches_source(path):
            db.clear()
        if path and db.matches_source(path) and db:
            self._status(f"负责人关系已从本地数据库加载 {db.count:,} 条")
        stat = self._source_stat()
        if stat is None:
            if not db:
                self._status("负责人关系表不可用")
            return
        if db and db.matches_source(path) and db.meta.get("stat") == stat:
            self._status(f"负责人关系已加载 {db.count:,} 条")
            return
        self.reload()

    def _watch(self):
        while not self._stop.is_set():
            interval = max(
                10, int(self.cfg.get("check_interval_seconds", 30))
            )
            if self._stop.wait(interval):
                break
            if self._loading:
                continue
            stat = self._source_stat()
            known = self._ops.meta.get("stat")
            if stat is not None and stat != known:
                self.reload()
            elif stat is None and not self._ops:
                self.reload()

    def reload(self):
        if self._loading:
            return
        self._loading = True
        threading.Thread(target=self._load_worker, daemon=True).start()

    def _load_worker(self):
        try:
            temp_path, _meta = build_ops_db(self.cfg["path"], self.cfg)
            try:
                self._ops.replace(temp_path)
            except Exception:
                try:
                    os.remove(temp_path)
                except OSError:
                    pass
                raise
            if not self._ops:
                raise RuntimeError("负责人关系数据库重建后无法打开")
            gc.collect()
            self._status(f"负责人关系已加载 {self._ops.count:,} 条")
            logging.info("负责人关系加载完成：%s 条", self._ops.count)
        except Exception:
            logging.exception("读取负责人关系失败")
            self._status("负责人关系读取失败")
        finally:
            self._loading = False

    def update_config(self, cfg):
        old_path = self.cfg.get("path")
        self.cfg = cfg.get("operations", {}) or {}
        new_path = self.cfg.get("path")
        if old_path != new_path:
            self._ops.clear()
        self.reload()


class DataStore:
    def __init__(self, cfg, status_callback):
        self.cfg = cfg
        self.status_callback = status_callback
        self._mapping = SkuDatabase()
        self._lock = threading.Lock()
        self._loading = False
        self._stop = threading.Event()

    @property
    def mapping(self):
        return self._mapping

    @property
    def meta(self):
        return self._mapping.meta

    def _status(self, text, level="info"):
        try:
            self.status_callback(text, level)
        except Exception:
            logging.debug("状态回调失败", exc_info=True)

    def start(self):
        threading.Thread(target=self._boot, daemon=True).start()
        threading.Thread(target=self._watch, daemon=True).start()

    def stop(self):
        self._stop.set()

    def _boot(self):
        path = self.cfg["data_source"]["path"]
        db = self._mapping
        if db and not db.matches_source(path):
            db.clear()

        if (
            self.cfg["data_source"].get("use_local_cache", True)
            and db.matches_source(path)
            and db
        ):
            self._status(
                f"已从本地数据库加载 {db.count:,} 条 · 正在检查共享盘…",
                "info",
            )

        stat = self._source_stat()
        if stat is None:
            if db:
                self._status(
                    f"共享盘暂时不可用，使用本地数据库 {db.count:,} 条",
                    "warn",
                )
            else:
                self._status("数据源不可用，请在设置里检查路径", "error")
            return

        if db and db.matches_source(path) and db.meta.get("stat") == stat:
            self._status(
                f"已加载 {db.count:,} 条 · 数据源无变化", "ok"
            )
            return

        self.reload(silent=True)

    def _source_stat(self):
        try:
            info = os.stat(self.cfg["data_source"]["path"])
            return (info.st_mtime, info.st_size)
        except OSError:
            return None

    def reload(self, silent=False):
        if self._loading:
            return
        self._loading = True
        self._status("正在读取数据源…", "info")
        threading.Thread(target=self._load_worker, args=(silent,), daemon=True).start()

    def _load_worker(self, silent):
        try:
            temp_path, _meta = build_sku_db(
                self.cfg["data_source"]["path"], self.cfg
            )
            try:
                self._mapping.replace(temp_path)
            except Exception:
                try:
                    os.remove(temp_path)
                except OSError:
                    pass
                raise
            if not self._mapping:
                raise RuntimeError("SKU 数据库重建后无法打开")
            gc.collect()
            meta = self._mapping.meta
            suffix = (
                f" · 数据时间 {meta['data_date']}"
                if meta.get("data_date")
                else ""
            )
            self._status(f"已加载 {self._mapping.count:,} 条{suffix}", "ok")
            logging.info(
                "加载完成：%s 条，来源 %s",
                self._mapping.count,
                meta.get("source"),
            )
        except Exception as exc:
            logging.exception("读取数据源失败")
            self._status(f"读取失败：{exc}", "error")
        finally:
            self._loading = False

    def _watch(self):
        while not self._stop.is_set():
            interval = max(5, int(self.cfg["data_source"].get("check_interval_seconds", 30)))
            if self._stop.wait(interval):
                break
            if self._loading:
                continue
            stat = self._source_stat()
            known = self._mapping.meta.get("stat")
            if stat is not None and stat != known:
                logging.info("检测到数据源变化，自动重载")
                self.reload(silent=True)
            elif stat is None and not self._mapping:
                self.reload(silent=True)

    def update_config(self, cfg):
        old_path = self.cfg["data_source"]["path"]
        self.cfg = cfg
        new_path = cfg["data_source"]["path"]
        if old_path != new_path:
            self._mapping.clear()
            self.reload(silent=True)


# ---------------------------------------------------------------- 剪贴板


def _html_to_text(raw):
    start = raw.find("<!--StartFragment-->")
    if start >= 0:
        raw = raw[start + len("<!--StartFragment-->"):]
    end = raw.find("<!--EndFragment-->")
    if end >= 0:
        raw = raw[:end]
    text = re.sub(r"(?is)<(script|style).*?</\1>", " ", raw)
    text = re.sub(r"(?s)<[^>]+>", " ", text)
    try:
        import html as _html

        text = _html.unescape(text)
    except Exception:
        pass
    return re.sub(r"\s+", " ", text).strip()


def get_clipboard_text():
    try:
        win32clipboard.OpenClipboard()
        try:
            if win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                text = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
                if text and str(text).strip():
                    return str(text)

            if win32clipboard.IsClipboardFormatAvailable(win32con.CF_TEXT):
                data = win32clipboard.GetClipboardData(win32con.CF_TEXT)
                if isinstance(data, bytes):
                    for encoding in ("utf-8", "gbk", "mbcs"):
                        try:
                            text = data.decode(encoding)
                            if text.strip():
                                return text
                        except (UnicodeDecodeError, LookupError):
                            continue
                elif data and str(data).strip():
                    return str(data)

            html_format = win32clipboard.RegisterClipboardFormat("HTML Format")
            if win32clipboard.IsClipboardFormatAvailable(html_format):
                data = win32clipboard.GetClipboardData(html_format)
                if isinstance(data, bytes):
                    data = data.decode("utf-8", errors="ignore")
                text = _html_to_text(str(data or ""))
                if text:
                    return text
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        logging.debug("读取剪贴板失败", exc_info=True)
    return ""


def set_clipboard_text(text):
    try:
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(text, win32con.CF_UNICODETEXT)
        finally:
            win32clipboard.CloseClipboard()
        return True
    except Exception:
        logging.debug("写回剪贴板失败", exc_info=True)
        return False


def copy_selection():
    old = get_clipboard_text()
    try:
        before_seq = ctypes.windll.user32.GetClipboardSequenceNumber()
    except Exception:
        before_seq = None

    win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
    win32api.keybd_event(0x43, 0, 0, 0)
    time.sleep(0.03)
    win32api.keybd_event(0x43, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)

    new = ""
    deadline = time.time() + 1.2
    while time.time() < deadline:
        current = get_clipboard_text()
        changed = False
        try:
            changed = (
                before_seq is not None
                and ctypes.windll.user32.GetClipboardSequenceNumber() != before_seq
            )
        except Exception:
            changed = False
        if current and (changed or current != old):
            new = current
            break
        time.sleep(0.04)
    return new, old


# ---------------------------------------------------------------- 热键


VK_NAMES = {
    "f1": 0x70, "f2": 0x71, "f3": 0x72, "f4": 0x73, "f5": 0x74, "f6": 0x75,
    "f7": 0x76, "f8": 0x77, "f9": 0x78, "f10": 0x79, "f11": 0x7A, "f12": 0x7B,
    "esc": 0x1B, "enter": 0x0D, "space": 0x20, "tab": 0x09,
    "insert": 0x2D, "delete": 0x2E, "home": 0x24, "end": 0x23,
}


def parse_hotkey(combo):
    mods = set()
    vk = None
    for part in (combo or "").replace(" ", "").split("+"):
        key = part.lower()
        if key in ("ctrl", "control"):
            mods.add(win32con.VK_CONTROL)
        elif key == "alt":
            mods.add(win32con.VK_MENU)
        elif key == "shift":
            mods.add(win32con.VK_SHIFT)
        elif key in ("win", "windows"):
            mods.add(win32con.VK_LWIN)
        elif key in VK_NAMES:
            vk = VK_NAMES[key]
        elif len(key) == 1 and key.isalpha():
            vk = ord(key.upper())
        elif len(key) == 1 and key.isdigit():
            vk = ord(key)
        else:
            raise ValueError(f"不认识的按键：{part}")
    if vk is None:
        raise ValueError(f"热键缺少主键：{combo}")
    return mods, vk


# ---------------------------------------------------------------- OCR


class OcrUnavailable(Exception):
    pass


class OcrEnginePool:
    def __init__(self):
        self._local = threading.local()
        self.available = False
        self.languages = []
        if OcrEngine is None:
            return
        try:
            self.languages = [
                lang.language_tag for lang in OcrEngine.available_recognizer_languages
            ]
            self.available = bool(self.languages)
        except Exception:
            logging.exception("检查 OCR 语言包失败")

    def _create_engine(self):
        tags = self.languages
        for tag in ("zh-Hans-CN", "zh-Hans", "en-US", "en-GB"):
            if tag in tags:
                engine = OcrEngine.try_create_from_language(Language(tag))
                if engine is not None:
                    return engine
        for tag in tags:
            engine = OcrEngine.try_create_from_language(Language(tag))
            if engine is not None:
                return engine
        engine = OcrEngine.try_create_from_user_profile_languages()
        if engine is not None:
            return engine
        raise OcrUnavailable("系统没有可用的 OCR 语言包")

    def _engine(self):
        engine = getattr(self._local, "engine", None)
        if engine is None:
            engine = self._create_engine()
            self._local.engine = engine
        return engine

    def recognize(self, image, scale=2):
        if not self.available:
            raise OcrUnavailable("系统没有可用的 OCR 语言包")

        width, height = image.size
        image = image.resize((width * scale, height * scale), Image.LANCZOS)
        if image.mode != "RGB":
            image = image.convert("RGB")

        buffer = io.BytesIO()
        image.save(buffer, format="PNG")
        data = buffer.getvalue()

        async def run():
            stream = InMemoryRandomAccessStream()
            writer = DataWriter(stream)
            writer.write_bytes(data)
            await writer.store_async()
            await writer.flush_async()
            writer.detach_stream()
            stream.seek(0)

            decoder = await BitmapDecoder.create_async(stream)
            bitmap = await decoder.get_software_bitmap_async(
                BitmapPixelFormat.BGRA8, BitmapAlphaMode.PREMULTIPLIED
            )
            return await self._engine().recognize_async(bitmap)

        return asyncio.run(run())


def capture_around(cx, cy, width, height):
    user32 = ctypes.windll.user32
    vx = user32.GetSystemMetrics(76)  # SM_XVIRTUALSCREEN
    vy = user32.GetSystemMetrics(77)  # SM_YVIRTUALSCREEN
    vw = user32.GetSystemMetrics(78)  # SM_CXVIRTUALSCREEN
    vh = user32.GetSystemMetrics(79)  # SM_CYVIRTUALSCREEN

    width = min(width, vw)
    height = min(height, vh)

    left = max(vx, min(cx - width // 2, vx + vw - width))
    top = max(vy, min(cy - height // 2, vy + vh - height))
    right = left + width
    bottom = top + height

    image = ImageGrab.grab(bbox=(left, top, right, bottom), all_screens=True)
    return image, left, top


def _normalize_compact(compact):
    chars = list(compact)
    for index, char in enumerate(chars):
        if (
            char in SEP_CHARS
            and 0 < index < len(chars) - 1
            and chars[index - 1].isdigit()
            and chars[index + 1].isdigit()
        ):
            chars[index] = "."
    return "".join(chars)


def _line_spans(line):
    """返回 (拼好的文本, 压缩文本, 压缩索引→原文索引, 词区间列表)。"""
    words = []
    for word in line.words:
        text = unicodedata.normalize("NFKC", word.text or "")
        words.append((word, text))

    text = ""
    spans = []
    for index, (word, word_text) in enumerate(words):
        if index:
            text += " "
        start = len(text)
        text += word_text
        spans.append((start, len(text), word))

    compact_chars = []
    compact_orig = []
    for index, char in enumerate(text):
        if not char.isspace():
            compact_chars.append(char)
            compact_orig.append(index)

    return text, "".join(compact_chars), compact_orig, spans


_DIGIT_CONFUSION = str.maketrans(
    {
        "O": "0", "o": "0",
        "I": "1", "l": "1", "i": "1", "|": "1",
        "S": "5", "s": "5",
        "B": "8", "b": "6",
        "Z": "2", "z": "2",
        "G": "6", "q": "9", "g": "9",
    }
)

_CONFUSION_RUN = re.compile(r"[0-9OoIlSBZzbGqg|]{9,16}")


def _iter_code_matches(normalized, mapping):
    """从 OCR 行里找出所有能查表的编码。

    除了标准的 9 位码 / 带点码，还会在一个更长的数字串里滑动 9 位窗口。
    有些后台界面会把编码、角标、序号连在一起被 OCR 读成一串，例如
    「101077933 0」会压成 10 位；只要其中某 9 位命中数据源，就认它。
    小字号下 OCR 也常把 5 读成 S、0 读成 O，这里一并做常见字形纠正。
    """
    seen = set()

    for regex in (RE_NINE, RE_DOTTED):
        for match in regex.finditer(normalized):
            found = lookup_code(match.group(), mapping)
            if not found:
                continue
            code, penalty = found
            key = (match.start(), match.end(), code)
            if key in seen:
                continue
            seen.add(key)
            yield match.start(), match.end(), code, penalty

    for run in re.finditer(r"\d{10,16}", normalized):
        digits = run.group()
        for offset in range(0, len(digits) - 8):
            window = digits[offset : offset + 9]
            if window not in mapping:
                continue
            start = run.start() + offset
            key = (start, start + 9, window)
            if key in seen:
                continue
            seen.add(key)
            yield start, start + 9, window, 24.0

    # OCR 字形混淆：S→5、O→0、I/l→1、B→8 等
    for run in _CONFUSION_RUN.finditer(normalized):
        raw = run.group()
        if not any(char.isdigit() for char in raw):
            continue
        digits = raw.translate(_DIGIT_CONFUSION)
        if not digits.isdigit():
            continue
        for offset in range(0, len(digits) - 8):
            window = digits[offset : offset + 9]
            if window not in mapping:
                continue
            start = run.start() + offset
            key = (start, start + 9, window)
            if key in seen:
                continue
            seen.add(key)
            yield start, start + 9, window, 30.0

    # 带点编码里的字形混淆，例如 l.01.08.3072、1.O1.08.3O72
    confused = normalized.translate(_DIGIT_CONFUSION)
    if confused != normalized:
        for regex in (RE_NINE, RE_DOTTED):
            for match in regex.finditer(confused):
                found = lookup_code(match.group(), mapping)
                if not found:
                    continue
                code, penalty = found
                key = (match.start(), match.end(), code)
                if key in seen:
                    continue
                seen.add(key)
                yield match.start(), match.end(), code, penalty + 30.0


def extract_best_hit(result, region_left, region_top, scale, mapping, cursor, max_distance):
    """在 OCR 结果里找离光标最近、且能查到表的编码。"""
    hits = []
    cursor_x, cursor_y = cursor

    for line in result.lines:
        text, compact, compact_orig, spans = _line_spans(line)
        if not compact:
            continue

        normalized = _normalize_compact(compact)
        for start, end, code, penalty in _iter_code_matches(normalized, mapping):
            orig_start = compact_orig[start]
            orig_end = compact_orig[end - 1]
            rects = [
                word.bounding_rect
                for word_start, word_end, word in spans
                if word_start <= orig_end and word_end >= orig_start
            ]

            if rects:
                x1 = min(rect.x for rect in rects) / scale + region_left
                y1 = min(rect.y for rect in rects) / scale + region_top
                x2 = max(rect.x + rect.width for rect in rects) / scale + region_left
                y2 = max(rect.y + rect.height for rect in rects) / scale + region_top
            else:
                x1 = region_left
                y1 = region_top
                x2 = region_left + 1
                y2 = region_top + 1

            dx = max(x1 - cursor_x, 0, cursor_x - x2)
            dy = max(y1 - cursor_y, 0, cursor_y - y2)
            distance = math.hypot(dx, dy)
            if distance > max_distance:
                continue

            entry = mapping.get(code)
            if entry is None:
                continue
            hits.append(
                {
                    "code": code,
                    "entry": entry,
                    "score": distance + penalty,
                    "distance": distance,
                    "bbox": (x1, y1, x2, y2),
                }
            )

        # OCR 有时把点读丢，只留下「1 01 08 3072」这样的数字词组；
        # 用词之间的空格还原成 1.01.08.3072 再查一次。
        for match in re.finditer(
            r"(?<![\d.])\d{1,3}(?:\s+\d{1,4}){2,3}(?![\d.])", text
        ):
            candidate = ".".join(match.group(0).split())
            found = lookup_code(candidate, mapping)
            if not found:
                continue
            code, penalty = found
            rects = [
                word.bounding_rect
                for word_start, word_end, word in spans
                if word_start < match.end() and word_end > match.start()
            ]
            if not rects:
                continue
            x1 = min(rect.x for rect in rects) / scale + region_left
            y1 = min(rect.y for rect in rects) / scale + region_top
            x2 = max(rect.x + rect.width for rect in rects) / scale + region_left
            y2 = max(rect.y + rect.height for rect in rects) / scale + region_top

            dx = max(x1 - cursor_x, 0, cursor_x - x2)
            dy = max(y1 - cursor_y, 0, cursor_y - y2)
            distance = math.hypot(dx, dy)
            if distance > max_distance:
                continue
            entry = mapping.get(code)
            if entry is None:
                continue
            hits.append(
                {
                    "code": code,
                    "entry": entry,
                    "score": distance + penalty + 18,
                    "distance": distance,
                    "bbox": (x1, y1, x2, y2),
                }
            )

    if not hits:
        return None
    hits.sort(key=lambda item: item["score"])
    return hits[0]


# ---------------------------------------------------------------- 弹窗


_MODEL_HINTS = (
    "苹果", "iPhone", "iPad", "MacBook", "AirPods", "Apple Watch",
    "华为", "荣耀", "小米", "红米", "OPPO", "vivo", "iQOO", "realme",
    "三星", "魅族", "一加", "努比亚", "红魔", "摩托罗拉", "联想",
    "戴尔", "惠普", "华硕", "宏碁", "微软", "Surface", "索尼",
    "任天堂", "Switch", "大疆", "DJI", "安卓", "适用", "通用",
    "Google", "Pixel", "Nothing", "ROG", "ThinkPad", "拯救者",
)


def _looks_like_model(prefix):
    """判断第一个 '-' 前面那段是不是机型，而不是“彩盒/赠品”之类的分类。"""
    text = (prefix or "").strip()
    if not text or len(text) > 36:
        return False
    if any(char.isdigit() for char in text) or " " in text:
        return True
    lowered = text.lower()
    return any(hint.lower() in lowered for hint in _MODEL_HINTS)


def split_sku_name(name):
    """把 SKU 名称拆成（机型, 描述）；拆不出来就原样当一行。

    数据里的名称大多是「苹果 17 Pro Max/18 Pro Max-无尘舱钢化膜-...」，
    但也有一些是「彩盒-亚马逊-...」这种包装分类开头，不能一律按第一个
    连字符切，否则会把“彩盒”当成机型显示。
    """
    text = (name or "").strip()
    match = re.search(r"[-－—]", text)
    if not match:
        return text, ""
    model = text[: match.start()].strip()
    desc = text[match.start() + 1 :].strip()
    if model and desc and _looks_like_model(model):
        return model, desc
    return text, ""


class Popup:
    """HUD 风格的查码弹窗：四角瞄准框 + 品红脊线。"""

    BG = COLOR_INPUT

    def __init__(self, root, cfg):
        self.root = root
        self.cfg = cfg
        self.win = None
        self.after_id = None
        self.anchor = None
        self.batch_callback = None

    def close(self):
        if self.after_id:
            try:
                self.root.after_cancel(self.after_id)
            except Exception:
                pass
            self.after_id = None
        if self.win is not None:
            try:
                self.win.destroy()
            except Exception:
                pass
            self.win = None
        self.anchor = None

    @staticmethod
    def _draw_ticks(win):
        color = COLOR_DIM_CYAN
        length = 12
        ticks = [
            (0, 0, 1, 1, length, 1),
            (0, 0, 1, 1, 1, length),
            (1, 0, -length - 1, 1, length, 1),
            (1, 0, -1, 1, 1, length),
            (0, 1, 1, -2, length, 1),
            (0, 1, 1, -length - 1, 1, length),
            (1, 1, -length - 1, -2, length, 1),
            (1, 1, -1, -length - 1, 1, length),
        ]
        for rel_x, rel_y, x, y, width, height in ticks:
            tk.Frame(win, bg=color, width=width, height=height).place(
                relx=rel_x, rely=rel_y, x=x, y=y
            )

    def show(self, title, lines, tags=None, anchor=None, split_name=False, ops=None, module_tags=None):
        self.close()

        size = int(self.cfg["display"].get("font_size", 13))
        win = tk.Toplevel(self.root)
        self.win = win
        win.configure(bg=self.BG)
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        try:
            opacity = float(self.cfg["display"].get("opacity", 0.98))
            win.attributes("-alpha", max(0.3, min(1.0, opacity)))
        except Exception:
            pass

        frame = tk.Frame(
            win,
            bg=self.BG,
            highlightbackground=COLOR_DIM_CYAN,
            highlightthickness=1,
        )
        frame.pack(fill="both", expand=True)

        # 左侧品红脊线：赛博朋克信号色
        tk.Frame(frame, bg=COLOR_MAGENTA, width=2).pack(side="left", fill="y")

        inner = tk.Frame(frame, bg=self.BG)
        inner.pack(side="left", fill="both", expand=True, padx=14, pady=10)

        top = tk.Frame(inner, bg=self.BG)
        top.pack(fill="x")

        tk.Label(
            top,
            text=title,
            bg=self.BG,
            fg=COLOR_CYAN,
            font=(FONT_MONO, size, "bold"),
        ).pack(side="left")

        for label, color in tags or []:
            tk.Label(
                top,
                text=label,
                bg=color,
                fg=COLOR_BADGE_INK,
                font=(FONT_UI, max(9, size - 3), "bold"),
                padx=7,
                pady=2,
            ).pack(side="left", padx=(10, 0))

        if getattr(self, "batch_callback", None):
            batch_label = tk.Label(
                top,
                text="批量",
                bg=self.BG,
                fg=COLOR_CYAN,
                font=(FONT_UI, max(9, size - 4), "bold"),
                cursor="hand2",
                padx=4,
            )
            batch_label.pack(side="right", padx=(0, 10))

            def open_batch(_event):
                callback = self.batch_callback
                self.close()
                if callback:
                    callback()
                return "break"

            batch_label.bind("<Button-1>", open_batch)

        tk.Label(
            top,
            text="SKU // MATCH",
            bg=self.BG,
            fg=COLOR_MUTED,
            font=(FONT_MONO, max(8, size - 5)),
        ).pack(side="right")

        tk.Frame(inner, bg=COLOR_LINE, height=1).pack(fill="x", pady=(7, 0))

        main_text = lines[0] if lines else ""
        model, desc = split_sku_name(main_text) if split_name else (main_text, "")
        extra_lines = list(lines[1:]) if lines else []

        if model:
            tk.Label(
                inner,
                text=model,
                bg=self.BG,
                fg=COLOR_TEXT,
                font=(FONT_UI, size + (1 if desc else 0), "bold"),
                justify="left",
                anchor="w",
                wraplength=520,
            ).pack(fill="x", pady=(9, 0))

        if desc:
            tk.Label(
                inner,
                text=desc,
                bg=self.BG,
                fg=COLOR_MUTED,
                font=(FONT_UI, max(9, size - 2)),
                justify="left",
                anchor="w",
                wraplength=520,
            ).pack(fill="x", pady=(4, 0))

        for line in extra_lines:
            tk.Label(
                inner,
                text=line,
                bg=self.BG,
                fg=COLOR_MUTED,
                font=(FONT_UI, max(9, size - 4)),
                justify="left",
                anchor="w",
                wraplength=520,
            ).pack(fill="x", pady=(4, 0))

        if ops:
            owner = " / ".join(ops.get("owner") or [])
            assistant = " / ".join(ops.get("assistant") or [])
            if owner or assistant:
                tk.Frame(inner, bg=COLOR_LINE, height=1).pack(
                    fill="x", pady=(8, 0)
                )
                ops_row = tk.Frame(inner, bg=self.BG)
                ops_row.pack(fill="x", pady=(6, 0))
                keyword_font = (FONT_UI, max(9, size - 3), "bold")
                name_font = (FONT_UI, max(9, size - 3), "bold")
                if owner:
                    tk.Label(
                        ops_row, text="负责人", bg=self.BG,
                        fg=COLOR_CYAN, font=keyword_font,
                    ).pack(side="left")
                    tk.Label(
                        ops_row, text=owner, bg=self.BG,
                        fg=COLOR_MAGENTA, font=name_font,
                    ).pack(side="left", padx=(4, 0))
                if assistant:
                    tk.Label(
                        ops_row, text="助理", bg=self.BG,
                        fg=COLOR_CYAN, font=keyword_font,
                    ).pack(side="left", padx=(14, 0) if owner else (0, 0))
                    tk.Label(
                        ops_row, text=assistant, bg=self.BG,
                        fg=COLOR_MAGENTA, font=name_font,
                    ).pack(side="left", padx=(4, 0))

        if module_tags:
            tk.Frame(inner, bg=COLOR_LINE, height=1).pack(
                fill="x", pady=(8, 0)
            )
            tags_box = tk.Frame(inner, bg=self.BG)
            tags_box.pack(fill="x", pady=(6, 0))
            show_module_name = bool(
                self.cfg.get("display", {}).get("show_module_name", False)
            )
            for module_index, module in enumerate(module_tags[:5]):
                module_box = tk.Frame(tags_box, bg=self.BG)
                module_box.pack(
                    fill="x", anchor="w",
                    pady=(0 if module_index == 0 else 9, 0),
                )
                module_color = module.get("color") or COLOR_CYAN
                if show_module_name:
                    header = tk.Frame(module_box, bg=self.BG)
                    header.pack(fill="x")
                    tk.Frame(
                        header, bg=module_color, width=2, height=12,
                    ).pack(side="left", padx=(0, 6))
                    tk.Label(
                        header,
                        text=str(module.get("module") or ""),
                        bg=self.BG,
                        fg=module_color,
                        font=(FONT_UI, max(9, size - 3), "bold"),
                    ).pack(side="left")

                grid = tk.Frame(module_box, bg=self.BG)
                grid.pack(
                    fill="x",
                    padx=(9 if show_module_name else 0, 0),
                    pady=(3 if show_module_name else 0, 0),
                )
                grid.columnconfigure(1, weight=1)
                for column_index, column in enumerate(module.get("columns") or []):
                    name = str(column.get("name") or "")
                    value = str(column.get("value") or "")
                    value_color = column.get("color") or COLOR_TEXT
                    # 纯黑在深色 HUD 上不可读；没有显式颜色时也回退默认色。
                    if str(value_color).upper() in ("#000000", "#000", "BLACK"):
                        value_color = COLOR_TEXT
                    tk.Label(
                        grid,
                        text=name,
                        bg=self.BG,
                        fg=COLOR_DIM_CYAN,
                        font=(FONT_UI, max(9, size - 4)),
                        anchor="nw",
                        justify="left",
                    ).grid(
                        row=column_index, column=0, sticky="nw",
                        padx=(0, 14), pady=(1, 0),
                    )
                    tk.Label(
                        grid,
                        text=value,
                        bg=self.BG,
                        fg=value_color,
                        font=(FONT_UI, max(9, size - 4)),
                        anchor="nw",
                        justify="left",
                        wraplength=360,
                    ).grid(
                        row=column_index, column=1, sticky="nw",
                        pady=(1, 0),
                    )

        win.update_idletasks()
        width = win.winfo_reqwidth()
        height = win.winfo_reqheight()

        if anchor is None:
            anchor = win32api.GetCursorPos()
        self.anchor = anchor
        x, y = self._place(anchor[0], anchor[1], width, height)
        win.geometry(f"+{x}+{y}")

        self._draw_ticks(win)

        win.bind("<Button-1>", lambda _event: self.close())

        seconds = int(self.cfg["display"].get("auto_close_seconds", 15))
        if seconds > 0:
            self.after_id = win.after(seconds * 1000, self.close)

    @staticmethod
    def _work_area(x, y):
        class RECT(ctypes.Structure):
            _fields_ = [
                ("left", ctypes.c_long),
                ("top", ctypes.c_long),
                ("right", ctypes.c_long),
                ("bottom", ctypes.c_long),
            ]

        class MONITORINFO(ctypes.Structure):
            _fields_ = [
                ("cbSize", ctypes.c_ulong),
                ("rcMonitor", RECT),
                ("rcWork", RECT),
                ("dwFlags", ctypes.c_ulong),
            ]

        class POINT(ctypes.Structure):
            _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

        user32 = ctypes.windll.user32
        monitor = user32.MonitorFromPoint(POINT(x, y), 2)
        info = MONITORINFO()
        info.cbSize = ctypes.sizeof(info)
        if user32.GetMonitorInfoW(monitor, ctypes.byref(info)):
            area = info.rcWork
            return area.left, area.top, area.right, area.bottom
        return 0, 0, user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)

    @classmethod
    def _place(cls, cursor_x, cursor_y, width, height):
        left, top, right, bottom = cls._work_area(cursor_x, cursor_y)

        x = cursor_x + 18
        y = cursor_y + 26
        if x + width > right - 8:
            x = cursor_x - width - 18
        if y + height > bottom - 8:
            y = cursor_y - height - 18

        x = max(left + 8, min(x, right - width - 8))
        y = max(top + 8, min(y, bottom - height - 8))
        return x, y


# ---------------------------------------------------------------- 悬停控制


def read_excel_cell_text(pos):
    """如果光标在 Excel 单元格上，直接用 COM 读单元格，绕开 OCR。

    OCR 对文本型数字、缩小后的表格、窄列（显示 ####）都不稳；Excel 自己
    有准确的单元格值。只有窗口是 Excel、且单元格在光标下时才走这条路。
    """
    if pythoncom is None or win32com is None:
        return None
    try:
        hwnd = win32gui.WindowFromPoint(pos)
        root = win32gui.GetAncestor(hwnd, win32con.GA_ROOT)
        if not root or win32gui.GetClassName(root) not in ("XLMAIN", "OpusApp"):
            return None
    except Exception:
        return None

    initialized = False
    try:
        pythoncom.CoInitialize()
        initialized = True
        excel = win32com.client.GetActiveObject("Excel.Application")
        windows = []
        try:
            windows = list(excel.Windows)
        except Exception:
            windows = []
        target = None
        for window in windows:
            try:
                win_hwnd = int(window.Hwnd)
                # 光标下的 EXCEL7 grid 或它的父窗口，优先精确匹配
                if win_hwnd == int(hwnd):
                    target = window
                    break
                if win32gui.IsChild(win_hwnd, int(hwnd)) or win32gui.IsChild(
                    int(hwnd), win_hwnd
                ):
                    target = window
                    break
            except Exception:
                continue
        if target is None and len(windows) == 1:
            target = windows[0]
        if target is None and len(windows) == 1:
            target = windows[0]
        if target is None:
            return None

        cell = target.RangeFromPoint(int(pos[0]), int(pos[1]))
        if cell is None:
            return None
        try:
            cell = cell.Cells(1, 1)
        except Exception:
            pass

        value = None
        try:
            value = cell.Value
        except Exception:
            value = None
        if isinstance(value, float) and value.is_integer():
            value = int(value)
        if value is not None and not isinstance(value, (list, tuple)):
            text = str(value).strip()
            if text and text.lower() != "none":
                return text

        try:
            text = str(cell.Text or "").strip()
        except Exception:
            text = ""
        if text and "#" not in text:
            return text
    except Exception:
        logging.debug("读取 Excel 单元格失败", exc_info=True)
    finally:
        if initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
    return None


_EXCEL_COM_LOCK = threading.Lock()
_CODE_HINT = re.compile(r"\d{6,16}|\d{1,3}(?:[.\-]\d{1,4}){2,3}")


def _looks_like_code_text(texts):
    """只有出现像编码的片段才值得重试/记日志，避免普通聊天文字刷屏。"""
    return bool(_CODE_HINT.search("".join(texts)))


def read_excel_cell_text_safe(pos, timeout=0.8):
    """带超时的 Excel 读取；COM 卡住时不再阻塞悬停，后续回退 OCR。"""
    if not _EXCEL_COM_LOCK.acquire(blocking=False):
        return None
    result = {}

    def work():
        try:
            result["text"] = read_excel_cell_text(pos)
        except Exception:
            result["text"] = None
        finally:
            _EXCEL_COM_LOCK.release()

    thread = threading.Thread(target=work, daemon=True)
    thread.start()
    thread.join(timeout)
    return result.get("text")


class HoverController:
    def __init__(self, app):
        self.app = app
        self.last_pos = None
        self.still_since = time.monotonic()
        self.trigger_point = None
        self.detecting = False
        self.last_hit = None
        self._cursor_warn_at = 0.0
        self.generation = 0
        self._last_miss_hash = None
        self._last_miss_pos = None

    def reset(self):
        """切换模式/关闭触发时，丢掉正在进行或已经过期的悬停结果。"""
        self.generation += 1
        self.last_pos = None
        self.still_since = time.monotonic()
        self.trigger_point = None
        self.detecting = False
        self._last_miss_hash = None
        self._last_miss_pos = None

    @staticmethod
    def _cursor_pos():
        try:
            return win32api.GetCursorPos()
        except Exception:
            return None

    def start(self):
        self.app.root.after(100, self._tick)

    def _tick(self):
        try:
            self._check()
        except Exception:
            logging.exception("悬停检测异常")
        finally:
            self.app.root.after(100, self._tick)

    def _check(self):
        cfg = self.app.cfg
        if not cfg["trigger"].get("hover_enabled", True):
            if self.last_pos is not None or self.trigger_point is not None or self.detecting:
                self.reset()
            return

        pos = self._cursor_pos()
        now = time.monotonic()
        if pos is None:
            # 锁屏、UAC 安全桌面等情况下会短暂读不到鼠标位置，过一会儿会自动恢复。
            if now - self._cursor_warn_at > 60:
                logging.warning("暂时读不到鼠标位置（锁屏或安全桌面），稍后自动重试")
                self._cursor_warn_at = now
            return

        if self.last_pos is None:
            self.last_pos = pos
            self.still_since = now
            return

        moved = math.hypot(pos[0] - self.last_pos[0], pos[1] - self.last_pos[1])
        if moved > 2:
            self.last_pos = pos
            self.still_since = now

            if self.app.popup.anchor:
                anchor = self.app.popup.anchor
                if math.hypot(pos[0] - anchor[0], pos[1] - anchor[1]) > 120:
                    self.app.popup.close()

            if self.trigger_point and math.hypot(
                pos[0] - self.trigger_point[0], pos[1] - self.trigger_point[1]
            ) > 40:
                self.trigger_point = None
            return

        if self.detecting or self.trigger_point is not None:
            return
        if not self.app.store.mapping:
            return
        if now - self.still_since < cfg["trigger"]["hover_dwell_ms"] / 1000:
            return
        if self._is_own_window(pos):
            return

        self.trigger_point = pos
        self.detecting = True
        self.app.ocr_executor.submit(self._detect, pos, self.generation)

    @staticmethod
    def _is_own_window(pos):
        try:
            hwnd = win32gui.WindowFromPoint(pos)
            _thread_id, pid = win32process.GetWindowThreadProcessId(hwnd)
            return pid == os.getpid()
        except Exception:
            return False

    def _detect(self, pos, generation):
        hit = None
        result = None
        digest = None
        try:
            cfg = self.app.cfg
            mapping = self.app.store.mapping
            max_distance = 70 * self.app.dpi_scale

            # Excel 优先走 COM 读单元格：文本型数字、缩小、窄列 #### 都不怕
            cell_text = read_excel_cell_text_safe(pos)
            if cell_text:
                found = lookup_text(cell_text, mapping)
                if found:
                    code, entry = found
                    hit = {
                        "code": code,
                        "entry": entry,
                        "score": 0.0,
                        "distance": 0.0,
                        "bbox": (pos[0], pos[1], pos[0] + 1, pos[1] + 1),
                    }
                    logging.debug(
                        "Excel 单元格命中：%r -> %s", cell_text[:100], code
                    )
                else:
                    logging.debug("Excel 单元格未命中：%r", cell_text[:120])

            if hit is None:
                scale = max(3, int(cfg["trigger"].get("ocr_scale", 3)))
                width = int(
                    cfg["trigger"]["hover_region_width"] * self.app.dpi_scale
                )
                height = int(
                    cfg["trigger"]["hover_region_height"] * self.app.dpi_scale
                )
                image, left, top = capture_around(pos[0], pos[1], width, height)
                digest = None
                try:
                    digest = hash(image.tobytes())
                except Exception:
                    digest = None
                # 屏幕内容没变、上次也失败过，就不重复 OCR 了
                if digest is not None and digest == self._last_miss_hash:
                    return
                result = self.app.ocr.recognize(image, scale)
                hit = extract_best_hit(
                    result, left, top, scale, mapping, pos, max_distance
                )
                texts = [line.text for line in result.lines]
                hint = _looks_like_code_text(texts)
                # 小字号/Excel 缩小时 3 倍可能丢点；只有像编码时才按 4、5 倍重试
                if hit is None and hint:
                    for retry_scale in range(scale + 1, 6):
                        try:
                            retry_result = self.app.ocr.recognize(
                                image, retry_scale
                            )
                        except Exception:
                            break
                        retry_hit = extract_best_hit(
                            retry_result,
                            left,
                            top,
                            retry_scale,
                            mapping,
                            pos,
                            max_distance,
                        )
                        if retry_hit:
                            hit = retry_hit
                            result = retry_result
                            scale = retry_scale
                            texts = [line.text for line in result.lines]
                            break
                if hit is None and hint:
                    logging.debug(
                        "悬停未命中，OCR 内容：%s", " | ".join(texts)[:600]
                    )
            if hit is None:
                if digest is not None:
                    self._last_miss_hash = digest
                    self._last_miss_pos = pos
            else:
                self._last_miss_hash = None
                self._last_miss_pos = None
        except OcrUnavailable as exc:
            logging.warning("OCR 不可用：%s", exc)
        except Exception:
            logging.exception("悬停识别失败")
        finally:
            self.app.post(lambda: self._finish(hit, pos, generation))

    def _finish(self, hit, pos, generation):
        if generation != self.generation:
            return
        self.detecting = False
        if not self.app.cfg["trigger"].get("hover_enabled", True):
            return
        if self.trigger_point is None:
            return
        if not hit:
            return

        # OCR 期间鼠标已经移开，就不再弹旧位置的结果
        try:
            current = win32api.GetCursorPos()
            if math.hypot(current[0] - pos[0], current[1] - pos[1]) > 40:
                return
        except Exception:
            pass

        code = hit["code"]
        entry = hit["entry"]
        name = entry["name"]

        lines = [name]
        stock = entry.get("stock") or {}
        if self.app.cfg["display"].get("show_stock", True) and stock:
            lines.append("　".join(f"{key} {value}" for key, value in stock.items()))

        ops = None
        if getattr(self.app, "ops_store", None):
            ops = self.app.ops_store.lookup(code)
        module_tags = None
        module_store = getattr(self.app, "module_store", None)
        if module_store is not None:
            try:
                module_tags = module_store.lookup(code)
            except Exception:
                logging.debug("模块标签查询失败", exc_info=True)
        self.app.popup.show(
            code,
            lines,
            detect_styles(name),
            anchor=pos,
            split_name=True,
            ops=ops,
            module_tags=module_tags,
        )
        self.last_hit = code


# ---------------------------------------------------------------- 托盘

import pystray
from PIL import Image, ImageDraw, ImageFont, ImageGrab

try:
    from winrt.windows.globalization import Language
    from winrt.windows.graphics.imaging import (
        BitmapAlphaMode,
        BitmapDecoder,
        BitmapPixelFormat,
    )
    from winrt.windows.media.ocr import OcrEngine
    from winrt.windows.storage.streams import DataWriter, InMemoryRandomAccessStream
    OCR_AVAILABLE = True
except Exception as exc:  # pragma: no cover
    # OCR 只是悬停查询的增强功能；缺少 WinRT OCR 组件的电脑仍应能打开主程序。
    logging.exception("OCR 组件加载失败：%s", exc)
    OCR_AVAILABLE = False
    Language = None
    BitmapAlphaMode = None
    BitmapDecoder = None
    BitmapPixelFormat = None
    OcrEngine = None
    DataWriter = None
    InMemoryRandomAccessStream = None


def make_tray_image(size=64):
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    pad = max(1, size // 32)
    stroke = max(1, size // 32)

    draw.rounded_rectangle(
        [pad, pad, size - pad - 1, size - pad - 1],
        radius=max(6, size // 5),
        fill=(5, 7, 13, 255),
        outline=(53, 231, 255, 255),
        width=stroke,
    )
    # 品红角标
    tick = max(3, size // 9)
    draw.line(
        [(size - pad - tick, pad + 1), (size - pad - 1, pad + 1)],
        fill=(255, 61, 154, 255), width=stroke,
    )
    draw.line(
        [(size - pad - 1, pad + 1), (size - pad - 1, pad + tick)],
        fill=(255, 61, 154, 255), width=stroke,
    )

    font = None
    for font_path, ratio in (
        (r"C:\Windows\Fonts\bahnschrift.ttf", 0.52),
        (r"C:\Windows\Fonts\msyhbd.ttc", 0.44),
    ):
        try:
            font = ImageFont.truetype(font_path, int(size * ratio))
            break
        except Exception:
            continue
    if font is None:
        font = ImageFont.load_default()
    draw.text(
        (size / 2, size / 2 + 1), "S", font=font, fill=(232, 243, 255, 255), anchor="mm"
    )
    return image


# ---------------------------------------------------------------- 开机启动


RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
RUN_NAME = "STag"


def startup_command():
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}" --tray'
    exe = sys.executable
    pythonw = os.path.join(os.path.dirname(exe), "pythonw.exe")
    if os.path.exists(pythonw):
        exe = pythonw
    return f'"{exe}" "{os.path.abspath(__file__)}" --tray'


def startup_enabled():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as key:
            value, _ = winreg.QueryValueEx(key, RUN_NAME)
            return bool(value)
    except OSError:
        return False


def set_startup(enabled):
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE
        ) as key:
            if enabled:
                winreg.SetValueEx(key, RUN_NAME, 0, winreg.REG_SZ, startup_command())
            else:
                try:
                    winreg.DeleteValue(key, RUN_NAME)
                except FileNotFoundError:
                    pass
        return True
    except OSError:
        logging.exception("设置开机启动失败")
        return False


# ---------------------------------------------------------------- 主界面


class App:
    def __init__(self):
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

        self.dpi_scale = self._dpi_scale()
        self.cfg, self._config_existed = self._load_initial_config()

        self.root = tk.Tk()
        self.root.title("SKU // 速查终端")
        self.root.configure(bg=COLOR_BG)
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        window_w, window_h = 660, 600
        pos_x = max(0, (screen_w - window_w) // 2)
        pos_y = max(0, (screen_h - window_h) // 2 - 30)
        self.root.geometry(f"{window_w}x{window_h}+{pos_x}+{pos_y}")
        self.root.minsize(620, 560)
        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)
        self._set_window_icon()
        self._enable_dark_titlebar()
        self._init_fonts()

        self._callbacks = queue.Queue()
        self.root.after(50, self._drain_callbacks)

        self.store = DataStore(self.cfg, self._on_data_status)
        self.ops_store = OperationsStore(self.cfg.get("operations", {}))
        self.ocr = OcrEnginePool()
        self.ocr_executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
        self.popup = Popup(self.root, self.cfg)
        self.hotkey = self._parse_hotkey_safe(self.cfg["trigger"]["hotkey"])
        self.hotkey_state = False
        self.icon = None
        self._quitting = False

        self._build_ui()
        self._sync_ui_from_config()

        self.store.start()
        self.ops_store.start()
        self._start_tray()
        self.hover = HoverController(self)
        self.hover.start()

        self.root.after(100, self._hotkey_tick)

        if "--tray" in sys.argv:
            self.root.withdraw()
        else:
            self.root.after(200, self.open_settings)
        self.root.after(500, self._check_show_signal)

        if not self.ocr.available:
            self._on_data_status(
                "未找到系统 OCR 语言包，悬停模式不可用；按键模式仍可使用", "warn"
            )

    # ---- 基础

    @staticmethod
    def _dpi_scale():
        try:
            dpi = ctypes.windll.user32.GetDpiForSystem()
            return max(1.0, dpi / 96.0)
        except Exception:
            return 1.0

    def _load_initial_config(self):
        existed = os.path.exists(CONFIG_PATH) or os.path.exists(USER_SETTINGS_PATH)
        cfg = load_config()
        if not existed:
            try:
                save_config(cfg)
            except OSError:
                logging.exception("首次保存配置失败")
        return cfg, existed

    def _set_window_icon(self):
        icon = resource_path(os.path.join("assets", "icon.ico"))
        if os.path.exists(icon):
            try:
                self.root.iconbitmap(icon)
            except Exception:
                logging.debug("设置窗口图标失败", exc_info=True)

    def _enable_dark_titlebar(self):
        """让 Windows 标题栏也变成深色，和 HUD 主题一致。"""
        try:
            hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
            value = ctypes.c_int(1)
            for attribute in (20, 19):  # 20 = Win10 20H1+，19 = 更早版本
                result = ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd, attribute, ctypes.byref(value), ctypes.sizeof(value)
                )
                if result == 0:
                    break
        except Exception:
            logging.debug("设置深色标题栏失败", exc_info=True)

    def _init_fonts(self):
        self.root.option_add("*Font", (FONT_UI, 10))
        self.root.option_add("*Background", COLOR_BG)
        self.root.option_add("*Foreground", COLOR_TEXT)
        self.root.option_add("*Entry.Background", COLOR_INPUT)
        self.root.option_add("*Entry.Foreground", COLOR_TEXT)
        self.root.option_add("*SelectBackground", COLOR_CYAN)
        self.root.option_add("*SelectForeground", COLOR_BADGE_INK)

    # ---- 赛博朋克控件工厂

    def _label(self, parent, text, fg=COLOR_TEXT, size=10, bold=False,
               mono=False, bg=COLOR_PANEL, **kwargs):
        return tk.Label(
            parent,
            text=text,
            bg=bg,
            fg=fg,
            font=(FONT_MONO if mono else FONT_UI, size, "bold" if bold else "normal"),
            **kwargs,
        )

    def _button(self, parent, text, command, primary=False, danger=False):
        fg = COLOR_CYAN if primary else (COLOR_MAGENTA if danger else COLOR_TEXT)
        bg = "#08222B" if primary else COLOR_PANEL_2
        border = COLOR_CYAN if primary else (COLOR_MAGENTA if danger else COLOR_LINE)
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground="#0C3540" if primary else "#132038",
            activeforeground=COLOR_CYAN if not danger else COLOR_MAGENTA,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=border,
            highlightcolor=COLOR_CYAN,
            font=(FONT_UI, 10, "bold" if primary else "normal"),
            padx=13,
            pady=5,
            cursor="hand2",
        )

    def _entry(self, parent, textvariable=None, width=None, mono=False):
        kwargs = {}
        if textvariable is not None:
            kwargs["textvariable"] = textvariable
        if width is not None:
            kwargs["width"] = width
        return tk.Entry(
            parent,
            bg=COLOR_INPUT,
            fg=COLOR_TEXT,
            insertbackground=COLOR_CYAN,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=COLOR_LINE,
            highlightcolor=COLOR_CYAN,
            font=(FONT_MONO if mono else FONT_UI, 10),
            **kwargs,
        )

    def _spinbox(self, parent, from_, to, textvariable, width=6):
        return tk.Spinbox(
            parent,
            from_=from_,
            to=to,
            textvariable=textvariable,
            width=width,
            bg=COLOR_INPUT,
            fg=COLOR_TEXT,
            insertbackground=COLOR_CYAN,
            buttonbackground=COLOR_PANEL_2,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=COLOR_LINE,
            highlightcolor=COLOR_CYAN,
            font=(FONT_MONO, 10),
        )

    def _section(self, title, accent=COLOR_CYAN):
        wrap = tk.Frame(self._content, bg=COLOR_BG)
        wrap.pack(fill="x", pady=(0, 10))
        tk.Frame(wrap, bg=accent, width=2).pack(side="left", fill="y")
        panel = tk.Frame(
            wrap, bg=COLOR_PANEL, highlightthickness=1,
            highlightbackground=COLOR_LINE,
        )
        panel.pack(side="left", fill="x", expand=True)

        head = tk.Frame(panel, bg=COLOR_PANEL)
        head.pack(fill="x", padx=12, pady=(9, 5))
        self._label(head, title, fg=accent, size=10, bold=True).pack(side="left")

        body = tk.Frame(panel, bg=COLOR_PANEL)
        body.pack(fill="x", padx=12, pady=(0, 11))
        return body

    @staticmethod
    def _parse_hotkey_safe(combo):
        try:
            return parse_hotkey(combo)
        except Exception:
            logging.exception("热键解析失败：%s", combo)
            return parse_hotkey("F8")

    # ---- 界面

    def _build_ui(self):
        outer = tk.Frame(self.root, bg=COLOR_BG, padx=14, pady=12)
        outer.pack(fill="both", expand=True)

        # 顶栏：SKU//速查 + 数据链状态
        header = tk.Frame(outer, bg=COLOR_BG)
        header.pack(fill="x")

        brand = tk.Frame(header, bg=COLOR_BG)
        brand.pack(side="left", anchor="w")
        tk.Label(
            brand, text="SKU", bg=COLOR_BG, fg=COLOR_TEXT,
            font=(FONT_DISPLAY, 26, "bold"),
        ).pack(side="left")
        tk.Label(
            brand, text="//", bg=COLOR_BG, fg=COLOR_CYAN,
            font=(FONT_MONO, 17, "bold"),
        ).pack(side="left", padx=(3, 0), pady=(7, 0))
        tk.Label(
            brand, text="速查", bg=COLOR_BG, fg=COLOR_TEXT,
            font=(FONT_UI, 19, "bold"),
        ).pack(side="left", padx=(1, 0), pady=(4, 0))

        self.header_status_var = tk.StringVar(value="● 数据链 初始化")
        self.header_status = tk.Label(
            header, textvariable=self.header_status_var, bg=COLOR_BG,
            fg=COLOR_MUTED, font=(FONT_MONO, 9),
        )
        self.header_status.pack(side="right", anchor="e")

        tk.Label(
            outer,
            text="客服查码终端 · 悬停或选中即读取完整 SKU 名称与款式",
            bg=COLOR_BG, fg=COLOR_MUTED, font=(FONT_UI, 10),
        ).pack(anchor="w", pady=(3, 0))

        # 信号线：青色主段 + 品红脉冲段
        signal_bar = tk.Frame(outer, bg=COLOR_LINE, height=2)
        signal_bar.pack(fill="x", pady=(10, 12))
        signal_bar.pack_propagate(False)
        tk.Frame(signal_bar, bg=COLOR_CYAN, width=96).pack(side="left", fill="y")
        tk.Frame(signal_bar, bg=COLOR_MAGENTA, width=18).pack(
            side="left", fill="y", padx=(2, 0)
        )

        # 可滚动内容区
        body = tk.Frame(outer, bg=COLOR_BG)
        body.pack(fill="both", expand=True)

        canvas = tk.Canvas(body, bg=COLOR_BG, highlightthickness=0, bd=0)
        scrollbar = tk.Scrollbar(
            body, orient="vertical", command=canvas.yview,
            bg=COLOR_PANEL, troughcolor=COLOR_BG,
            activebackground=COLOR_CYAN, highlightthickness=0, bd=0, width=10,
        )
        self._content = tk.Frame(canvas, bg=COLOR_BG)

        window_id = canvas.create_window((0, 0), window=self._content, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self._content.bind(
            "<Configure>",
            lambda _event: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.bind(
            "<Configure>",
            lambda event: canvas.itemconfigure(window_id, width=event.width),
        )
        canvas.bind_all(
            "<MouseWheel>",
            lambda event: canvas.yview_scroll(int(-event.delta / 120), "units"),
        )

        # 数据源
        data_box = self._section("数据源")
        self._label(data_box, "Excel 路径", fg=COLOR_MUTED, size=9).pack(anchor="w")
        path_row = tk.Frame(data_box, bg=COLOR_PANEL)
        path_row.pack(fill="x", pady=(6, 0))
        self.path_var = tk.StringVar()
        self._entry(path_row, self.path_var).pack(
            side="left", fill="x", expand=True, ipady=5
        )
        self._button(path_row, "浏览", self.browse_source).pack(
            side="left", padx=(8, 0)
        )

        ops_row = tk.Frame(data_box, bg=COLOR_PANEL)
        ops_row.pack(fill="x", pady=(6, 0))
        self._label(
            ops_row, "负责人关系表", fg=COLOR_MUTED, size=9
        ).pack(side="left", padx=(0, 8))
        self.ops_path_var = tk.StringVar()
        self._entry(ops_row, self.ops_path_var).pack(
            side="left", fill="x", expand=True, ipady=5
        )
        self._button(ops_row, "浏览", self.browse_ops_source).pack(
            side="left", padx=(8, 0)
        )

        self.status_var = tk.StringVar(value="正在启动…")
        self.status_label = tk.Label(
            data_box, textvariable=self.status_var, bg=COLOR_PANEL,
            fg=COLOR_MUTED, font=(FONT_UI, 9), anchor="w",
            justify="left", wraplength=560,
        )
        self.status_label.pack(fill="x", pady=(8, 0))

        row2 = tk.Frame(data_box, bg=COLOR_PANEL)
        row2.pack(fill="x", pady=(10, 0))
        self._button(row2, "立即重新读取", self.reload_data).pack(side="left")
        self._button(row2, "打开所在文件夹", self.open_source_folder).pack(
            side="left", padx=(8, 0)
        )

        # 触发方式
        trigger_box = self._section("触发方式")
        self.hover_enabled_var = tk.BooleanVar(value=True)
        self.hotkey_enabled_var = tk.BooleanVar(value=True)

        row3 = tk.Frame(trigger_box, bg=COLOR_PANEL)
        row3.pack(fill="x")
        self.hover_switch = tk.Label(
            row3, text="鼠标悬停显示", bg=COLOR_PANEL, fg=COLOR_MUTED,
            font=(FONT_UI, 10, "bold"), padx=10, pady=4, cursor="hand2",
        )
        self.hover_switch.pack(side="left")
        self.hover_switch.bind(
            "<Button-1>", lambda _event: self._toggle_trigger_ui("hover")
        )
        self._label(row3, "停留", fg=COLOR_MUTED, size=9).pack(
            side="left", padx=(18, 5)
        )
        self.dwell_var = tk.IntVar(value=600)
        self._spinbox(row3, 200, 2000, self.dwell_var).pack(side="left")
        self._label(row3, "毫秒", fg=COLOR_MUTED, size=9).pack(
            side="left", padx=(4, 0)
        )

        row4 = tk.Frame(trigger_box, bg=COLOR_PANEL)
        row4.pack(fill="x", pady=(9, 0))
        self.hotkey_switch = tk.Label(
            row4, text="选中后按热键", bg=COLOR_PANEL, fg=COLOR_MUTED,
            font=(FONT_UI, 10, "bold"), padx=10, pady=4, cursor="hand2",
        )
        self.hotkey_switch.pack(side="left")
        self.hotkey_switch.bind(
            "<Button-1>", lambda _event: self._toggle_trigger_ui("hotkey")
        )
        self.hotkey_var = tk.StringVar(value="F8")
        self._entry(row4, self.hotkey_var, width=10, mono=True).pack(
            side="left", padx=(12, 8), ipady=4
        )
        self._label(
            row4, "如 F8 或 Ctrl+Alt+S", fg=COLOR_MUTED, size=9
        ).pack(side="left")

        self._label(
            trigger_box,
            "两个开关相互独立，可以同时打开；都关掉时不响应悬停和热键。改动立即生效。",
            fg=COLOR_MUTED, size=9,
        ).pack(anchor="w", pady=(9, 0))

        # 显示
        display_box = self._section("显示")
        row5 = tk.Frame(display_box, bg=COLOR_PANEL)
        row5.pack(fill="x")

        self._label(row5, "自动关闭", fg=COLOR_MUTED, size=9).pack(side="left")
        self.autoclose_var = tk.IntVar(value=15)
        self._spinbox(row5, 0, 120, self.autoclose_var).pack(
            side="left", padx=(6, 4)
        )
        self._label(row5, "秒", fg=COLOR_MUTED, size=9).pack(side="left")

        self._label(row5, "字号", fg=COLOR_MUTED, size=9).pack(
            side="left", padx=(20, 6)
        )
        self.font_var = tk.IntVar(value=13)
        self._spinbox(row5, 9, 24, self.font_var).pack(side="left")

        self.show_stock_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            row5, text="显示库存", variable=self.show_stock_var,
            bg=COLOR_PANEL, fg=COLOR_TEXT, activebackground=COLOR_PANEL,
            activeforeground=COLOR_CYAN, selectcolor=COLOR_CYAN,
            highlightthickness=0, bd=0, font=(FONT_UI, 10), cursor="hand2",
        ).pack(side="left", padx=(20, 0))

        row_opacity = tk.Frame(display_box, bg=COLOR_PANEL)
        row_opacity.pack(fill="x", pady=(9, 0))
        self._label(row_opacity, "弹窗透明度", fg=COLOR_MUTED, size=9).pack(
            side="left"
        )
        self.opacity_var = tk.IntVar(value=98)
        self._spinbox(row_opacity, 30, 100, self.opacity_var).pack(
            side="left", padx=(6, 4)
        )
        self._label(row_opacity, "%", fg=COLOR_MUTED, size=9).pack(side="left")
        self._label(
            row_opacity, "30–100，数值越小越透明", fg=COLOR_MUTED, size=9
        ).pack(side="left", padx=(10, 0))

        # 启动
        system_box = self._section("启动", accent=COLOR_AMBER)
        self.startup_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            system_box, text="开机自动启动（登录后缩到托盘）",
            variable=self.startup_var, bg=COLOR_PANEL, fg=COLOR_TEXT,
            activebackground=COLOR_PANEL, activeforeground=COLOR_AMBER,
            selectcolor=COLOR_AMBER, highlightthickness=0, bd=0,
            font=(FONT_UI, 10), cursor="hand2",
        ).pack(anchor="w")

        # 测试
        test_box = self._section("测试查询", accent=COLOR_MAGENTA)
        row6 = tk.Frame(test_box, bg=COLOR_PANEL)
        row6.pack(fill="x")
        self._label(row6, "编码", fg=COLOR_MUTED, size=9).pack(side="left")
        self.test_var = tk.StringVar()
        self._entry(row6, self.test_var, width=22, mono=True).pack(
            side="left", padx=(8, 0), ipady=4
        )
        self._button(row6, "查询", self.test_query).pack(side="left", padx=(8, 0))

        self.test_result_var = tk.StringVar(value="")
        self.test_result_label = tk.Label(
            test_box, textvariable=self.test_result_var, bg=COLOR_PANEL,
            fg=COLOR_TEXT, font=(FONT_UI, 9), anchor="w",
            justify="left", wraplength=560,
        )
        self.test_result_label.pack(fill="x", pady=(8, 0))

        # 固定底部
        footer = tk.Frame(outer, bg=COLOR_BG)
        footer.pack(fill="x", pady=(8, 0))
        tk.Frame(footer, bg=COLOR_LINE, height=1).pack(fill="x", pady=(0, 8))
        tk.Label(
            footer,
            text="关闭窗口会缩到右下角托盘。彻底退出：托盘右键 → 退出。",
            bg=COLOR_BG, fg=COLOR_MUTED, font=(FONT_UI, 9),
        ).pack(anchor="w", pady=(0, 8))

        buttons = tk.Frame(footer, bg=COLOR_BG)
        buttons.pack(fill="x")
        self._button(
            buttons, "保存并应用", self.save_settings, primary=True
        ).pack(side="left")
        self._button(buttons, "关闭窗口", self.hide_window).pack(
            side="left", padx=(8, 0)
        )
        self._button(buttons, "退出软件", self.quit, danger=True).pack(side="right")

    def _sync_ui_from_config(self):
        cfg = self.cfg
        self.path_var.set(cfg["data_source"]["path"])
        self.ops_path_var.set(cfg.get("operations", {}).get("path", ""))
        self.hover_enabled_var.set(bool(cfg["trigger"].get("hover_enabled", True)))
        self.hotkey_enabled_var.set(bool(cfg["trigger"].get("hotkey_enabled", True)))
        self.dwell_var.set(int(cfg["trigger"]["hover_dwell_ms"]))
        self.hotkey_var.set(cfg["trigger"]["hotkey"])
        self.autoclose_var.set(int(cfg["display"]["auto_close_seconds"]))
        self.font_var.set(int(cfg["display"]["font_size"]))
        self.opacity_var.set(
            int(round(float(cfg["display"].get("opacity", 0.98)) * 100))
        )
        self.show_stock_var.set(bool(cfg["display"]["show_stock"]))
        self.startup_var.set(startup_enabled())
        self._refresh_switch_ui()

    def _collect_config(self):
        cfg = json.loads(json.dumps(self.cfg))  # 深拷贝
        cfg["data_source"]["path"] = self.path_var.get().strip()
        cfg.setdefault("operations", {})
        cfg["operations"]["path"] = self.ops_path_var.get().strip()
        cfg["trigger"]["hover_enabled"] = bool(self.hover_enabled_var.get())
        cfg["trigger"]["hotkey_enabled"] = bool(self.hotkey_enabled_var.get())
        cfg["trigger"]["mode"] = trigger_mode(cfg["trigger"])
        cfg["trigger"]["hover_dwell_ms"] = max(200, int(self.dwell_var.get()))
        cfg["trigger"]["hotkey"] = self.hotkey_var.get().strip() or "F8"
        cfg["display"]["auto_close_seconds"] = max(0, int(self.autoclose_var.get()))
        cfg["display"]["font_size"] = max(9, int(self.font_var.get()))
        cfg["display"]["opacity"] = max(
            0.3, min(1.0, int(self.opacity_var.get()) / 100.0)
        )
        cfg["display"]["show_stock"] = bool(self.show_stock_var.get())
        cfg["system"]["start_with_windows"] = bool(self.startup_var.get())
        return cfg

    def _toggle_trigger_ui(self, kind):
        var = (
            self.hover_enabled_var
            if kind == "hover"
            else self.hotkey_enabled_var
        )
        var.set(not bool(var.get()))
        self._apply_triggers()

    def _refresh_switch_ui(self):
        try:
            hover_on = bool(self.hover_enabled_var.get())
            self.hover_switch.configure(
                bg=COLOR_CYAN if hover_on else COLOR_PANEL,
                fg=COLOR_BADGE_INK if hover_on else COLOR_MUTED,
            )
            hotkey_on = bool(self.hotkey_enabled_var.get())
            self.hotkey_switch.configure(
                bg=COLOR_MAGENTA if hotkey_on else COLOR_PANEL,
                fg=COLOR_BADGE_INK if hotkey_on else COLOR_MUTED,
            )
        except Exception:
            pass

    def _apply_triggers(self):
        """设置里的两个开关即时生效，并分别保存。"""
        trigger = self.cfg["trigger"]
        trigger["hover_enabled"] = bool(self.hover_enabled_var.get())
        trigger["hotkey_enabled"] = bool(self.hotkey_enabled_var.get())
        self._persist_triggers()

    def _persist_triggers(self, announce=False):
        trigger = self.cfg["trigger"]
        trigger["mode"] = trigger_mode(trigger)
        if getattr(self, "hover", None):
            self.hover.reset()
        if not trigger["hover_enabled"]:
            self.popup.close()
        try:
            save_config(self.cfg)
        except OSError:
            logging.exception("保存触发开关失败")
        if self.icon:
            try:
                self.icon.update_menu()
            except Exception:
                pass
        self._update_tray_title()
        if announce:
            hover = trigger["hover_enabled"]
            hotkey = trigger["hotkey_enabled"]
            if not hover and not hotkey:
                message = "悬停和热键都已关闭"
            elif hover and hotkey:
                message = "悬停和热键都已打开"
            elif hover:
                message = "鼠标悬停已打开，热键已关闭"
            else:
                message = "热键已打开，鼠标悬停已关闭"
            self._on_data_status(message, "ok")
        self._refresh_switch_ui()

    # ---- 线程安全回调

    def post(self, callback):
        """把回调交给 Tk 主线程执行（工作线程 / 托盘线程都要走这里）。"""
        self._callbacks.put(callback)

    def _drain_callbacks(self):
        while True:
            try:
                callback = self._callbacks.get_nowait()
            except queue.Empty:
                break
            try:
                callback()
            except Exception:
                logging.exception("界面回调执行失败")
        try:
            self.root.after(50, self._drain_callbacks)
        except Exception:
            pass

    # ---- 状态

    def _on_data_status(self, text, level="info"):
        status_colors = {
            "ok": COLOR_CYAN,
            "warn": COLOR_AMBER,
            "error": COLOR_MAGENTA,
            "info": COLOR_MUTED,
        }

        def apply():
            self.status_var.set(text)
            self.status_label.configure(fg=status_colors.get(level, COLOR_MUTED))

            count = len(self.store.mapping) if hasattr(self, "store") else 0
            if level == "ok":
                head = f"● 在线 · {count:,}" if count else "● 在线"
                color = COLOR_CYAN
            elif level == "warn":
                head = f"● 缓存 · {count:,}" if count else "● 缓存"
                color = COLOR_AMBER
            elif level == "error":
                head = "● 离线"
                color = COLOR_MAGENTA
            else:
                head = "● 同步中"
                color = COLOR_MUTED
            self.header_status_var.set(head)
            self.header_status.configure(fg=color)
            self._update_tray_title()

        self.post(apply)

    def _update_tray_title(self):
        if not self.icon:
            return
        hover = bool(self.cfg["trigger"].get("hover_enabled", True))
        hotkey = bool(self.cfg["trigger"].get("hotkey_enabled", True))
        if hover and hotkey:
            mode = "悬停+热键"
        elif hover:
            mode = "悬停"
        elif hotkey:
            mode = "热键"
        else:
            mode = "关闭"
        count = len(self.store.mapping)
        try:
            self.icon.title = f"SKU//速查 · {count:,} 条 · {mode}"
        except Exception:
            pass

    # ---- 操作

    def _check_show_signal(self):
        try:
            if os.path.exists(SHOW_SIGNAL_PATH):
                os.remove(SHOW_SIGNAL_PATH)
                self.open_settings()
        except Exception:
            pass
        try:
            self.root.after(600, self._check_show_signal)
        except Exception:
            pass

    def open_settings(self):
        try:
            self.root.state("normal")
        except Exception:
            pass
        self.root.deiconify()
        self.root.lift()
        try:
            self.root.focus_force()
        except Exception:
            pass

    def hide_window(self):
        self.root.withdraw()

    def browse_source(self):
        current = self.path_var.get().strip()
        initial = os.path.dirname(current) if current else BASE_DIR
        path = filedialog.askopenfilename(
            title="选择 SKU 数据源",
            initialdir=initial if os.path.isdir(initial) else BASE_DIR,
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.path_var.set(path)

    def browse_ops_source(self):
        current = self.ops_path_var.get().strip()
        initial = os.path.dirname(current) if current else BASE_DIR
        path = filedialog.askopenfilename(
            title="选择负责人关系表",
            initialdir=initial if os.path.isdir(initial) else BASE_DIR,
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.ops_path_var.set(path)

    def open_source_folder(self):
        path = self.path_var.get().strip()
        folder = os.path.dirname(path)
        if folder and os.path.isdir(folder):
            try:
                os.startfile(folder)
            except Exception:
                messagebox.showerror("打不开", f"无法打开文件夹：\n{folder}")
        else:
            messagebox.showwarning("路径不存在", f"找不到文件夹：\n{folder or path}")

    def reload_data(self):
        self.store.reload(silent=False)
        if getattr(self, "ops_store", None):
            self.ops_store.reload()

    def save_settings(self):
        try:
            cfg = self._collect_config()
            if cfg["trigger"]["hotkey_enabled"]:
                parse_hotkey(cfg["trigger"]["hotkey"])
        except Exception as exc:
            messagebox.showerror("设置有误", str(exc))
            return

        path = cfg["data_source"]["path"]
        if not os.path.exists(path):
            if not messagebox.askyesno(
                "数据源不存在",
                f"当前路径读不到：\n{path}\n\n仍要保存吗？（共享盘未连接时可以保存）",
            ):
                return

        self.cfg = cfg
        self.popup.cfg = cfg
        if self.popup.win is not None:
            try:
                self.popup.win.attributes(
                    "-alpha",
                    max(
                        0.3,
                        min(1.0, float(cfg["display"].get("opacity", 0.98))),
                    ),
                )
            except Exception:
                pass
        if getattr(self, "hover", None):
            self.hover.reset()
        if not cfg["trigger"].get("hover_enabled", True):
            self.popup.close()
        try:
            saved_to = save_config(cfg)
            logging.info("配置已保存到 %s", saved_to)
        except OSError as exc:
            messagebox.showerror("保存失败", str(exc))
            return

        self.hotkey = self._parse_hotkey_safe(cfg["trigger"]["hotkey"])
        self.store.update_config(cfg)
        self.ops_store.update_config(cfg)
        set_startup(cfg["system"]["start_with_windows"])
        if self.icon:
            try:
                self.icon.update_menu()
            except Exception:
                pass
        self._update_tray_title()
        where = "程序目录 config.json（共享）" if saved_to == CONFIG_PATH else "本机 AppData"
        self._on_data_status(f"设置已保存 · 数据源写入{where}", "ok")

    def test_query(self):
        text = self.test_var.get().strip()
        if not text:
            return
        mapping = self.store.mapping
        if not mapping:
            self.test_result_var.set("数据还没加载完，请稍候")
            self.test_result_label.configure(fg=COLOR_AMBER)
            return
        hit = lookup_text(text, mapping)
        if not hit:
            self.test_result_var.set("未找到，检查编码是否在当前数据源里")
            self.test_result_label.configure(fg=COLOR_MAGENTA)
            return
        code, entry = hit
        tags = "、".join(tag for tag, _ in detect_styles(entry["name"]))
        stock = entry.get("stock") or {}
        stock_text = "　".join(f"{k} {v}" for k, v in stock.items())
        self.test_result_var.set(
            f"{code} · {tags}\n{entry['name']}" + (f"\n{stock_text}" if stock_text else "")
        )
        self.test_result_label.configure(fg=COLOR_CYAN)

    # ---- 热键

    def _hotkey_tick(self):
        try:
            if self.cfg["trigger"].get("hotkey_enabled", False):
                mods, vk = self.hotkey
                down = (win32api.GetAsyncKeyState(vk) & 0x8000) != 0
                if down:
                    down = all(
                        (win32api.GetAsyncKeyState(mod) & 0x8000) != 0 for mod in mods
                    )
                if down and not self.hotkey_state:
                    self.root.after(1, self.do_hotkey_lookup)
                self.hotkey_state = down
            else:
                self.hotkey_state = False
        except Exception:
            logging.exception("热键检测失败")
        finally:
            self.root.after(60, self._hotkey_tick)

    def do_hotkey_lookup(self):
        if not self.cfg["trigger"].get("hotkey_enabled", False):
            return
        mapping = self.store.mapping
        if not mapping:
            self.popup.show("SKU 速查", ["数据还在加载，请稍候……"], [])
            return

        text, old = copy_selection()
        text = (text or "").strip()
        if old:
            set_clipboard_text(old)

        logging.info("F8 复制内容：%r", text[:300])

        if not text:
            logging.info("F8 剪贴板为空，改用鼠标附近 OCR 兜底")
            self._hotkey_ocr_fallback("")
            return

        hit = lookup_text(text, mapping)
        logging.info(
            "F8 查询：%r -> %s", text[:160], hit[0] if hit else "未命中"
        )
        if not hit:
            self._hotkey_ocr_fallback(text)
            return

        self._show_hotkey_hit(hit)

    def _show_hotkey_hit(self, hit):
        code, entry = hit
        name = entry["name"]
        lines = [name]
        stock = entry.get("stock") or {}
        if self.cfg["display"].get("show_stock", True) and stock:
            lines.append("　".join(f"{key} {value}" for key, value in stock.items()))
        ops = self.ops_store.lookup(code) if self.ops_store else None
        self.popup.show(
            code, lines, detect_styles(name), split_name=True, ops=ops
        )

    def _hotkey_ocr_fallback(self, copied_text):
        """剪贴板拿不到或匹配不上时，直接 OCR 鼠标附近，兼容 RPA/富文本卡片。"""
        try:
            pos = win32api.GetCursorPos()
        except Exception:
            pos = (0, 0)
        cfg = self.cfg["trigger"]
        scale = max(3, int(cfg.get("ocr_scale", 3)))
        width = max(
            int(cfg["hover_region_width"] * self.dpi_scale), 800
        )
        height = max(
            int(cfg["hover_region_height"] * self.dpi_scale), 140
        )

        def work():
            hit = None
            try:
                image, left, top = capture_around(pos[0], pos[1], width, height)
                result = self.ocr.recognize(image, scale)
                hit = extract_best_hit(
                    result,
                    left,
                    top,
                    scale,
                    self.store.mapping,
                    pos,
                    120 * self.dpi_scale,
                )
                if hit is None:
                    logging.info(
                        "F8 OCR 兜底未命中：%s",
                        " | ".join(line.text for line in result.lines)[:600],
                    )
            except OcrUnavailable as exc:
                logging.warning("F8 OCR 不可用：%s", exc)
            except Exception:
                logging.exception("F8 OCR 兜底失败")
            self.post(lambda: self._finish_hotkey_fallback(hit, copied_text))

        self.ocr_executor.submit(work)

    def _finish_hotkey_fallback(self, hit, copied_text):
        if hit:
            self._show_hotkey_hit(hit)
            return
        preview = copied_text if len(copied_text) <= 80 else copied_text[:80] + "…"
        if preview:
            self.popup.show(
                "未找到匹配", [f"选中内容：{preview}"], [("未找到", "#ef4444")]
            )
        else:
            self.popup.show(
                "没读到内容",
                ["请重新选中编码，或把鼠标停在编码上再按 F8"],
                [("提示", "#f59e0b")],
            )

    # ---- 托盘

    def _start_tray(self):
        try:
            menu = pystray.Menu(
                pystray.MenuItem(
                    "打开设置", self._tray_action(self.open_settings), default=True
                ),
                pystray.MenuItem(
                    "重新加载数据", self._tray_action(self.reload_data)
                ),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem(
                    "鼠标悬停",
                    self._tray_action(lambda: self.toggle_trigger("hover")),
                    checked=lambda _item: bool(
                        self.cfg["trigger"].get("hover_enabled", True)
                    ),
                ),
                pystray.MenuItem(
                    "选中 + 热键",
                    self._tray_action(lambda: self.toggle_trigger("hotkey")),
                    checked=lambda _item: bool(
                        self.cfg["trigger"].get("hotkey_enabled", True)
                    ),
                ),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("退出", self._tray_action(self.quit)),
            )
            self.icon = pystray.Icon(
                "STag",
                make_tray_image(),
                "STag SKU 速查",
                menu,
            )
            self.icon.run_detached()
            self.root.after(1200, self._tray_notify_ready)
        except Exception:
            logging.exception("托盘图标启动失败")
            self.icon = None

    def _tray_action(self, callback):
        def handler(_icon=None, _item=None):
            self.post(callback)

        return handler

    def _tray_notify_ready(self):
        if not self.icon:
            return
        try:
            self.icon.notify("已启动。右键右下角托盘图标可打开设置或退出。", "STag SKU 速查")
        except Exception:
            logging.debug("托盘通知失败", exc_info=True)

    def toggle_trigger(self, kind):
        """托盘菜单里的两个独立开关。"""
        trigger = self.cfg["trigger"]
        key = "hover_enabled" if kind == "hover" else "hotkey_enabled"
        trigger[key] = not bool(trigger.get(key, True))
        if kind == "hover":
            self.hover_enabled_var.set(trigger[key])
        else:
            self.hotkey_enabled_var.set(trigger[key])
        self._persist_triggers(announce=True)

    # ---- 退出

    def quit(self):
        if self._quitting:
            return
        self._quitting = True
        logging.info("退出程序")
        self.store.stop()
        self.ops_store.stop()
        self.ocr_executor.shutdown(wait=False, cancel_futures=True)
        self.popup.close()
        if self.icon:
            try:
                self.icon.stop()
            except Exception:
                pass
        try:
            self.root.destroy()
        except Exception:
            pass

    def run(self):
        self.root.mainloop()


# ---------------------------------------------------------------- 单实例


def single_instance_guard():
    kernel32 = ctypes.windll.kernel32
    handle = kernel32.CreateMutexW(None, False, "STAG_SKU_LOOKUP_APP")
    if kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
        # 再次双击 exe：通知已经在运行的那个实例把窗口显示出来
        try:
            os.makedirs(CACHE_DIR, exist_ok=True)
            with open(SHOW_SIGNAL_PATH, "w", encoding="utf-8") as signal_file:
                signal_file.write("show")
        except Exception:
            pass
        return None
    return handle


# ---------------------------------------------------------------- 命令行自检


def _make_test_image(text):
    image = Image.new("RGB", (760, 90), "white")
    draw = ImageDraw.Draw(image)
    try:
        font = ImageFont.truetype(r"C:\Windows\Fonts\msyh.ttc", 28)
    except Exception:
        font = ImageFont.load_default()
    draw.text((10, 25), text, fill="black", font=font)
    return image


def _open_temp_sku_db(cfg):
    """自检/命令行用临时库，避免影响正在运行的程序缓存。"""
    temp_path, _meta = build_sku_db(cfg["data_source"]["path"], cfg)
    return temp_path, SkuDatabase(temp_path)


def _close_temp_db(temp_path, db):
    try:
        db.close()
    except Exception:
        pass
    try:
        os.remove(temp_path)
    except OSError:
        pass


def run_selftest():
    lines = []
    ok = True
    temp_path = None
    db = None
    try:
        cfg = load_config()
        temp_path, db = _open_temp_sku_db(cfg)
        lines.append(
            f"数据加载 OK：{db.count:,} 条，"
            f"数据时间 {db.meta.get('data_date','-')}"
        )

        sample = db.first_code()
        if sample:
            lines.append(f"样例编码：{sample} -> {db[sample]['name'][:40]}")

        pool = OcrEnginePool()
        lines.append(f"OCR 语言包：{pool.languages}")

        image = _make_test_image("101010519 苹果17 Pro Max 无边小弧边钢化膜")
        result = pool.recognize(image, 2)
        lines.append(f"OCR 文本：{result.text[:120]}")

        cursor = (image.width // 2, image.height // 2)
        hit = extract_best_hit(result, 0, 0, 2, db, cursor, 10000)
        if hit:
            lines.append(f"OCR 命中的编码：{hit['code']}")
        else:
            lines.append("OCR 未命中编码（如果文本里是真实编码，说明识别或匹配需要调整）")
            ok = False
    except Exception as exc:
        logging.exception("自检失败")
        lines.append(f"FAILED: {exc}")
        ok = False
    finally:
        if temp_path is not None and db is not None:
            _close_temp_db(temp_path, db)

    path = os.path.join(BASE_DIR, "selftest.log")
    try:
        with open(path, "w", encoding="utf-8") as handle:
            handle.write("\n".join(lines))
    except OSError:
        path = os.path.join(tempfile.gettempdir(), "sku_selftest.log")
        with open(path, "w", encoding="utf-8") as handle:
            handle.write("\n".join(lines))

    print("\n".join(lines))
    return 0 if ok else 1


def run_cli_lookup(code):
    cfg = load_config()
    temp_path, db = _open_temp_sku_db(cfg)
    try:
        print(
            f"已加载 {db.count:,} 条 · "
            f"数据时间 {db.meta.get('data_date','-')}"
        )
        hit = lookup_text(code, db)
        if not hit:
            print(f"未找到：{code}")
            return 1
        found_code, entry = hit
        print(f"编码：{found_code}")
        print(f"名称：{entry['name']}")
        print("款式：" + "、".join(tag for tag, _ in detect_styles(entry["name"])))
        stock = entry.get("stock") or {}
        if stock:
            print("库存：" + "　".join(f"{k} {v}" for k, v in stock.items()))
        return 0
    finally:
        _close_temp_db(temp_path, db)


def run_cli_ocr(image_path):
    cfg = load_config()
    temp_path, db = _open_temp_sku_db(cfg)
    try:
        pool = OcrEnginePool()
        image = Image.open(image_path)
        result = pool.recognize(image, 2)
        print("OCR:", result.text)
        cursor = (image.width // 2, image.height // 2)
        hit = extract_best_hit(result, 0, 0, 2, db, cursor, 10000)
        print("HIT:", hit["code"] if hit else None)
        return 0
    finally:
        _close_temp_db(temp_path, db)


# ---------------------------------------------------------------- 入口


def main():
    _setup_logging()
    _cleanup_legacy_cache()

    if "--selftest" in sys.argv:
        return run_selftest()

    if "--test" in sys.argv:
        index = sys.argv.index("--test")
        if index + 1 >= len(sys.argv):
            print("用法：sku_lookup.exe --test 101010519")
            return 2
        return run_cli_lookup(sys.argv[index + 1])

    if "--test-ocr" in sys.argv:
        index = sys.argv.index("--test-ocr")
        if index + 1 >= len(sys.argv):
            print("用法：sku_lookup.exe --test-ocr 图片路径")
            return 2
        return run_cli_ocr(sys.argv[index + 1])

    guard = single_instance_guard()
    if guard is None:
        logging.info("已有实例在运行，退出本次启动")
        return 0

    try:
        app = App()
        if "--settings" in sys.argv:
            app.root.after(300, app.open_settings)
        app.run()
    except Exception:
        logging.exception("程序异常退出")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
