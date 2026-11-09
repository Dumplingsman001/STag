# -*- coding: utf-8 -*-
"""STag · Modular SKU Tag System

在 SKU 查询/OCR 基础上，加入“模块化标签”：
  - 每个模块是一张独立的小表，SKU 编码作为 join key
  - 模块可以导入、更新、停用、删除、导出
  - 查询时只取当前 SKU 的模块行，不把模块全表读进内存
  - 支持批量查码和模块列横向展示
"""

import concurrent.futures
import ctypes
import hashlib
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
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import winreg

import win32api
import pystray
from PIL import Image, ImageDraw, ImageFont, ImageTk
import tkinter.font as tkfont

import sku_lookup as core

APP_NAME = "STag"
APP_TITLE = "STag // SKU TAG SYSTEM"
LOG_PATH = os.path.join(core.BASE_DIR, "stag.log")
MODULE_DB_PATH = os.path.join(core.CACHE_DIR, "modules.db")
MODULE_SOURCE_DIR = os.path.join(core.CACHE_DIR, "module_sources")
STAG_SIGNAL_PATH = os.path.join(core.CACHE_DIR, "stag_show.signal")
RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
RUN_NAME = "STagSkuTag"

MODULE_COLORS = [
    "#35E7FF",
    "#FF3D9A",
    "#8B5CF6",
    "#FFB547",
    "#59FFA8",
    "#FF5C7A",
]


def setup_logging():
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    for path in (LOG_PATH, os.path.join(core.CACHE_DIR, "stag.log")):
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            handler = RotatingFileHandler(
                path, maxBytes=1_000_000, backupCount=2, encoding="utf-8"
            )
            handler.setFormatter(formatter)
            logging.basicConfig(
                level=logging.INFO, handlers=[handler], force=True
            )
            return
        except OSError:
            continue
    logging.basicConfig(level=logging.INFO, force=True)


def enable_dark_app_mode():
    """在创建窗口前打开 Windows 10/11 的深色应用模式。"""
    try:
        uxtheme = ctypes.WinDLL("uxtheme")
        set_preferred = uxtheme[135]
        set_preferred.argtypes = [ctypes.c_int]
        set_preferred.restype = ctypes.c_int
        set_preferred(2)  # AllowDark
    except Exception:
        pass


def enable_dark_titlebar(window):
    try:
        window.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id()) or window.winfo_id()
        try:
            uxtheme = ctypes.WinDLL("uxtheme")
            set_preferred = uxtheme[135]
            set_preferred.argtypes = [ctypes.c_int]
            set_preferred.restype = ctypes.c_int
            set_preferred(2)
            allow = uxtheme[133]
            allow.argtypes = [ctypes.c_void_p, ctypes.c_bool]
            allow.restype = ctypes.c_bool
            allow(hwnd, True)
            uxtheme.SetWindowTheme(hwnd, "DarkMode_Explorer", None)
        except Exception:
            pass
        value = ctypes.c_int(1)
        for attribute in (20, 19):
            result = ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, attribute, ctypes.byref(value), ctypes.sizeof(value)
            )
            if result == 0:
                break
        try:
            ctypes.windll.user32.SetWindowPos(
                hwnd, 0, 0, 0, 0, 0, 0x0020 | 0x0001 | 0x0002 | 0x0004
            )
            ctypes.windll.user32.RedrawWindow(
                hwnd, None, None, 0x0400 | 0x0100 | 0x0001
            )
        except Exception:
            pass
    except Exception:
        pass


def _cut_points(w, h, cut, ox=0, oy=0):
    cut = max(1, min(cut, w / 3, h / 3))
    return [
        (ox + cut, oy),
        (ox + w - cut, oy),
        (ox + w, oy + cut),
        (ox + w, oy + h - cut),
        (ox + w - cut, oy + h),
        (ox + cut, oy + h),
        (ox, oy + h - cut),
        (ox, oy + cut),
    ]


def _scaled_cut_image(w, h, fill, outline, cut, width=1, shadow=False):
    scale = 3
    pad = 6 if shadow else 0
    canvas_w = max(1, int((w + pad) * scale))
    canvas_h = max(1, int((h + pad) * scale))
    image = Image.new("RGBA", (canvas_w, canvas_h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    if shadow:
        shadow_points = _cut_points(
            w * scale, h * scale, cut * scale, pad * scale // 2, pad * scale // 2
        )
        draw.polygon(shadow_points, fill=(0, 0, 0, 110))
    body = _cut_points(w * scale, h * scale, cut * scale, 0, 0)
    draw.polygon(body, fill=fill)
    draw.line(body + [body[0]], fill=outline, width=max(1, int(width * scale)))
    image = image.resize((int(w + pad), int(h + pad)), Image.LANCZOS)
    return image


def make_block_photo(w, h, color, disabled=False, highlight=False):
    fill = hex_blend(color, core.COLOR_BG, 0.05 if disabled else (0.18 if highlight else 0.12))
    outline = core.COLOR_MUTED if disabled else color
    image = _scaled_cut_image(
        int(w), int(h), fill, outline,
        cut=10 if w > 90 and h > 60 else 5,
        width=2 if highlight else 1,
        shadow=True,
    )
    # 顶部信号线
    draw = ImageDraw.Draw(image)
    draw.line(
        [(2, 2), (min(w - 10, 110), 2)],
        fill=color, width=2,
    )
    return ImageTk.PhotoImage(image)


_BUTTON_FONT_CACHE = {}


def _button_font(size):
    key = int(size)
    if key not in _BUTTON_FONT_CACHE:
        path = r"C:\Windows\Fonts\msyhbd.ttc"
        if not os.path.exists(path):
            path = r"C:\Windows\Fonts\bahnschrift.ttf"
        try:
            _BUTTON_FONT_CACHE[key] = ImageFont.truetype(
                path, max(8, int(size * 3))
            )
        except Exception:
            logging.exception("按钮字体加载失败，使用默认字体")
            try:
                _BUTTON_FONT_CACHE[key] = ImageFont.load_default(
                    size=max(8, int(size * 3))
                )
            except TypeError:
                _BUTTON_FONT_CACHE[key] = ImageFont.load_default()
    return _BUTTON_FONT_CACHE[key]


def make_button_photo(w, h, text, accent=False, hover=False, font_size=10):
    if accent:
        fill = "#6BF0FF" if hover else core.COLOR_CYAN
        outline = "#A9FAFF" if hover else core.COLOR_CYAN
    else:
        fill = "#122033" if hover else core.COLOR_PANEL_2
        outline = core.COLOR_CYAN if hover else core.COLOR_LINE_BRIGHT
    scale = 3
    canvas = Image.new("RGBA", (int(w) * scale, int(h) * scale), (0, 0, 0, 0))
    draw = ImageDraw.Draw(canvas)
    points = _cut_points(int(w) * scale, int(h) * scale, 7 * scale)
    draw.polygon(points, fill=fill)
    draw.line(points + [points[0]], fill=outline, width=scale)
    font = _button_font(font_size)
    draw.text(
        (int(w) * scale / 2, int(h) * scale / 2),
        text,
        font=font,
        fill=core.COLOR_BADGE_INK if accent else core.COLOR_TEXT,
        anchor="mm",
    )
    return ImageTk.PhotoImage(canvas.resize((int(w), int(h)), Image.LANCZOS))


class HudButton(tk.Canvas):
    def __init__(self, parent, text, command, accent=False, small=False):
        self.text = text
        self.command = command
        self.accent = accent
        self.small = small
        self._font = tkfont.Font(
            family=core.FONT_UI, size=9 if small else 10, weight="bold"
        )
        self._photos = {}
        self._measure()
        try:
            bg = parent.cget("bg")
        except Exception:
            bg = core.COLOR_BG
        super().__init__(
            parent, width=self.w, height=self.h, bg=bg,
            highlightthickness=0, bd=0, cursor="hand2",
        )
        self.bind("<Button-1>", lambda _event: self.command())
        self.bind("<Enter>", lambda _event: self._draw(hover=True))
        self.bind("<Leave>", lambda _event: self._draw(hover=False))
        self._draw()

    def _measure(self):
        self.w = self._font.measure(self.text) + (24 if self.small else 34)
        self.h = 32 if self.small else 40

    def _photo(self, hover):
        key = (self.w, self.h, self.accent, hover, self.text)
        if key not in self._photos:
            self._photos[key] = make_button_photo(
                self.w, self.h, self.text, self.accent, hover,
                font_size=9 if self.small else 10,
            )
        return self._photos[key]

    def _draw(self, hover=False):
        self.delete("all")
        self.create_image(0, 0, anchor="nw", image=self._photo(hover))

    def configure(self, **kwargs):
        text = kwargs.pop("text", None)
        if text is not None:
            self.text = text
            self._measure()
            super().configure(width=self.w, height=self.h)
            self._draw()
        if kwargs:
            super().configure(**kwargs)

    config = configure


def now_text():
    return time.strftime("%Y-%m-%d %H:%M:%S")


def hex_blend(hex_color, base="#070B13", ratio=0.12):
    try:
        color = hex_color.lstrip("#")
        base_color = base.lstrip("#")
        c = tuple(int(color[i:i + 2], 16) for i in (0, 2, 4))
        b = tuple(int(base_color[i:i + 2], 16) for i in (0, 2, 4))
        mixed = tuple(int(b[i] + (c[i] - b[i]) * ratio) for i in range(3))
        return "#%02X%02X%02X" % mixed
    except Exception:
        return base


def cut_rect(canvas, x, y, w, h, cut=8, **kwargs):
    """带切角的矩形，接近网页版积木的轮廓。"""
    cut = max(0, min(cut, w / 3, h / 3))
    points = [
        x + cut, y,
        x + w - cut, y,
        x + w, y + cut,
        x + w, y + h - cut,
        x + w - cut, y + h,
        x + cut, y + h,
        x, y + h - cut,
        x, y + cut,
    ]
    return canvas.create_polygon(points, **kwargs)


def treemap(items, width, height, gap=8):
    """按 weight 把模块块铺满矩形；返回 [(item, x, y, w, h)]。"""
    if not items or width <= 0 or height <= 0:
        return []
    total = sum(max(0.01, item["weight"]) for item in items)
    result = []

    def split(group, x, y, w, h):
        if not group:
            return
        if len(group) == 1:
            result.append((group[0]["item"], x, y, w, h))
            return
        group_total = sum(max(0.01, item["weight"]) for item in group)
        acc = 0.0
        split_index = 1
        for index in range(len(group) - 1):
            acc += max(0.01, group[index]["weight"])
            split_index = index + 1
            if acc >= group_total / 2:
                break
        split_index = max(1, min(split_index, len(group) - 1))
        first = group[:split_index]
        second = group[split_index:]
        first_weight = sum(max(0.01, item["weight"]) for item in first)
        ratio = first_weight / group_total
        if w >= h:
            w1 = w * ratio
            split(first, x, y, w1, h)
            split(second, x + w1, y, w - w1, h)
        else:
            h1 = h * ratio
            split(first, x, y, w, h1)
            split(second, x, y + h1, w, h - h1)

    split(items, 0, 0, width, height)
    padded = []
    for item, x, y, w, h in result:
        w = max(0, w - gap)
        h = max(0, h - gap)
        if w < 20 or h < 16:
            continue
        padded.append((item, x + gap / 2, y + gap / 2, w, h))
    return padded


class ToggleSwitch(tk.Canvas):
    def __init__(self, parent, variable, on_color=core.COLOR_CYAN, command=None, bg=None):
        super().__init__(
            parent, width=46, height=24, bg=bg or core.COLOR_BG,
            highlightthickness=0, bd=0, cursor="hand2",
        )
        self.variable = variable
        self.on_color = on_color
        self.command = command
        self.bind("<Button-1>", self.toggle)
        self.variable.trace_add("write", lambda *_args: self.redraw())
        self.redraw()

    def toggle(self, _event=None):
        self.variable.set(not bool(self.variable.get()))
        if self.command:
            self.command()

    def set(self, value):
        self.variable.set(bool(value))
        self.redraw()

    def redraw(self):
        self.delete("all")
        on = bool(self.variable.get())
        track = self.on_color if on else core.COLOR_LINE_BRIGHT
        self.create_oval(1, 3, 21, 21, fill=track, outline="")
        self.create_oval(25, 3, 45, 21, fill=track, outline="")
        self.create_rectangle(11, 3, 35, 21, fill=track, outline="")
        knob_x = 26 if on else 5
        self.create_oval(
            knob_x, 4, knob_x + 16, 20,
            fill=core.COLOR_BADGE_INK if on else core.COLOR_MUTED,
            outline="",
        )


class ModuleStore:
    """模块化标签库。每个模块独立成表，查询时按 SKU 编码 join。"""

    def __init__(self, path=MODULE_DB_PATH):
        self.path = path
        os.makedirs(os.path.dirname(path), exist_ok=True)
        os.makedirs(MODULE_SOURCE_DIR, exist_ok=True)
        self._lock = threading.RLock()
        self.conn = sqlite3.connect(
            path, check_same_thread=False, isolation_level=None
        )
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA journal_mode = WAL")
        self.conn.execute("PRAGMA synchronous = NORMAL")
        self.conn.execute("PRAGMA foreign_keys = ON")
        self._init_schema()

    def _init_schema(self):
        self.conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS module (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                description TEXT NOT NULL DEFAULT '',
                color TEXT NOT NULL DEFAULT '#35E7FF',
                priority INTEGER NOT NULL DEFAULT 100,
                status TEXT NOT NULL DEFAULT 'active',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS module_column (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                module_id INTEGER NOT NULL,
                name TEXT NOT NULL,
                position INTEGER NOT NULL DEFAULT 0,
                UNIQUE(module_id, name),
                FOREIGN KEY(module_id) REFERENCES module(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS module_value (
                module_id INTEGER NOT NULL,
                column_id INTEGER NOT NULL,
                code TEXT NOT NULL,
                value TEXT NOT NULL DEFAULT '',
                color TEXT NOT NULL DEFAULT '',
                updated_at TEXT NOT NULL,
                PRIMARY KEY(module_id, column_id, code),
                FOREIGN KEY(module_id) REFERENCES module(id) ON DELETE CASCADE,
                FOREIGN KEY(column_id) REFERENCES module_column(id) ON DELETE CASCADE
            );
            CREATE INDEX IF NOT EXISTS idx_module_value_code
                ON module_value(code);

            CREATE TABLE IF NOT EXISTS action (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kind TEXT NOT NULL,
                module_id INTEGER,
                target_action_id INTEGER,
                source_name TEXT NOT NULL DEFAULT '',
                file_hash TEXT NOT NULL DEFAULT '',
                user TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                summary TEXT NOT NULL DEFAULT '',
                status TEXT NOT NULL DEFAULT 'applied'
            );

            CREATE TABLE IF NOT EXISTS module_version (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                module_id INTEGER NOT NULL,
                version_no INTEGER NOT NULL,
                action_id INTEGER,
                source_name TEXT NOT NULL DEFAULT '',
                file_hash TEXT NOT NULL DEFAULT '',
                row_count INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'active',
                UNIQUE(module_id, version_no),
                FOREIGN KEY(module_id) REFERENCES module(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS module_version_column (
                version_id INTEGER NOT NULL,
                name TEXT NOT NULL,
                position INTEGER NOT NULL DEFAULT 0,
                PRIMARY KEY(version_id, name),
                FOREIGN KEY(version_id) REFERENCES module_version(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS module_version_value (
                version_id INTEGER NOT NULL,
                column_name TEXT NOT NULL,
                code TEXT NOT NULL,
                value TEXT NOT NULL DEFAULT '',
                color TEXT NOT NULL DEFAULT '',
                PRIMARY KEY(version_id, column_name, code),
                FOREIGN KEY(version_id) REFERENCES module_version(id) ON DELETE CASCADE
            );
            """
        )
        for table in ("module_value", "module_version_value"):
            columns = {
                row[1]
                for row in self.conn.execute(f"PRAGMA table_info({table})")
            }
            if "color" not in columns:
                self.conn.execute(
                    f"ALTER TABLE {table} ADD COLUMN color TEXT NOT NULL DEFAULT ''"
                )

    def close(self):
        with self._lock:
            try:
                self.conn.close()
            except Exception:
                pass

    def _row_to_module(self, row):
        return {
            "id": int(row["id"]),
            "name": row["name"],
            "description": row["description"] or "",
            "color": row["color"] or core.COLOR_CYAN,
            "priority": int(row["priority"] or 100),
            "status": row["status"] or "active",
            "created_at": row["created_at"],
            "updated_at": row["updated_at"],
            "row_count": int(row["row_count"] or 0) if "row_count" in row.keys() else 0,
            "col_count": int(row["col_count"] or 0) if "col_count" in row.keys() else 0,
        }

    def list_modules(self, include_deleted=False):
        where = "" if include_deleted else "WHERE m.status != 'deleted'"
        with self._lock:
            rows = self.conn.execute(
                f"""
                SELECT m.*,
                    (SELECT COUNT(DISTINCT v.code) FROM module_value v
                     WHERE v.module_id = m.id) AS row_count,
                    (SELECT COUNT(*) FROM module_column c
                     WHERE c.module_id = m.id) AS col_count
                FROM module m
                {where}
                ORDER BY m.priority, m.name
                """
            ).fetchall()
        return [self._row_to_module(row) for row in rows]

    def get_module(self, module_id):
        with self._lock:
            row = self.conn.execute(
                """
                SELECT m.*,
                    (SELECT COUNT(DISTINCT v.code) FROM module_value v
                     WHERE v.module_id = m.id) AS row_count,
                    (SELECT COUNT(*) FROM module_column c
                     WHERE c.module_id = m.id) AS col_count
                FROM module m WHERE m.id = ?
                """,
                (module_id,),
            ).fetchone()
        return self._row_to_module(row) if row else None

    def columns(self, module_id):
        with self._lock:
            rows = self.conn.execute(
                "SELECT id, name, position FROM module_column "
                "WHERE module_id = ? ORDER BY position, id",
                (module_id,),
            ).fetchall()
        return [dict(row) for row in rows]

    def lookup(self, code):
        code = core.normalize_code(code)
        if not code:
            return []
        with self._lock:
            rows = self.conn.execute(
                """
                SELECT m.id AS module_id, m.name AS module_name, m.color,
                       m.priority, c.name AS column_name, c.position,
                       v.value, v.color AS value_color
                FROM module_value v
                JOIN module m ON m.id = v.module_id
                JOIN module_column c ON c.id = v.column_id
                WHERE v.code = ? AND m.status = 'active'
                ORDER BY m.priority, m.id, c.position, c.id
                """,
                (code,),
            ).fetchall()
        result = {}
        for row in rows:
            mid = int(row["module_id"])
            entry = result.setdefault(
                mid,
                {
                    "module_id": mid,
                    "module": row["module_name"],
                    "color": row["color"] or core.COLOR_CYAN,
                    "columns": [],
                },
            )
            entry["columns"].append(
                {
                    "name": row["column_name"],
                    "value": row["value"] or "",
                    "color": row["value_color"] or "",
                }
            )
        return list(result.values())

    def lookup_many(self, codes):
        codes = [core.normalize_code(code) for code in codes]
        codes = [code for code in dict.fromkeys(codes) if code]
        result = {code: [] for code in codes}
        if not codes:
            return result
        for start in range(0, len(codes), 500):
            chunk = codes[start:start + 500]
            placeholders = ",".join("?" for _ in chunk)
            with self._lock:
                rows = self.conn.execute(
                    f"""
                    SELECT v.code, m.id AS module_id, m.name AS module_name,
                           m.color, m.priority, c.name AS column_name,
                           c.position, v.value, v.color AS value_color
                    FROM module_value v
                    JOIN module m ON m.id = v.module_id
                    JOIN module_column c ON c.id = v.column_id
                    WHERE v.code IN ({placeholders}) AND m.status = 'active'
                    ORDER BY m.priority, m.id, c.position, c.id
                    """,
                    chunk,
                ).fetchall()
            grouped = {}
            for row in rows:
                code = row["code"]
                mid = int(row["module_id"])
                entry = grouped.setdefault(
                    code,
                    {},
                ).setdefault(
                    mid,
                    {
                        "module_id": mid,
                        "module": row["module_name"],
                        "color": row["color"] or core.COLOR_CYAN,
                        "columns": [],
                    },
                )
                entry["columns"].append(
                    {
                        "name": row["column_name"],
                        "value": row["value"] or "",
                        "color": row["value_color"] or "",
                    }
                )
            for code, entries in grouped.items():
                result[code] = list(entries.values())
        return result

    def detail_rows(self, module_id, limit=500):
        columns = self.columns(module_id)
        with self._lock:
            total = self.conn.execute(
                "SELECT COUNT(DISTINCT code) FROM module_value WHERE module_id = ?",
                (module_id,),
            ).fetchone()[0]
            codes = [
                row["code"]
                for row in self.conn.execute(
                    "SELECT DISTINCT code FROM module_value "
                    "WHERE module_id = ? ORDER BY code LIMIT ?",
                    (module_id, limit),
                ).fetchall()
            ]
            rows = []
            if codes:
                placeholders = ",".join("?" for _ in codes)
                raw = self.conn.execute(
                    f"""
                    SELECT v.code, c.name AS column_name, v.value
                    FROM module_value v
                    JOIN module_column c ON c.id = v.column_id
                    WHERE v.module_id = ? AND v.code IN ({placeholders})
                    ORDER BY v.code, c.position, c.id
                    """,
                    [module_id] + codes,
                ).fetchall()
                by_code = {}
                for row in raw:
                    by_code.setdefault(row["code"], {})[row["column_name"]] = row["value"] or ""
                for code in codes:
                    rows.append({"code": code, "values": by_code.get(code, {})})
        return columns, rows, total

    @staticmethod
    def _normalize_excel_color(raw):
        if not raw:
            return ""
        text = str(raw).strip().lstrip("#")
        if len(text) == 8:
            text = text[2:]
        if len(text) == 6 and re.fullmatch(r"[0-9A-Fa-f]{6}", text):
            return "#" + text.upper()
        return ""

    @staticmethod
    def _cell_color(cell):
        """读取 Excel 单元格字体颜色；默认/主题色返回空，沿用 HUD 默认色。"""
        try:
            font = cell.font
            color = font.color if font else None
            if color is None:
                return ""
            kind = getattr(color, "type", None)
            raw = None
            if kind == "rgb":
                raw = color.rgb
            elif kind == "indexed":
                index = color.indexed
                if index is not None:
                    from openpyxl.styles.colors import COLOR_INDEX

                    if 0 <= int(index) < len(COLOR_INDEX):
                        raw = COLOR_INDEX[int(index)]
            return ModuleStore._normalize_excel_color(raw)
        except Exception:
            return ""

    def _unique_column_names(self, header):
        names = []
        used = {}
        for index, raw in enumerate(header):
            name = str(raw or "").strip()
            if not name or "名称" in name:
                continue
            base = name
            count = used.get(base, 0) + 1
            used[base] = count
            if count > 1:
                name = f"{base}（{count}）"
            names.append((index, name))
        return names

    def import_excel(self, path, module_name, description="", color=None, module_id=None):
        load_workbook = core._import_openpyxl()
        module_name = (module_name or "").strip()
        if not module_name:
            raise ValueError("模块名称不能为空")
        if not os.path.exists(path):
            raise ValueError("Excel 文件不存在")

        workbook = load_workbook(path, read_only=True, data_only=True, keep_links=False)
        try:
            sheet = workbook.worksheets[0]
            reset = getattr(sheet, "reset_dimensions", None)
            if reset is not None:
                try:
                    reset()
                except Exception:
                    pass
            rows = sheet.iter_rows(values_only=False)
            header = []
            for row in rows:
                candidate = [
                    ""
                    if cell is None or cell.value is None
                    else str(cell.value).strip()
                    for cell in row
                ]
                if any(candidate):
                    header = candidate
                    break
            code_index = core._column_index(header, "SKU编码", "编码")
            if code_index is None:
                raise ValueError("第一行找不到「SKU编码」或「编码」列")

            tag_columns = self._unique_column_names(header)
            tag_columns = [
                (index, name) for index, name in tag_columns if index != code_index
            ]
            if not tag_columns:
                raise ValueError(
                    "没有找到标签列，请在 C1 起填写自己的标签列名"
                )

            def cell_value(row, index):
                if index is None or index >= len(row):
                    return ""
                cell = row[index]
                return "" if cell is None or cell.value is None else cell.value

            data = []
            for row in rows:
                code = core.normalize_code(cell_value(row, code_index))
                if not code:
                    continue
                values = []
                colors = []
                has_value = False
                for index, _name in tag_columns:
                    value = str(cell_value(row, index) or "").strip()
                    cell_color = ""
                    if value:
                        cell = row[index] if index < len(row) else None
                        cell_color = (
                            self._cell_color(cell) if cell is not None else ""
                        )
                        has_value = True
                    values.append(value)
                    colors.append(cell_color)
                if has_value:
                    data.append((code, values, colors))
            if not data:
                raise ValueError("Excel 里没有有效的 SKU + 标签数据")
        finally:
            workbook.close()

        file_hash = ""
        try:
            digest = hashlib.sha256()
            with open(path, "rb") as handle:
                for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                    digest.update(chunk)
            file_hash = digest.hexdigest()
        except OSError:
            pass

        source_name = os.path.basename(path)
        stamp = time.strftime("%Y%m%d_%H%M%S")
        safe_name = re.sub(r"[^\w.\-]+", "_", source_name)
        source_copy = os.path.join(MODULE_SOURCE_DIR, f"{stamp}_{safe_name}")
        try:
            shutil.copy2(path, source_copy)
        except OSError:
            source_copy = ""

        return self._replace_module(
            module_id=module_id,
            name=module_name,
            description=description,
            color=color or MODULE_COLORS[0],
            columns=[name for _index, name in tag_columns],
            data=data,
            source_name=source_name,
            source_path=source_copy,
            file_hash=file_hash,
        )

    def _replace_module(
        self,
        module_id,
        name,
        description,
        color,
        columns,
        data,
        source_name,
        source_path,
        file_hash,
    ):
        user = os.environ.get("USERNAME", "")
        created = module_id is None
        with self._lock:
            self.conn.execute("BEGIN IMMEDIATE")
            try:
                if module_id is None:
                    row = self.conn.execute(
                        "SELECT id FROM module WHERE name = ?", (name,)
                    ).fetchone()
                    if row:
                        module_id = int(row["id"])
                        created = False
                    else:
                        cursor = self.conn.execute(
                            "INSERT INTO module(name, description, color, priority, "
                            "status, created_at, updated_at) VALUES(?, ?, ?, 100, "
                            "'active', ?, ?)",
                            (name, description, color, now_text(), now_text()),
                        )
                        module_id = int(cursor.lastrowid)
                else:
                    row = self.conn.execute(
                        "SELECT id FROM module WHERE id = ?", (module_id,)
                    ).fetchone()
                    if not row:
                        raise ValueError("模块不存在")

                action_cursor = self.conn.execute(
                    "INSERT INTO action(kind, module_id, source_name, file_hash, "
                    "user, created_at, summary, status) "
                    "VALUES('import', ?, ?, ?, ?, ?, '', 'applied')",
                    (module_id, source_name, file_hash, user, now_text()),
                )
                action_id = int(action_cursor.lastrowid)

                self.conn.execute(
                    "DELETE FROM module_value WHERE module_id = ?", (module_id,)
                )
                self.conn.execute(
                    "DELETE FROM module_column WHERE module_id = ?", (module_id,)
                )

                column_ids = {}
                for position, column_name in enumerate(columns):
                    cursor = self.conn.execute(
                        "INSERT INTO module_column(module_id, name, position) "
                        "VALUES(?, ?, ?)",
                        (module_id, column_name, position),
                    )
                    column_ids[column_name] = int(cursor.lastrowid)

                batch = []
                for code, values, colors in data:
                    for position, value in enumerate(values):
                        if not value:
                            continue
                        column_name = columns[position]
                        cell_color = (
                            colors[position] if position < len(colors) else ""
                        )
                        batch.append(
                            (
                                module_id,
                                column_ids[column_name],
                                code,
                                value,
                                cell_color,
                                now_text(),
                            )
                        )
                    if len(batch) >= 2000:
                        self.conn.executemany(
                            "INSERT OR REPLACE INTO module_value(module_id, "
                            "column_id, code, value, color, updated_at) "
                            "VALUES(?, ?, ?, ?, ?, ?)",
                            batch,
                        )
                        batch.clear()
                if batch:
                    self.conn.executemany(
                        "INSERT OR REPLACE INTO module_value(module_id, "
                        "column_id, code, value, color, updated_at) "
                        "VALUES(?, ?, ?, ?, ?, ?)",
                        batch,
                    )

                self.conn.execute(
                    "UPDATE module SET name = ?, description = ?, color = ?, "
                    "status = 'active', updated_at = ? WHERE id = ?",
                    (name, description, color, now_text(), module_id),
                )

                version_no = self.conn.execute(
                    "SELECT COALESCE(MAX(version_no), 0) + 1 "
                    "FROM module_version WHERE module_id = ?",
                    (module_id,),
                ).fetchone()[0]
                version_cursor = self.conn.execute(
                    "INSERT INTO module_version(module_id, version_no, action_id, "
                    "source_name, file_hash, row_count, created_at, status) "
                    "VALUES(?, ?, ?, ?, ?, ?, ?, 'active')",
                    (
                        module_id,
                        version_no,
                        action_id,
                        source_name,
                        file_hash,
                        len({code for code, _values, _colors in data}),
                        now_text(),
                    ),
                )
                version_id = int(version_cursor.lastrowid)
                self.conn.execute(
                    "INSERT INTO module_version_column(version_id, name, position) "
                    "SELECT ?, name, position FROM module_column WHERE module_id = ?",
                    (version_id, module_id),
                )
                self.conn.execute(
                    "INSERT INTO module_version_value(version_id, column_name, "
                    "code, value, color) "
                    "SELECT ?, c.name, v.code, v.value, v.color "
                    "FROM module_value v "
                    "JOIN module_column c ON c.id = v.column_id "
                    "WHERE v.module_id = ?",
                    (version_id, module_id),
                )
                self.conn.execute(
                    "UPDATE action SET summary = ? WHERE id = ?",
                    (
                        f"导入 {len(data):,} 行 · {len(columns)} 列"
                        + (" · 新建模块" if created else " · 更新模块"),
                        action_id,
                    ),
                )
                self.conn.execute("COMMIT")
                return module_id
            except Exception:
                self.conn.execute("ROLLBACK")
                raise

    def set_module_color(self, module_id, color):
        with self._lock:
            self.conn.execute("BEGIN IMMEDIATE")
            try:
                self.conn.execute(
                    "UPDATE module SET color = ?, updated_at = ? WHERE id = ?",
                    (color, now_text(), module_id),
                )
                self.conn.execute(
                    "INSERT INTO action(kind, module_id, user, created_at, "
                    "summary, status) VALUES('color', ?, ?, ?, ?, 'applied')",
                    (
                        module_id,
                        os.environ.get("USERNAME", ""),
                        now_text(),
                        "修改模块颜色",
                    ),
                )
                self.conn.execute("COMMIT")
            except Exception:
                self.conn.execute("ROLLBACK")
                raise

    def toggle_module(self, module_id):
        module = self.get_module(module_id)
        if not module:
            return
        new_status = "disabled" if module["status"] == "active" else "active"
        with self._lock:
            self.conn.execute("BEGIN IMMEDIATE")
            try:
                self.conn.execute(
                    "UPDATE module SET status = ?, updated_at = ? WHERE id = ?",
                    (new_status, now_text(), module_id),
                )
                self.conn.execute(
                    "INSERT INTO action(kind, module_id, user, created_at, "
                    "summary, status) VALUES('toggle', ?, ?, ?, ?, 'applied')",
                    (
                        module_id,
                        os.environ.get("USERNAME", ""),
                        now_text(),
                        "启用模块" if new_status == "active" else "停用模块",
                    ),
                )
                self.conn.execute("COMMIT")
            except Exception:
                self.conn.execute("ROLLBACK")
                raise
        return new_status

    def delete_module(self, module_id):
        module = self.get_module(module_id)
        if not module:
            return
        with self._lock:
            self.conn.execute("BEGIN IMMEDIATE")
            try:
                self.conn.execute(
                    "UPDATE module SET status = 'deleted', updated_at = ? "
                    "WHERE id = ?",
                    (now_text(), module_id),
                )
                self.conn.execute(
                    "INSERT INTO action(kind, module_id, user, created_at, "
                    "summary, status) VALUES('delete', ?, ?, ?, ?, 'applied')",
                    (
                        module_id,
                        os.environ.get("USERNAME", ""),
                        now_text(),
                        f"删除模块「{module['name']}」",
                    ),
                )
                self.conn.execute("COMMIT")
            except Exception:
                self.conn.execute("ROLLBACK")
                raise

    def undo_latest(self, module_id):
        module = self.get_module(module_id)
        if not module:
            raise ValueError("模块不存在")
        with self._lock:
            latest = self.conn.execute(
                "SELECT * FROM module_version WHERE module_id = ? "
                "ORDER BY version_no DESC LIMIT 1",
                (module_id,),
            ).fetchone()
            if latest is None:
                raise ValueError("这个模块还没有导入版本")
            if latest["status"] != "active":
                raise ValueError("最近一次导入已经撤回过了")
            previous = self.conn.execute(
                "SELECT * FROM module_version WHERE module_id = ? "
                "AND version_no < ? ORDER BY version_no DESC LIMIT 1",
                (module_id, latest["version_no"]),
            ).fetchone()

            self.conn.execute("BEGIN IMMEDIATE")
            try:
                undo_cursor = self.conn.execute(
                    "INSERT INTO action(kind, module_id, target_action_id, "
                    "user, created_at, summary, status) "
                    "VALUES('undo', ?, ?, ?, ?, ?, 'applied')",
                    (
                        module_id,
                        latest["action_id"],
                        os.environ.get("USERNAME", ""),
                        now_text(),
                        f"撤回 v{latest['version_no']} 导入",
                    ),
                )
                undo_action_id = int(undo_cursor.lastrowid)

                self.conn.execute(
                    "DELETE FROM module_value WHERE module_id = ?", (module_id,)
                )
                self.conn.execute(
                    "DELETE FROM module_column WHERE module_id = ?", (module_id,)
                )
                if previous is not None:
                    column_rows = self.conn.execute(
                        "SELECT name, position FROM module_version_column "
                        "WHERE version_id = ? ORDER BY position",
                        (previous["id"],),
                    ).fetchall()
                    column_ids = {}
                    for row in column_rows:
                        cursor = self.conn.execute(
                            "INSERT INTO module_column(module_id, name, position) "
                            "VALUES(?, ?, ?)",
                            (module_id, row["name"], row["position"]),
                        )
                        column_ids[row["name"]] = int(cursor.lastrowid)
                    value_rows = self.conn.execute(
                        "SELECT column_name, code, value, color "
                        "FROM module_version_value WHERE version_id = ?",
                        (previous["id"],),
                    ).fetchall()
                    self.conn.executemany(
                        "INSERT OR REPLACE INTO module_value(module_id, column_id, "
                        "code, value, color, updated_at) VALUES(?, ?, ?, ?, ?, ?)",
                        [
                            (
                                module_id,
                                column_ids[row["column_name"]],
                                row["code"],
                                row["value"],
                                row["color"] or "",
                                now_text(),
                            )
                            for row in value_rows
                            if row["column_name"] in column_ids
                        ],
                    )
                    self.conn.execute(
                        "UPDATE module SET status = 'active', updated_at = ? "
                        "WHERE id = ?",
                        (now_text(), module_id),
                    )
                else:
                    self.conn.execute("DELETE FROM module WHERE id = ?", (module_id,))

                self.conn.execute(
                    "UPDATE module_version SET status = 'undone' WHERE id = ?",
                    (latest["id"],),
                )
                if latest["action_id"]:
                    self.conn.execute(
                        "UPDATE action SET status = 'undone' WHERE id = ?",
                        (latest["action_id"],),
                    )
                self.conn.execute(
                    "UPDATE action SET summary = ? WHERE id = ?",
                    (
                        f"撤回 v{latest['version_no']} 导入"
                        + (f"，恢复到 v{previous['version_no']}" if previous else "，模块已移除"),
                        undo_action_id,
                    ),
                )
                self.conn.execute("COMMIT")
            except Exception:
                self.conn.execute("ROLLBACK")
                raise
        return previous["version_no"] if previous else None

    def versions(self, module_id):
        with self._lock:
            rows = self.conn.execute(
                "SELECT * FROM module_version WHERE module_id = ? "
                "ORDER BY version_no DESC",
                (module_id,),
            ).fetchall()
        return [dict(row) for row in rows]

    def list_actions(self, limit=20):
        with self._lock:
            rows = self.conn.execute(
                "SELECT * FROM action ORDER BY id DESC LIMIT ?", (limit,)
            ).fetchall()
        return [dict(row) for row in rows]

    def export_module(self, module_id, path, sku_store):
        module = self.get_module(module_id)
        if not module:
            raise ValueError("模块不存在")
        columns, rows, _total = self.detail_rows(module_id, limit=1000000)
        load_workbook = core._import_openpyxl  # only to ensure openpyxl available
        from openpyxl import Workbook

        workbook = Workbook(write_only=True)
        sheet = workbook.create_sheet("标签数据")
        headers = ["SKU编码", "SKU名称（参考）"] + [column["name"] for column in columns]
        sheet.append(headers)
        for row in rows:
            code = row["code"]
            entry = sku_store.get(code) if sku_store else None
            name = entry["name"] if entry else ""
            sheet.append(
                [code, name]
                + [row["values"].get(column["name"], "") for column in columns]
            )
        workbook.save(path)
        return len(rows)


class StagApp:
    def __init__(self):
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            pass
        self.dpi_scale = core.App._dpi_scale()
        self.cfg = core.load_config()
        enable_dark_app_mode()
        self.root = tk.Tk()
        self.root.title(APP_TITLE)
        self.root.configure(bg=core.COLOR_BG)
        self.root.overrideredirect(True)
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        window_w, window_h = 1280, 800
        pos_x = max(0, (screen_w - window_w) // 2)
        pos_y = max(0, (screen_h - window_h) // 2 - 20)
        self.root.geometry(f"{window_w}x{window_h}+{pos_x}+{pos_y}")
        self.root.minsize(1080, 680)
        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)
        self.root.option_add("*Background", core.COLOR_BG)
        self.root.option_add("*Foreground", core.COLOR_TEXT)
        self._maximized = False
        self._normal_geometry = None
        self._drag_offset = None
        self._set_icon()
        self._init_ttk()

        self._callbacks = queue.Queue()
        self.root.after(50, self._drain_callbacks)

        self.store = core.DataStore(self.cfg, self._on_data_status)
        self.ops_store = core.OperationsStore(
            self.cfg.get("operations", {}),
            lambda text: self._on_data_status(text, "info"),
        )
        self.module_store = ModuleStore()
        self.ocr = core.OcrEnginePool()
        self.ocr_executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
        self.popup = core.Popup(self.root, self.cfg)
        self.popup.batch_callback = self.open_batch

        try:
            self.hotkey = core.parse_hotkey(self.cfg["trigger"]["hotkey"])
        except Exception:
            self.hotkey = None
        self.hotkey_state = False
        self.icon = None
        self._quitting = False
        self._matrix_after = None
        self._status_level = "info"

        self._build_topbar()
        self._build_pages()
        self.show_page("query")
        self.root.after(150, self._add_taskbar_style)

        self.store.start()
        self.ops_store.start()
        self._start_tray()
        self.hover = core.HoverController(self)
        self.hover.start()
        self.root.after(100, self._hotkey_tick)
        self.root.after(500, self._check_show_signal)
        self.root.after(300, self.refresh_matrix)

        if not self.ocr.available:
            self._on_data_status(
                "OCR 组件或语言包不可用，悬停不可用；F8 和手动查询仍可使用", "warn"
            )

    # ---------- base ----------

    def _set_icon(self):
        icon = core.resource_path(os.path.join("assets", "icon.ico"))
        if os.path.exists(icon):
            try:
                self.root.iconbitmap(icon)
            except Exception:
                pass

    def _init_ttk(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        for tree_style in ("Stag.Treeview", "Stag.Flat.Treeview"):
            style.configure(
                tree_style,
                background=core.COLOR_PANEL,
                fieldbackground=core.COLOR_PANEL,
                foreground=core.COLOR_TEXT,
                rowheight=24,
                borderwidth=0,
                bordercolor=core.COLOR_PANEL,
                lightcolor=core.COLOR_PANEL,
                darkcolor=core.COLOR_PANEL,
                relief="flat",
                font=(core.FONT_UI, 9),
            )
            style.map(
                tree_style,
                background=[("selected", core.COLOR_DIM_CYAN)],
                foreground=[("selected", core.COLOR_TEXT)],
            )
        # The detail table draws its own header; remove the themed field border
        # entirely so no light 1px frame leaks around the data area.
        style.layout(
            "Stag.Flat.Treeview",
            [("Treeview.treearea", {"sticky": "nswe"})],
        )
        style.configure(
            "Stag.Treeview.Heading",
            background=core.COLOR_PANEL_2,
            foreground=core.COLOR_CYAN,
            relief="flat",
            borderwidth=0,
            padding=(6, 6),
            font=(core.FONT_UI, 9, "bold"),
        )
        style.map(
            "Stag.Treeview.Heading",
            background=[
                ("active", core.COLOR_PANEL_2),
                ("pressed", core.COLOR_PANEL_2),
            ],
            foreground=[
                ("active", core.COLOR_CYAN),
                ("pressed", core.COLOR_CYAN),
            ],
            relief=[("active", "flat"), ("pressed", "flat")],
        )
        for scroll_style in (
            "Stag.Vertical.TScrollbar",
            "Stag.Horizontal.TScrollbar",
        ):
            style.configure(
                scroll_style,
                background=core.COLOR_PANEL_2,
                troughcolor=core.COLOR_BG,
                bordercolor=core.COLOR_LINE,
                lightcolor=core.COLOR_LINE,
                darkcolor=core.COLOR_LINE,
                arrowcolor=core.COLOR_MUTED,
                relief="flat",
                gripcount=0,
            )
            style.map(
                scroll_style,
                background=[
                    ("disabled", core.COLOR_PANEL_2),
                    ("pressed", core.COLOR_DIM_CYAN),
                    ("active", "#152238"),
                ],
                troughcolor=[("disabled", core.COLOR_BG)],
                arrowcolor=[
                    ("disabled", "#3A4A5E"),
                    ("active", core.COLOR_CYAN),
                ],
                lightcolor=[
                    ("disabled", core.COLOR_LINE),
                    ("active", core.COLOR_LINE),
                ],
                darkcolor=[
                    ("disabled", core.COLOR_LINE),
                    ("active", core.COLOR_LINE),
                ],
                bordercolor=[("disabled", core.COLOR_LINE)],
            )

    def _label(self, parent, text, fg=core.COLOR_TEXT, size=10, bold=False, bg=None):
        return tk.Label(
            parent,
            text=text,
            bg=bg or core.COLOR_BG,
            fg=fg,
            font=(core.FONT_UI, size, "bold" if bold else "normal"),
            anchor="w",
            justify="left",
        )

    def _button(self, parent, text, command, accent=False, small=False):
        return HudButton(parent, text, command, accent=accent, small=small)

    def _entry(self, parent, textvariable=None, width=40):
        return tk.Entry(
            parent,
            textvariable=textvariable,
            width=width,
            bg=core.COLOR_INPUT,
            fg=core.COLOR_TEXT,
            insertbackground=core.COLOR_CYAN,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=core.COLOR_LINE_BRIGHT,
            highlightcolor=core.COLOR_CYAN,
            font=(core.FONT_UI, 10),
        )

    # ---------- topbar/pages ----------

    def _build_topbar(self):
        bar = tk.Frame(self.root, bg=core.COLOR_BG, height=64)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        brand = tk.Frame(bar, bg=core.COLOR_BG)
        brand.pack(side="left", padx=(24, 0), pady=1)
        logo_path = core.resource_path(os.path.join("assets", "stag_logo.png"))
        self.brand_photo = None
        if os.path.exists(logo_path):
            try:
                logo = Image.open(logo_path).convert("RGBA")
                logo = logo.resize((100, 42), Image.LANCZOS)
                self.brand_photo = ImageTk.PhotoImage(logo)
                tk.Label(
                    brand, image=self.brand_photo, bg=core.COLOR_BG, bd=0,
                ).pack(anchor="w")
            except Exception:
                self.brand_photo = None
        if self.brand_photo is None:
            word = tk.Frame(brand, bg=core.COLOR_BG)
            word.pack(anchor="w")
            tk.Label(
                word, text="S", bg=core.COLOR_BG, fg=core.COLOR_CYAN,
                font=(core.FONT_DISPLAY, 21, "bold"),
            ).pack(side="left")
            tk.Label(
                word, text="Tag", bg=core.COLOR_BG, fg=core.COLOR_TEXT,
                font=(core.FONT_DISPLAY, 21, "bold"),
            ).pack(side="left")
            tk.Label(
                brand, text="SKU TAG SYSTEM", bg=core.COLOR_BG,
                fg=core.COLOR_DIM_CYAN, font=(core.FONT_MONO, 8),
            ).pack(anchor="w", pady=(2, 0))

        nav = tk.Frame(bar, bg=core.COLOR_BG)
        nav.pack(side="left", padx=40)
        self.nav_query = tk.Button(
            nav, text="查询页", command=lambda: self.show_page("query"),
            bg=core.COLOR_BG, fg=core.COLOR_CYAN, activebackground=core.COLOR_BG,
            activeforeground=core.COLOR_CYAN, bd=0, relief="flat",
            font=(core.FONT_UI, 10, "bold"), padx=10, pady=8, cursor="hand2",
        )
        self.nav_query.pack(side="left")
        self.nav_settings = tk.Button(
            nav, text="设置页面", command=lambda: self.show_page("settings"),
            bg=core.COLOR_BG, fg=core.COLOR_MUTED, activebackground=core.COLOR_BG,
            activeforeground=core.COLOR_TEXT, bd=0, relief="flat",
            font=(core.FONT_UI, 10), padx=10, pady=8, cursor="hand2",
        )
        self.nav_settings.pack(side="left")

        right = tk.Frame(bar, bg=core.COLOR_BG)
        right.pack(side="right", padx=(12, 18))
        controls = tk.Frame(right, bg=core.COLOR_BG)
        controls.pack(side="right", padx=(12, 0))
        for text, command, hover in (
            ("—", self.hide_window, core.COLOR_CYAN),
            ("□", self.toggle_maximize, core.COLOR_CYAN),
            ("×", self.hide_window, core.COLOR_MAGENTA),
        ):
            button = tk.Label(
                controls, text=text, bg=core.COLOR_BG, fg=core.COLOR_MUTED,
                font=(core.FONT_DISPLAY, 13), width=3, cursor="hand2",
            )
            button.pack(side="left")
            button.bind("<Enter>", lambda _e, w=button, c=hover: w.configure(fg=c))
            button.bind("<Leave>", lambda _e, w=button: w.configure(fg=core.COLOR_MUTED))
            button.bind("<Button-1>", lambda _e, c=command: c())

        self.status_var = tk.StringVar(value="● 同步中")
        self.status_label = tk.Label(
            right, textvariable=self.status_var, bg=core.COLOR_BG,
            fg=core.COLOR_MUTED, font=(core.FONT_MONO, 9),
        )
        self.status_label.pack(side="right", padx=(0, 10))
        tk.Frame(bar, bg=core.COLOR_LINE, height=1).place(
            relx=0, rely=1, relwidth=1, anchor="sw"
        )

        for widget in (bar, brand, nav, self.status_label):
            widget.bind("<ButtonPress-1>", self._start_drag)
            widget.bind("<B1-Motion>", self._drag_window)
            widget.bind("<ButtonRelease-1>", self._end_drag)
        for child in brand.winfo_children():
            child.bind("<ButtonPress-1>", self._start_drag)
            child.bind("<B1-Motion>", self._drag_window)
            child.bind("<ButtonRelease-1>", self._end_drag)
        bar.bind("<Double-Button-1>", lambda _event: self.toggle_maximize())

        self.resize_grip = tk.Canvas(
            self.root, width=16, height=16, bg=core.COLOR_BG,
            highlightthickness=0, bd=0, cursor="size_nw_se",
        )
        self.resize_grip.create_line(4, 14, 14, 4, fill=core.COLOR_LINE_BRIGHT)
        self.resize_grip.create_line(9, 14, 14, 9, fill=core.COLOR_DIM_CYAN)
        self.resize_grip.place(relx=1.0, rely=1.0, anchor="se")
        self.resize_grip.bind("<ButtonPress-1>", self._resize_start)
        self.resize_grip.bind("<B1-Motion>", self._resize_drag)

    def _start_drag(self, event):
        self._drag_offset = (
            event.x_root - self.root.winfo_x(),
            event.y_root - self.root.winfo_y(),
        )

    def _drag_window(self, event):
        if not self._drag_offset:
            return
        x = event.x_root - self._drag_offset[0]
        y = event.y_root - self._drag_offset[1]
        self.root.geometry(f"+{x}+{y}")

    def _end_drag(self, _event=None):
        self._drag_offset = None

    def _resize_start(self, event):
        self._resize_origin = (
            event.x_root,
            event.y_root,
            self.root.winfo_width(),
            self.root.winfo_height(),
        )

    def _resize_drag(self, event):
        origin = getattr(self, "_resize_origin", None)
        if not origin:
            return
        start_x, start_y, start_w, start_h = origin
        width = max(1080, start_w + event.x_root - start_x)
        height = max(680, start_h + event.y_root - start_y)
        self.root.geometry(f"{width}x{height}")

    def _work_area(self):
        class RECT(ctypes.Structure):
            _fields_ = [
                ("left", ctypes.c_long),
                ("top", ctypes.c_long),
                ("right", ctypes.c_long),
                ("bottom", ctypes.c_long),
            ]

        rect = RECT()
        try:
            ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(rect), 0)
            return rect.left, rect.top, rect.right, rect.bottom
        except Exception:
            return 0, 0, self.root.winfo_screenwidth(), self.root.winfo_screenheight()

    def toggle_maximize(self):
        if self._maximized:
            if self._normal_geometry:
                self.root.geometry(self._normal_geometry)
            self._maximized = False
            return
        self._normal_geometry = self.root.geometry()
        left, top, right, bottom = self._work_area()
        self.root.geometry(f"{right - left}x{bottom - top}+{left}+{top}")
        self._maximized = True

    def _add_taskbar_style(self):
        try:
            hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id()) or self.root.winfo_id()
            ex_style = ctypes.windll.user32.GetWindowLongW(hwnd, -20)
            ex_style = (ex_style & ~0x00000080) | 0x00040000
            ctypes.windll.user32.SetWindowLongW(hwnd, -20, ex_style)
            ctypes.windll.user32.SetWindowPos(
                hwnd, 0, 0, 0, 0, 0, 0x0020 | 0x0001 | 0x0002 | 0x0004
            )
        except Exception:
            pass

    def _build_pages(self):
        self.pages = tk.Frame(self.root, bg=core.COLOR_BG)
        self.pages.pack(fill="both", expand=True, padx=20, pady=(0, 18))
        self.query_page = tk.Frame(self.pages, bg=core.COLOR_BG)
        self.settings_page = tk.Frame(self.pages, bg=core.COLOR_BG)
        self._build_query_page()
        self._build_settings_page()
        try:
            self.resize_grip.lift()
        except Exception:
            pass

    def show_page(self, page):
        self.query_page.pack_forget()
        self.settings_page.pack_forget()
        if page == "settings":
            self.settings_page.pack(fill="both", expand=True)
            self.nav_settings.configure(fg=core.COLOR_CYAN)
            self.nav_query.configure(fg=core.COLOR_MUTED)
            self._sync_settings_from_cfg()
            self.root.bind_all("<MouseWheel>", self._settings_wheel)
        else:
            self.query_page.pack(fill="both", expand=True)
            self.nav_query.configure(fg=core.COLOR_CYAN)
            self.nav_settings.configure(fg=core.COLOR_MUTED)
            self.root.unbind_all("<MouseWheel>")
            self.refresh_matrix()

    def _settings_wheel(self, event):
        try:
            if self.settings_page.winfo_ismapped():
                self.settings_canvas.yview_scroll(
                    int(-event.delta / 120), "units"
                )
                return "break"
        except Exception:
            pass

    # ---------- query page ----------

    def _build_query_page(self):
        page = self.query_page
        page.grid_rowconfigure(1, weight=1)
        page.grid_columnconfigure(0, weight=1)
        page.grid_columnconfigure(1, weight=0)

        head = tk.Frame(page, bg=core.COLOR_BG)
        head.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(4, 14))
        left = tk.Frame(head, bg=core.COLOR_BG)
        left.pack(side="left")
        tk.Label(
            left, text="QUERY CONSOLE", bg=core.COLOR_BG, fg=core.COLOR_CYAN,
            font=(core.FONT_MONO, 8),
        ).pack(anchor="w")
        tk.Label(
            left, text="SKU 查询", bg=core.COLOR_BG,
            fg=core.COLOR_TEXT, font=(core.FONT_UI, 14, "bold"),
        ).pack(anchor="w", pady=(5, 0))

        actions = tk.Frame(
            head, bg=core.COLOR_PANEL,
            highlightthickness=1, highlightbackground=core.COLOR_LINE,
        )
        actions.pack(side="right", anchor="s")
        actions_inner = tk.Frame(actions, bg=core.COLOR_PANEL)
        actions_inner.pack(padx=9, pady=7)
        self.query_var = tk.StringVar()
        entry = self._entry(actions_inner, self.query_var, width=34)
        entry.pack(side="left", ipady=7, padx=(0, 8))
        entry.bind("<Return>", lambda _event: self.query_current())
        self._button(actions_inner, "查询", self.query_current, accent=True).pack(side="left")
        self._button(actions_inner, "批量查询", self.open_batch).pack(side="left", padx=(8, 0))
        self.query_entry = entry

        matrix_wrap = tk.Frame(page, bg=core.COLOR_BG)
        matrix_wrap.grid(row=1, column=0, sticky="nsew", padx=(0, 12))
        matrix_wrap.grid_rowconfigure(1, weight=1)
        matrix_wrap.grid_columnconfigure(0, weight=1)

        matrix_head = tk.Frame(matrix_wrap, bg=core.COLOR_BG)
        matrix_head.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        tk.Label(
            matrix_head, text="STAG MATRIX / 模块矩阵", bg=core.COLOR_BG,
            fg=core.COLOR_TEXT, font=(core.FONT_UI, 13, "bold"),
        ).pack(side="left")
        self.matrix_stats = tk.Label(
            matrix_head, text="", bg=core.COLOR_BG, fg=core.COLOR_MUTED,
            font=(core.FONT_MONO, 8),
        )
        self.matrix_stats.pack(side="left", padx=(14, 0))
        self._button(matrix_head, "+ 新建模块", self.open_new_module, small=True).pack(
            side="right"
        )
        self._button(matrix_head, "导出模板", self.export_blank_template, small=True).pack(
            side="right", padx=(0, 6)
        )

        self.matrix_canvas = tk.Canvas(
            matrix_wrap, bg=core.COLOR_PANEL, highlightthickness=1,
            highlightbackground=core.COLOR_LINE_BRIGHT, bd=0,
        )
        self.matrix_canvas.grid(row=1, column=0, sticky="nsew")
        self.matrix_canvas.bind("<Configure>", self._on_matrix_resize)
        self.matrix_footer = tk.Label(
            matrix_wrap,
            text="面积 = 数据量（行 × 列，开方压缩） · 左键打开模块，右键停用/删除",
            bg=core.COLOR_BG, fg=core.COLOR_DIM_CYAN,
            font=(core.FONT_MONO, 8), anchor="w",
        )
        self.matrix_footer.grid(row=2, column=0, sticky="ew", pady=(6, 0))

        rail = tk.Frame(page, bg=core.COLOR_BG, width=286)
        rail.grid(row=1, column=1, sticky="ns")
        rail.pack_propagate(False)
        rail.grid_rowconfigure(0, weight=1)
        rail.grid_rowconfigure(1, weight=1)
        rail.grid_columnconfigure(0, weight=1)

        columns_card = tk.Frame(rail, bg=core.COLOR_PANEL, highlightthickness=1,
                                highlightbackground=core.COLOR_LINE)
        columns_card.grid(row=0, column=0, sticky="nsew", pady=(0, 8))
        tk.Label(
            columns_card, text="已有列", bg=core.COLOR_PANEL, fg=core.COLOR_MUTED,
            font=(core.FONT_MONO, 9), anchor="w", padx=12, pady=8,
        ).pack(fill="x")
        self.columns_text = tk.Text(
            columns_card, bg=core.COLOR_PANEL, fg=core.COLOR_MUTED, bd=0,
            highlightthickness=0, wrap="word", font=(core.FONT_UI, 9),
            padx=12, pady=4,
        )
        self.columns_text.pack(fill="both", expand=True)
        self.columns_text.configure(state="disabled")

        actions_card = tk.Frame(rail, bg=core.COLOR_PANEL, highlightthickness=1,
                                highlightbackground=core.COLOR_LINE)
        actions_card.grid(row=1, column=0, sticky="nsew")
        tk.Label(
            actions_card, text="最近动作", bg=core.COLOR_PANEL, fg=core.COLOR_MUTED,
            font=(core.FONT_MONO, 9), anchor="w", padx=12, pady=8,
        ).pack(fill="x")
        self.actions_text = tk.Text(
            actions_card, bg=core.COLOR_PANEL, fg=core.COLOR_MUTED, bd=0,
            highlightthickness=0, wrap="word", font=(core.FONT_UI, 9),
            padx=12, pady=4,
        )
        self.actions_text.pack(fill="both", expand=True)
        self.actions_text.configure(state="disabled")

    def _on_matrix_resize(self, _event=None):
        if self._matrix_after:
            try:
                self.root.after_cancel(self._matrix_after)
            except Exception:
                pass
        self._matrix_after = self.root.after(120, self.refresh_matrix)

    def refresh_matrix(self):
        canvas = self.matrix_canvas
        width = canvas.winfo_width()
        height = canvas.winfo_height()
        if width < 80 or height < 80:
            self.root.after(120, self.refresh_matrix)
            return
        modules = self.module_store.list_modules()
        canvas.delete("all")
        self._block_photos = {}
        # HUD 网格与四角框，给 Canvas 一点网页版的空间感
        for gx in range(18, width, 36):
            canvas.create_line(gx, 12, gx, height - 12, fill="#0B1220")
        for gy in range(18, height, 36):
            canvas.create_line(12, gy, width - 12, gy, fill="#0B1220")
        for cx, cy, dx, dy in (
            (8, 8, 1, 1), (width - 8, 8, -1, 1),
            (8, height - 8, 1, -1), (width - 8, height - 8, -1, -1),
        ):
            canvas.create_line(cx, cy, cx + 18 * dx, cy, fill=core.COLOR_DIM_CYAN)
            canvas.create_line(cx, cy, cx, cy + 18 * dy, fill=core.COLOR_DIM_CYAN)
        if not modules:
            canvas.create_text(
                width / 2, height / 2, text="还没有模块\n点击右上角「+ 新建模块」开始",
                fill=core.COLOR_MUTED, font=(core.FONT_UI, 12), justify="center",
            )
            self.matrix_stats.configure(text="0 个模块")
            self._refresh_rail(modules)
            return

        modules.sort(
            key=lambda m: (m["row_count"] * max(1, m["col_count"])),
            reverse=True,
        )
        items = [
            {
                "item": module,
                "weight": max(1.0, math.sqrt(module["row_count"] * max(1, module["col_count"]))),
            }
            for module in modules
        ]
        for module, x, y, w, h in treemap(items, width - 16, height - 16, gap=8):
            self._draw_block(canvas, module, x + 8, y + 8, w, h)

        total_rows = sum(module["row_count"] for module in modules)
        self.matrix_stats.configure(
            text=f"{len(modules)} 个模块 · {total_rows:,} 条模块记录"
        )
        self._refresh_rail(modules)

    def _draw_block(self, canvas, module, x, y, w, h):
        color = module["color"] or core.COLOR_CYAN
        disabled = module["status"] == "disabled"
        tag = f"module_{module['id']}"
        normal_photo = make_block_photo(w, h, color, disabled=disabled)
        hover_photo = make_block_photo(w, h, color, disabled=disabled, highlight=True)
        self._block_photos[(tag, "normal")] = normal_photo
        self._block_photos[(tag, "hover")] = hover_photo
        body_id = canvas.create_image(
            x, y, anchor="nw", image=normal_photo, tags=(tag,)
        )
        canvas.create_line(
            x + 1, y + 1, x + min(w, 110), y + 1,
            fill=color, width=2, tags=(tag,),
        )
        name = module["name"] + ("  [停用]" if disabled else "")
        canvas.create_text(
            x + 11, y + 9, text=name, anchor="nw", fill=core.COLOR_MUTED if disabled else core.COLOR_TEXT,
            font=(core.FONT_UI, 12 if w > 180 else 10, "bold"),
            width=max(40, w - 22), tags=(tag,),
        )
        if w > 120 and h > 58:
            columns = self.module_store.columns(module["id"])
            column_text = " / ".join(column["name"] for column in columns[:4])
            if len(columns) > 4:
                column_text += f"  +{len(columns) - 4}"
            canvas.create_text(
                x + 11, y + 37, text=column_text, anchor="nw",
                fill=core.COLOR_MUTED, font=(core.FONT_UI, 8),
                width=max(40, w - 22), tags=(tag,),
            )
        if w > 90:
            canvas.create_text(
                x + w - 9, y + 9, text=f"{module['row_count']:,}", anchor="ne",
                fill=color, font=(core.FONT_MONO, 8), tags=(tag,),
            )
        if h > 78:
            canvas.create_text(
                x + 11, y + h - 9, text=f"{module['col_count']} 列 · {module['updated_at'][5:16]}",
                anchor="sw", fill=core.COLOR_DIM_CYAN, font=(core.FONT_MONO, 8),
                tags=(tag,),
            )
        if not disabled:
            canvas.tag_bind(body_id, "<Enter>", lambda _event: (
                canvas.itemconfigure(body_id, image=hover_photo),
                canvas.tag_raise(tag),
            ))
            canvas.tag_bind(body_id, "<Leave>", lambda _event: (
                canvas.itemconfigure(body_id, image=normal_photo),
            ))
        canvas.tag_bind(tag, "<Button-1>", lambda _event, mid=module["id"]: self.open_module(mid))
        canvas.tag_bind(tag, "<Button-3>", lambda event, mid=module["id"]: self._module_menu(event, mid))

    def _module_menu(self, event, module_id):
        module = self.module_store.get_module(module_id)
        if not module:
            return
        menu = tk.Menu(
            self.root, tearoff=0, bg=core.COLOR_PANEL_2, fg=core.COLOR_TEXT,
            activebackground=core.COLOR_DIM_CYAN, activeforeground=core.COLOR_TEXT,
            bd=0,
        )
        menu.add_command(label="打开模块", command=lambda: self.open_module(module_id))
        menu.add_command(
            label="启用模块" if module["status"] == "disabled" else "停用模块",
            command=lambda: self._toggle_module(module_id),
        )
        menu.add_separator()
        menu.add_command(label="删除模块", command=lambda: self._delete_module(module_id))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _refresh_rail(self, modules):
        self.columns_text.configure(state="normal")
        self.columns_text.delete("1.0", "end")
        for index, module in enumerate(modules):
            tag = f"m{module['id']}"
            self.columns_text.tag_configure(
                tag, foreground=module["color"] or core.COLOR_CYAN,
                font=(core.FONT_UI, 9, "bold"),
            )
            self.columns_text.insert("end", module["name"] + "\n", tag)
            columns = self.module_store.columns(module["id"])
            text = "  " + " / ".join(column["name"] for column in columns) + "\n\n"
            self.columns_text.insert("end", text, "muted")
        self.columns_text.tag_configure("muted", foreground=core.COLOR_MUTED)
        self.columns_text.configure(state="disabled")

        self.actions_text.configure(state="normal")
        self.actions_text.delete("1.0", "end")
        for action in self.module_store.list_actions(12):
            created = (action["created_at"] or "")[5:16]
            summary = action["summary"] or action["kind"]
            self.actions_text.insert("end", f"{created}  {summary}\n", "action")
            self.actions_text.insert("end", f"          {action['kind']} · {action['status']}\n\n", "muted")
        self.actions_text.tag_configure("action", foreground=core.COLOR_TEXT)
        self.actions_text.tag_configure("muted", foreground=core.COLOR_DIM_CYAN)
        self.actions_text.configure(state="disabled")

    def query_current(self):
        code = core.normalize_code(self.query_var.get())
        if not code:
            self.query_entry.focus_set()
            return
        entry = self.store.mapping.get(code)
        if entry is None:
            self.popup.show("未找到匹配", [f"当前数据源里没有 {code}"], [("未找到", core.COLOR_MAGENTA)])
            return
        self._show_hit(code, entry)

    def _show_hit(self, code, entry, anchor=None):
        name = entry["name"]
        lines = [name]
        stock = entry.get("stock") or {}
        if self.cfg["display"].get("show_stock", True) and stock:
            lines.append("　".join(f"{key} {value}" for key, value in stock.items()))
        ops = self.ops_store.lookup(code)
        module_tags = self.module_store.lookup(code)
        self.popup.show(
            code, lines, core.detect_styles(name), anchor=anchor,
            split_name=True, ops=ops, module_tags=module_tags,
        )

    def open_module(self, module_id):
        ModuleWindow(self, module_id)

    def open_new_module(self):
        path = filedialog.askopenfilename(
            parent=self.root,
            title="选择模块 Excel",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            NewModuleWindow(self, path)

    def export_blank_template(self):
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title="导出空白标签模板",
            defaultextension=".xlsx",
            initialfile="STag_标签模板.xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not path:
            return
        try:
            from openpyxl import Workbook
            from openpyxl.comments import Comment
            from openpyxl.styles import Alignment, Font, PatternFill

            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "标签数据"
            # Only the two fixed columns are pre-filled. Tag columns are left
            # completely blank so users can name their own columns; the import
            # step must never invent 标签1 / 标签2 / 标签3.
            headers = ["SKU编码", "SKU名称（参考）"]
            sheet.append(headers)
            for column, width in zip("ABCDEF", (18, 72, 22, 22, 22, 22)):
                sheet.column_dimensions[column].width = width
            for cell in sheet[1][:2]:
                cell.fill = PatternFill("solid", fgColor="0A101C")
                cell.font = Font(
                    name="Microsoft YaHei", bold=True, color="35E7FF", size=10
                )
                cell.alignment = Alignment(horizontal="center", vertical="center")
            sheet["C1"].comment = Comment(
                "从 C1 开始填写你自己的标签列名，例如“它真的是肖肖吗”“testA”。"
                "不需要的列可以留空。",
                "STag",
            )
            sheet.freeze_panes = "C2"

            guide = workbook.create_sheet("填写说明")
            guide.column_dimensions["A"].width = 18
            guide.column_dimensions["B"].width = 92
            guide.append(["项目", "说明"])
            for row in (
                ("用途", "给 SKU 打自定义标签。导入时才填写模块名称。"),
                ("第一列", "SKU编码，必填；支持 9 位码和 1.01.01.xxxx 带点码。"),
                ("第二列", "SKU名称仅供参考，导入时忽略。"),
                ("第三列起", "自己在第一行填写标签列名，例如“它真的是肖肖吗”“testA”；不需要的列留空。"),
                ("表头", "第一行就是表头，不要在上面再加标题行，也不要合并单元格。"),
                ("每一行", "一个 SKU；数据行可以不填。"),
                ("空单元格", "不写入该标签。"),
            ):
                guide.append(list(row))
            workbook.save(path)
            messagebox.showinfo("模板已导出", path, parent=self.root)
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc), parent=self.root)

    def open_batch(self):
        BatchWindow(self)

    def _toggle_module(self, module_id):
        try:
            self.module_store.toggle_module(module_id)
            self.refresh_matrix()
        except Exception as exc:
            messagebox.showerror("操作失败", str(exc), parent=self.root)

    def _delete_module(self, module_id):
        module = self.module_store.get_module(module_id)
        if not module:
            return
        if not messagebox.askyesno(
            "删除模块",
            f"删除模块「{module['name']}」？\n\n模块会从查询页移除，历史版本和动作记录仍保留。",
            parent=self.root,
        ):
            return
        try:
            self.module_store.delete_module(module_id)
            self.refresh_matrix()
        except Exception as exc:
            messagebox.showerror("删除失败", str(exc), parent=self.root)

    # ---------- settings page ----------

    def _build_settings_page(self):
        page = self.settings_page
        canvas = tk.Canvas(page, bg=core.COLOR_BG, highlightthickness=0, bd=0)
        self.settings_canvas = canvas
        scrollbar = ttk.Scrollbar(
            page, orient="vertical", command=canvas.yview,
            style="Stag.Vertical.TScrollbar",
        )
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        content = tk.Frame(canvas, bg=core.COLOR_BG)
        window = canvas.create_window((0, 0), window=content, anchor="nw")
        content.bind(
            "<Configure>",
            lambda _event: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.bind(
            "<Configure>",
            lambda event: canvas.itemconfigure(window, width=event.width),
        )
        body = tk.Frame(content, bg=core.COLOR_BG)
        body.pack(fill="x", padx=18, pady=(4, 22))

        tk.Label(
            body, text="SETTINGS / 设置页面", bg=core.COLOR_BG,
            fg=core.COLOR_CYAN, font=(core.FONT_MONO, 9),
        ).pack(anchor="w", pady=(4, 4))
        tk.Label(
            body, text="把数据源和触发方式调顺。", bg=core.COLOR_BG,
            fg=core.COLOR_TEXT, font=(core.FONT_UI, 20),
        ).pack(anchor="w")
        tk.Label(
            body, text="保存后立即生效；数据源路径会写入共享 config.json，个人偏好写入本机。",
            bg=core.COLOR_BG, fg=core.COLOR_MUTED, font=(core.FONT_UI, 9),
        ).pack(anchor="w", pady=(6, 18))

        self.path_var = tk.StringVar()
        self.ops_path_var = tk.StringVar()
        data_body = self._settings_card(
            body, "数据源", "SKU 表与负责人关系表；支持本地路径和共享盘。程序会先复制到本地再读。"
        )
        row = 0
        row = self._settings_field(
            data_body, row, "SKU 数据源", self.path_var, self.browse_source,
            hint="例如 D:/data/inventory.xlsx",
        )
        row = self._settings_field(
            data_body, row, "负责人关系表", self.ops_path_var, self.browse_ops,
            hint="负责人和助理从这里读取；共享盘断开时使用本机数据库。",
        )
        sync_row = tk.Frame(data_body, bg=core.COLOR_PANEL)
        sync_row.grid(row=row, column=1, columnspan=2, sticky="w", pady=(3, 0))
        self._button(
            sync_row, "立即同步", self.reload_data, accent=True, small=True
        ).pack(side="left")
        tk.Label(
            sync_row, text="立即重新读取 SKU 和负责人关系表",
            bg=core.COLOR_PANEL, fg=core.COLOR_DIM_CYAN,
            font=(core.FONT_UI, 8),
        ).pack(side="left", padx=(9, 0))

        self.hover_var = tk.BooleanVar()
        self.hotkey_var = tk.BooleanVar()
        trigger_body = self._settings_card(
            body, "触发方式", "两个开关互相独立，可以同时打开；都关掉就是完全停用。"
        )
        row = 0
        row = self._settings_switch(
            trigger_body, row, "鼠标悬停", "光标停在编码上约 0.6 秒自动弹出。",
            self.hover_var, core.COLOR_CYAN,
        )
        row = self._settings_switch(
            trigger_body, row, "选中 + 热键", "选中编码后按热键查询，适合后台/RPA 卡片。",
            self.hotkey_var, core.COLOR_MAGENTA,
        )
        self.hotkey_value_var = tk.StringVar()
        self.dwell_var = tk.StringVar()
        row = self._settings_field(
            trigger_body, row, "热键", self.hotkey_value_var,
            hint="默认 F8，也可以写 Ctrl+Alt+S。",
        )
        row = self._settings_field(
            trigger_body, row, "悬停停留（毫秒）", self.dwell_var,
            hint="越大越不容易误触，默认 600。",
        )

        self.font_var = tk.StringVar()
        self.opacity_var = tk.StringVar()
        self.auto_close_var = tk.StringVar()
        self.show_module_name_var = tk.BooleanVar()
        self.show_stock_var = tk.BooleanVar()
        display_body = self._settings_card(
            body, "弹窗显示", "控制查码 HUD 的阅读体验；透明度和字号会立即生效。"
        )
        row = 0
        row = self._settings_field(display_body, row, "弹窗字号", self.font_var)
        row = self._settings_field(
            display_body, row, "弹窗透明度（0.3–1.0）", self.opacity_var,
            hint="0.98 接近不透明，0.6 会更轻。",
        )
        row = self._settings_field(
            display_body, row, "自动关闭（秒，0=不关）", self.auto_close_var,
            hint="鼠标移开后也会自动关闭。",
        )
        row = self._settings_switch(
            display_body, row, "显示模块名称",
            "关闭后，弹窗只显示列名和值，不显示“肖肖？”这类模块名。",
            self.show_module_name_var, core.COLOR_CYAN,
        )
        row = self._settings_switch(
            display_body, row, "显示库存", "在名称下面显示库存A和库存B。",
            self.show_stock_var, core.COLOR_AMBER,
        )

        self.startup_var = tk.BooleanVar()
        startup_body = self._settings_card(
            body, "启动与模块库", "模块库当前保存在本机 SQLite；后续可切换共享库。"
        )
        row = 0
        row = self._settings_switch(
            startup_body, row, "开机自动启动", "登录 Windows 后自动缩到托盘。",
            self.startup_var, "#59FFA8",
        )
        self._settings_info(startup_body, row, "模块库位置", MODULE_DB_PATH)
        self._settings_info(startup_body, row + 1, "模块源文件副本", MODULE_SOURCE_DIR)

        buttons = tk.Frame(body, bg=core.COLOR_BG)
        buttons.pack(fill="x", pady=(20, 0))
        self._button(buttons, "保存并应用", self.save_settings, accent=True).pack(side="left")
        self._button(buttons, "立即重新读取", self.reload_data).pack(side="left", padx=(9, 0))

        self.settings_status = tk.Label(
            body, text="", bg=core.COLOR_BG, fg=core.COLOR_MUTED,
            font=(core.FONT_UI, 9),
        )
        self.settings_status.pack(anchor="w", pady=(12, 0))

        self._sync_settings_from_cfg()

    def _settings_card(self, parent, title, hint):
        card = tk.Frame(
            parent, bg=core.COLOR_PANEL,
            highlightthickness=1, highlightbackground=core.COLOR_LINE,
        )
        card.pack(fill="x", pady=(0, 12))
        inner = tk.Frame(card, bg=core.COLOR_PANEL)
        inner.pack(fill="x", padx=16, pady=14)
        tk.Label(
            inner, text=title, bg=core.COLOR_PANEL, fg=core.COLOR_TEXT,
            font=(core.FONT_UI, 12, "bold"),
        ).pack(anchor="w")
        tk.Label(
            inner, text=hint, bg=core.COLOR_PANEL, fg=core.COLOR_MUTED,
            font=(core.FONT_UI, 8), justify="left", wraplength=900,
        ).pack(anchor="w", pady=(4, 10))
        body = tk.Frame(inner, bg=core.COLOR_PANEL)
        body.pack(fill="x")
        body.grid_columnconfigure(1, weight=1)
        return body

    def _settings_field(self, parent, row, label, variable, browse_command=None, hint=None):
        tk.Label(
            parent, text=label, bg=core.COLOR_PANEL, fg=core.COLOR_MUTED,
            font=(core.FONT_UI, 9),
        ).grid(row=row, column=0, sticky="nw", pady=6, padx=(0, 14))
        entry = self._entry(parent, variable, width=56)
        entry.grid(row=row, column=1, sticky="ew", pady=6)
        if browse_command:
            self._button(parent, "浏览", browse_command, small=True).grid(
                row=row, column=2, sticky="w", padx=(8, 0)
            )
        if hint:
            tk.Label(
                parent, text=hint, bg=core.COLOR_PANEL, fg=core.COLOR_DIM_CYAN,
                font=(core.FONT_UI, 8), anchor="w", justify="left",
                wraplength=760,
            ).grid(row=row + 1, column=1, columnspan=2, sticky="w", pady=(0, 4))
            return row + 2
        return row + 1

    def _settings_switch(self, parent, row, title, hint, variable, color):
        frame = tk.Frame(parent, bg=core.COLOR_PANEL)
        frame.grid(row=row, column=0, columnspan=3, sticky="ew", pady=5)
        text = tk.Frame(frame, bg=core.COLOR_PANEL)
        text.pack(side="left", fill="x", expand=True)
        tk.Label(
            text, text=title, bg=core.COLOR_PANEL, fg=core.COLOR_TEXT,
            font=(core.FONT_UI, 9, "bold"),
        ).pack(anchor="w")
        tk.Label(
            text, text=hint, bg=core.COLOR_PANEL, fg=core.COLOR_DIM_CYAN,
            font=(core.FONT_UI, 8),
        ).pack(anchor="w", pady=(2, 0))
        switch = ToggleSwitch(frame, variable, color, bg=core.COLOR_PANEL)
        switch.pack(side="right", padx=(12, 0))
        return row + 1

    def _settings_info(self, parent, row, label, value):
        tk.Label(
            parent, text=label, bg=core.COLOR_PANEL, fg=core.COLOR_MUTED,
            font=(core.FONT_UI, 9),
        ).grid(row=row, column=0, sticky="nw", pady=5, padx=(0, 14))
        tk.Label(
            parent, text=value, bg=core.COLOR_PANEL, fg=core.COLOR_DIM_CYAN,
            font=(core.FONT_MONO, 8), anchor="w", justify="left",
            wraplength=760,
        ).grid(row=row, column=1, columnspan=2, sticky="w", pady=5)

    def _sync_settings_from_cfg(self):
        if not hasattr(self, "path_var"):
            return
        self.path_var.set(self.cfg["data_source"].get("path", ""))
        self.ops_path_var.set(self.cfg.get("operations", {}).get("path", ""))
        self.hover_var.set(bool(self.cfg["trigger"].get("hover_enabled", True)))
        self.hotkey_var.set(bool(self.cfg["trigger"].get("hotkey_enabled", True)))
        self.hotkey_value_var.set(self.cfg["trigger"].get("hotkey", "F8"))
        self.dwell_var.set(str(self.cfg["trigger"].get("hover_dwell_ms", 600)))
        self.font_var.set(str(self.cfg["display"].get("font_size", 13)))
        self.opacity_var.set(str(self.cfg["display"].get("opacity", 0.98)))
        self.auto_close_var.set(str(self.cfg["display"].get("auto_close_seconds", 15)))
        self.show_stock_var.set(bool(self.cfg["display"].get("show_stock", True)))
        self.show_module_name_var.set(
            bool(self.cfg["display"].get("show_module_name", False))
        )
        self.startup_var.set(startup_enabled())

    def browse_source(self):
        path = filedialog.askopenfilename(
            title="选择 SKU 数据源",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.path_var.set(path)

    def browse_ops(self):
        path = filedialog.askopenfilename(
            title="选择负责人关系表",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.ops_path_var.set(path)

    def save_settings(self):
        try:
            hotkey_text = self.hotkey_value_var.get().strip()
            if self.hotkey_var.get():
                core.parse_hotkey(hotkey_text)
            cfg = json.loads(json.dumps(self.cfg))
            cfg["data_source"]["path"] = self.path_var.get().strip()
            cfg["operations"]["path"] = self.ops_path_var.get().strip()
            cfg["trigger"]["hover_enabled"] = bool(self.hover_var.get())
            cfg["trigger"]["hotkey_enabled"] = bool(self.hotkey_var.get())
            cfg["trigger"]["hotkey"] = hotkey_text or "F8"
            cfg["trigger"]["hover_dwell_ms"] = max(100, int(float(self.dwell_var.get())))
            cfg["display"]["font_size"] = max(9, int(float(self.font_var.get())))
            cfg["display"]["opacity"] = max(0.3, min(1.0, float(self.opacity_var.get())))
            cfg["display"]["auto_close_seconds"] = max(0, int(float(self.auto_close_var.get())))
            cfg["display"]["show_stock"] = bool(self.show_stock_var.get())
            cfg["display"]["show_module_name"] = bool(
                self.show_module_name_var.get()
            )
            core.save_config(cfg)
            self.cfg = cfg
            self.popup.cfg = cfg
            self.store.update_config(cfg)
            self.ops_store.update_config(cfg)
            try:
                self.hotkey = core.parse_hotkey(cfg["trigger"]["hotkey"])
            except Exception:
                self.hotkey = None
            set_startup(bool(self.startup_var.get()))
            self.settings_status.configure(text="设置已保存", fg=core.COLOR_CYAN)
            self._on_data_status("设置已保存", "ok")
        except Exception as exc:
            messagebox.showerror("保存失败", str(exc), parent=self.root)

    def reload_data(self):
        self.store.reload(silent=False)
        self.ops_store.reload()
        self.refresh_matrix()

    # ---------- status/tray/hotkey ----------

    def _on_data_status(self, text, level="info"):
        colors = {
            "ok": core.COLOR_CYAN,
            "warn": core.COLOR_AMBER,
            "error": core.COLOR_MAGENTA,
            "info": core.COLOR_MUTED,
        }
        self._status_level = level

        def apply():
            self.status_var.set(("● " if not text.startswith("●") else "") + text)
            self.status_label.configure(fg=colors.get(level, core.COLOR_MUTED))

        try:
            self.post(apply)
        except Exception:
            pass

    def post(self, callback):
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
                logging.exception("界面回调失败")
        try:
            self.root.after(50, self._drain_callbacks)
        except Exception:
            pass

    def _start_tray(self):
        try:
            menu = pystray.Menu(
                pystray.MenuItem("打开查询页", self._tray_action(self.show_query), default=True),
                pystray.MenuItem("批量查询", self._tray_action(self.open_batch)),
                pystray.MenuItem("设置页面", self._tray_action(lambda: self.show_page("settings"))),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("重新加载数据", self._tray_action(self.reload_data)),
                pystray.MenuItem("退出", self._tray_action(self.quit)),
            )
            self.icon = pystray.Icon("STagSkuTag", core.make_tray_image(), "STag", menu)
            self.icon.run_detached()
        except Exception:
            logging.exception("托盘启动失败")
            self.icon = None

    def _tray_action(self, callback):
        def handler(_icon=None, _item=None):
            self.post(callback)
        return handler

    def _hotkey_tick(self):
        try:
            if self.cfg["trigger"].get("hotkey_enabled") and self.hotkey:
                mods, vk = self.hotkey
                down = (win32api.GetAsyncKeyState(vk) & 0x8000) != 0
                if down:
                    down = all((win32api.GetAsyncKeyState(mod) & 0x8000) != 0 for mod in mods)
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
        if not self.cfg["trigger"].get("hotkey_enabled"):
            return
        text, old = core.copy_selection()
        text = (text or "").strip()
        if old:
            core.set_clipboard_text(old)
        hit = core.lookup_text(text, self.store.mapping) if text else None
        if hit:
            self._show_hit(hit[0], hit[1])
            return
        self._hotkey_ocr_fallback(text)

    def _hotkey_ocr_fallback(self, copied_text):
        try:
            pos = win32api.GetCursorPos()
        except Exception:
            pos = (0, 0)
        cfg = self.cfg["trigger"]
        scale = max(3, int(cfg.get("ocr_scale", 3)))
        width = max(int(cfg.get("hover_region_width", 520) * self.dpi_scale), 800)
        height = max(int(cfg.get("hover_region_height", 48) * self.dpi_scale), 140)

        def work():
            hit = None
            try:
                image, left, top = core.capture_around(pos[0], pos[1], width, height)
                result = self.ocr.recognize(image, scale)
                hit = core.extract_best_hit(
                    result, left, top, scale, self.store.mapping, pos, 120 * self.dpi_scale
                )
            except core.OcrUnavailable as exc:
                logging.warning("F8 OCR 不可用：%s", exc)
            except Exception:
                logging.exception("F8 OCR 兜底失败")
            self.post(lambda: self._finish_hotkey_fallback(hit, copied_text))

        self.ocr_executor.submit(work)

    def _finish_hotkey_fallback(self, hit, copied_text):
        if hit:
            self._show_hit(hit["code"], hit["entry"])
            return
        preview = copied_text if len(copied_text) <= 80 else copied_text[:80] + "…"
        if preview:
            self.popup.show("未找到匹配", [f"选中内容：{preview}"], [("未找到", core.COLOR_MAGENTA)])
        else:
            self.popup.show("没读到内容", ["请重新选中编码，或把鼠标停在编码上再按 F8"], [("提示", core.COLOR_AMBER)])

    def _ensure_window_on_screen(self):
        """窗口被隐藏后，如果落在已断开的副屏上，拉回主屏中间。"""
        try:
            screen_w = self.root.winfo_screenwidth()
            screen_h = self.root.winfo_screenheight()
            width = self.root.winfo_width()
            height = self.root.winfo_height()
            if width <= 1 or height <= 1:
                width, height = 1280, 800
            x = self.root.winfo_x()
            y = self.root.winfo_y()
            if (
                x + width < 40
                or y + height < 40
                or x > screen_w - 40
                or y > screen_h - 40
            ):
                x = max(0, (screen_w - width) // 2)
                y = max(0, (screen_h - height) // 2 - 20)
                self.root.geometry(f"{width}x{height}+{x}+{y}")
        except Exception:
            pass

    def show_query(self):
        try:
            self.root.state("normal")
        except Exception:
            pass
        self.root.deiconify()
        self._ensure_window_on_screen()
        self.root.lift()
        self.show_page("query")
        try:
            self.root.focus_force()
        except Exception:
            pass

    def hide_window(self):
        self.root.withdraw()

    def _check_show_signal(self):
        try:
            if os.path.exists(STAG_SIGNAL_PATH):
                os.remove(STAG_SIGNAL_PATH)
                self.show_query()
        except Exception:
            pass
        try:
            self.root.after(600, self._check_show_signal)
        except Exception:
            pass

    def quit(self):
        if self._quitting:
            return
        self._quitting = True
        try:
            self.store.stop()
            self.ops_store.stop()
            self.module_store.close()
            self.ocr_executor.shutdown(wait=False, cancel_futures=True)
        except Exception:
            pass
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


class HudToplevel(tk.Toplevel):
    """无原生白边标题栏的暗色弹窗；自带标题栏、关闭和拖动。"""

    def __init__(self, master, title, width, height):
        super().__init__(master)
        self.overrideredirect(True)
        self.configure(
            bg=core.COLOR_BG,
            highlightthickness=1,
            highlightbackground=core.COLOR_LINE_BRIGHT,
        )
        self._drag_offset = None
        self.header = tk.Frame(self, bg=core.COLOR_PANEL, height=34)
        self.header.pack(fill="x")
        self.header.pack_propagate(False)
        tk.Label(
            self.header, text="STAG", bg=core.COLOR_PANEL,
            fg=core.COLOR_CYAN, font=(core.FONT_DISPLAY, 10, "bold"),
        ).pack(side="left", padx=(12, 8))
        self.title_label = tk.Label(
            self.header, text=title, bg=core.COLOR_PANEL,
            fg=core.COLOR_MUTED, font=(core.FONT_UI, 9),
        )
        self.title_label.pack(side="left")
        close = tk.Label(
            self.header, text="×", bg=core.COLOR_PANEL, fg=core.COLOR_MUTED,
            font=(core.FONT_DISPLAY, 13), width=3, cursor="hand2",
        )
        close.pack(side="right")
        close.bind("<Enter>", lambda _e: close.configure(fg=core.COLOR_MAGENTA))
        close.bind("<Leave>", lambda _e: close.configure(fg=core.COLOR_MUTED))
        close.bind("<Button-1>", lambda _e: self.destroy())
        for widget in (self.header, self.title_label):
            widget.bind("<ButtonPress-1>", self._start_drag)
            widget.bind("<B1-Motion>", self._drag)
            widget.bind("<ButtonRelease-1>", self._end_drag)
        self.update_idletasks()
        try:
            x = master.winfo_rootx() + max(0, (master.winfo_width() - width) // 2)
            y = master.winfo_rooty() + max(0, (master.winfo_height() - height) // 2)
        except Exception:
            x = y = 80
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _start_drag(self, event):
        self._drag_offset = (
            event.x_root - self.winfo_x(),
            event.y_root - self.winfo_y(),
        )

    def _drag(self, event):
        if not self._drag_offset:
            return
        self.geometry(
            f"+{event.x_root - self._drag_offset[0]}+{event.y_root - self._drag_offset[1]}"
        )

    def _end_drag(self, _event=None):
        self._drag_offset = None


class ModuleWindow:
    def __init__(self, app, module_id):
        self.app = app
        self.module_id = module_id
        self.sort_key = None
        self.sort_desc = False
        self._module_columns = []
        self._rows = []
        self._total = 0
        self._keys = []
        self._labels = {}
        self._column_names = {}
        self._base_widths = {}
        self._col_widths = {}
        self._header_offset = 0.0
        self._hover_index = None
        self._hover_t = 0.0
        self._hover_step = 0
        self._hover_after = None
        self._resize_after = None
        self.color_swatches = []
        self.win = HudToplevel(app.root, "模块详情", 980, 714)
        self._build()
        self.refresh()

    def _build(self):
        outer = tk.Frame(self.win, bg=core.COLOR_BG)
        outer.pack(fill="both", expand=True, padx=18, pady=14)

        header_card = tk.Frame(
            outer, bg=core.COLOR_PANEL,
            highlightthickness=1, highlightbackground=core.COLOR_LINE,
        )
        header_card.pack(fill="x", pady=(0, 10))
        header_inner = tk.Frame(header_card, bg=core.COLOR_PANEL)
        header_inner.pack(fill="x", padx=16, pady=(12, 10))

        tk.Label(
            header_inner, text="MODULE DETAIL", bg=core.COLOR_PANEL,
            fg=core.COLOR_CYAN, font=(core.FONT_MONO, 8),
        ).pack(anchor="w")
        self.title_label = tk.Label(
            header_inner, text="", bg=core.COLOR_PANEL, fg=core.COLOR_TEXT,
            font=(core.FONT_UI, 19, "bold"),
        )
        self.title_label.pack(anchor="w", pady=(4, 0))
        self.meta_label = tk.Label(
            header_inner, text="", bg=core.COLOR_PANEL, fg=core.COLOR_MUTED,
            font=(core.FONT_MONO, 8),
        )
        self.meta_label.pack(anchor="w", pady=(4, 0))

        actions = tk.Frame(header_inner, bg=core.COLOR_PANEL)
        actions.pack(fill="x", pady=(10, 0))
        for text, command in (
            ("导入更新", self.import_update),
            ("导出当前", self.export_current),
            ("版本历史", self.version_history),
        ):
            self.app._button(actions, text, command, small=True).pack(side="left", padx=(0, 7))
        self.toggle_button = self.app._button(actions, "停用", self.toggle, small=True)
        self.toggle_button.pack(side="left", padx=(0, 7))
        self.app._button(actions, "删除模块", self.delete, small=True).pack(side="left")

        self.columns_label = tk.Label(
            header_inner, text="", bg=core.COLOR_PANEL, fg=core.COLOR_DIM_CYAN,
            font=(core.FONT_UI, 9), anchor="w", justify="left", wraplength=880,
        )
        self.columns_label.pack(fill="x", pady=(8, 0))

        color_row = tk.Frame(header_inner, bg=core.COLOR_PANEL)
        color_row.pack(fill="x", pady=(8, 0))
        tk.Label(
            color_row, text="模块颜色", bg=core.COLOR_PANEL,
            fg=core.COLOR_MUTED, font=(core.FONT_UI, 9),
        ).pack(side="left", padx=(0, 8))
        for color in MODULE_COLORS:
            swatch = ColorSwatch(
                color_row, color, self.change_color, size=22,
                bg=core.COLOR_PANEL,
            )
            swatch.pack(side="left", padx=2)
            self.color_swatches.append(swatch)

        table_card = tk.Frame(
            outer, bg=core.COLOR_PANEL,
            highlightthickness=1, highlightbackground=core.COLOR_LINE,
        )
        table_card.pack(fill="both", expand=True)

        table_head = tk.Frame(table_card, bg=core.COLOR_PANEL)
        table_head.pack(fill="x", padx=14, pady=(10, 6))
        tk.Label(
            table_head, text="当前数据", bg=core.COLOR_PANEL, fg=core.COLOR_TEXT,
            font=(core.FONT_UI, 10, "bold"),
        ).pack(side="left")
        self.table_count_label = tk.Label(
            table_head, text="", bg=core.COLOR_PANEL, fg=core.COLOR_DIM_CYAN,
            font=(core.FONT_MONO, 8),
        )
        self.table_count_label.pack(side="right")

        table_frame = tk.Frame(table_card, bg=core.COLOR_PANEL)
        table_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.header_canvas = tk.Canvas(
            table_frame, bg=core.COLOR_PANEL, height=32,
            highlightthickness=0, bd=0, cursor="arrow",
        )
        self.tree = ttk.Treeview(
            table_frame, style="Stag.Flat.Treeview", show="",
        )
        self._xscroll = ttk.Scrollbar(
            table_frame, orient="horizontal", command=self.tree.xview,
            style="Stag.Horizontal.TScrollbar",
        )
        self._yscroll = ttk.Scrollbar(
            table_frame, orient="vertical", command=self.tree.yview,
            style="Stag.Vertical.TScrollbar",
        )
        self.tree.configure(
            yscrollcommand=self._on_yscroll, xscrollcommand=self._on_xscroll
        )
        self.header_canvas.grid(row=0, column=0, sticky="ew")
        self.tree.grid(row=1, column=0, sticky="nsew")
        self._yscroll.grid(row=1, column=1, sticky="ns")
        self._xscroll.grid(row=2, column=0, sticky="ew")
        table_frame.grid_rowconfigure(1, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        self.header_canvas.bind("<Configure>", lambda _e: self._draw_header())
        self.header_canvas.bind("<Enter>", self._on_header_motion)
        self.header_canvas.bind("<Motion>", self._on_header_motion)
        self.header_canvas.bind("<Leave>", self._on_header_leave)
        self.header_canvas.bind("<Button-1>", self._on_header_click)
        self.tree.bind("<Configure>", self._on_tree_configure)
        self.tree.tag_configure("odd", background="#0A111D")
        self.tree.tag_configure("even", background="#070B13")
        self.tree.tag_configure("disabled", foreground=core.COLOR_MUTED)

        self.footer = tk.Label(
            outer, text="", bg=core.COLOR_BG, fg=core.COLOR_DIM_CYAN,
            font=(core.FONT_MONO, 8), anchor="w",
        )
        self.footer.pack(fill="x", pady=(8, 0))

    def refresh(self):
        module = self.app.module_store.get_module(self.module_id)
        if not module or module["status"] == "deleted":
            self.win.destroy()
            self.app.refresh_matrix()
            return
        self.title_label.configure(text=module["name"])
        for swatch in self.color_swatches:
            swatch.set_selected(
                swatch.color.lower() == (module["color"] or "").lower()
            )
        self.meta_label.configure(
            text=f"{module['row_count']:,} 个 SKU · {module['col_count']} 列 · "
                 f"{module['status']} · {module['updated_at']}"
        )
        columns, rows, total = self.app.module_store.detail_rows(self.module_id, limit=500)
        self._module_columns = columns
        self._rows = rows
        self._total = total
        self.columns_label.configure(
            text="列：" + " / ".join(column["name"] for column in columns)
        )
        self.table_count_label.configure(text=f"{total:,} 条 · 显示 {len(rows):,}")
        self.toggle_button.configure(
            text="启用模块" if module["status"] == "disabled" else "停用模块"
        )
        self._setup_columns(columns)
        self._render_rows()
        self.win.after_idle(self._apply_column_widths)
        self.footer.configure(
            text=f"显示前 {len(rows):,} / {total:,} 个 SKU · 点击列头排序 · "
                 f"当前模块独立存储，查询时按编码 join"
        )

    # ---------- module detail table ----------

    def _setup_columns(self, columns):
        ids = [f"c{i}" for i in range(len(columns))]
        keys = ["code", "name"] + ids
        labels = {"code": "SKU编码", "name": "SKU名称"}
        column_names = {}
        widths = {"code": 110, "name": 280}
        for index, column in enumerate(columns):
            key = ids[index]
            labels[key] = column["name"]
            column_names[key] = column["name"]
            widths[key] = 110
        if self.sort_key not in keys:
            self.sort_key = None
            self.sort_desc = False
        self._keys = keys
        self._labels = labels
        self._column_names = column_names
        self._base_widths = widths
        self._col_widths = dict(widths)
        self.tree["columns"] = tuple(keys)
        for key in keys:
            self.tree.heading(key, text=labels[key], anchor="w")
            self.tree.column(
                key, width=widths[key], minwidth=60, anchor="w", stretch=False
            )

    def _render_rows(self):
        self.tree.delete(*self.tree.get_children())
        for index, row in enumerate(self._sorted_rows()):
            entry = self.app.store.mapping.get(row["code"])
            name = entry["name"] if entry else ""
            values = [
                row["values"].get(column["name"], "")
                for column in self._module_columns
            ]
            self.tree.insert(
                "", "end", values=[row["code"], name] + values,
                tags=("odd" if index % 2 else "even",),
            )
        self._draw_header()
        self.win.after_idle(self._sync_header_offset)

    def _sync_header_offset(self):
        if not self.win.winfo_exists():
            return
        self._on_xscroll(*self.tree.xview())

    def _row_sort_value(self, row):
        key = self.sort_key
        if key == "code":
            return str(row["code"] or "")
        if key == "name":
            entry = self.app.store.mapping.get(row["code"])
            return str(entry["name"] if entry else "")
        column_name = self._column_names.get(key, "")
        return str(row["values"].get(column_name, "") or "")

    def _sorted_rows(self):
        rows = list(self._rows)
        if not self.sort_key:
            return rows

        def sort_key(item):
            text = item[0].strip()
            try:
                return 0, float(text.replace(",", "").replace("%", ""))
            except ValueError:
                return 1, text.casefold()

        filled = []
        empty = []
        for row in rows:
            text = self._row_sort_value(row)
            if text.strip():
                filled.append((text, row))
            else:
                empty.append((text, row))
        filled.sort(key=sort_key, reverse=self.sort_desc)
        return [row for _text, row in filled] + [row for _text, row in empty]

    def _toggle_scrollbar(self, scrollbar, visible):
        try:
            mapped = bool(scrollbar.winfo_ismapped())
        except Exception:
            return
        if visible and not mapped:
            scrollbar.grid()
        elif not visible and mapped:
            scrollbar.grid_remove()

    @staticmethod
    def _scrollbar_needed(first, last):
        return not (float(first) <= 0.0001 and float(last) >= 0.9999)

    def _on_yscroll(self, first, last):
        self._yscroll.set(first, last)
        self._toggle_scrollbar(
            self._yscroll, self._scrollbar_needed(first, last)
        )

    def _on_xscroll(self, first, last):
        self._xscroll.set(first, last)
        self._toggle_scrollbar(
            self._xscroll, self._scrollbar_needed(first, last)
        )
        total = sum(self._actual_column_widths().values())
        self._header_offset = float(first) * total
        self._draw_header()

    def _actual_column_widths(self):
        widths = {}
        for key in self._keys:
            try:
                widths[key] = max(1, int(self.tree.column(key, "width")))
            except Exception:
                widths[key] = max(1, int(self._col_widths.get(key, 0)))
        return widths

    def _on_tree_configure(self, _event=None):
        if self._resize_after:
            try:
                self.win.after_cancel(self._resize_after)
            except Exception:
                pass
        self._resize_after = self.win.after_idle(self._apply_column_widths)

    def _apply_column_widths(self):
        self._resize_after = None
        if not self.win.winfo_exists():
            return
        if not self._keys or not self._base_widths:
            return
        available = int(self.tree.winfo_width())
        if available <= 10:
            return
        base_total = sum(self._base_widths.values())
        widths = dict(self._base_widths)
        if base_total and available > base_total:
            factor = available / base_total
            widths = {
                key: max(60, int(round(width * factor)))
                for key, width in self._base_widths.items()
            }
            widths[self._keys[-1]] += available - sum(widths.values())
        for key, width in widths.items():
            try:
                current = int(self.tree.column(key, "width"))
            except Exception:
                current = None
            if current != width:
                self.tree.column(key, width=width, stretch=False)
        self._col_widths = self._actual_column_widths()
        self._draw_header()

    def _header_index_at(self, x):
        position = x + self._header_offset
        edge = 0.0
        for index, key in enumerate(self._keys):
            edge += self._col_widths.get(key, 0)
            if position < edge:
                return index
        return None

    def _on_header_motion(self, event):
        index = self._header_index_at(event.x)
        self.header_canvas.configure(
            cursor="hand2" if index is not None else "arrow"
        )
        if index == self._hover_index:
            return
        if self._hover_after:
            try:
                self.win.after_cancel(self._hover_after)
            except Exception:
                pass
            self._hover_after = None
        self._hover_index = index
        self._hover_step = 0
        self._hover_t = 0.0
        if index is not None:
            self._animate_header_hover()
        else:
            self._draw_header()

    def _on_header_leave(self, _event=None):
        if self._hover_after:
            try:
                self.win.after_cancel(self._hover_after)
            except Exception:
                pass
            self._hover_after = None
        self._hover_index = None
        self._hover_t = 0.0
        self.header_canvas.configure(cursor="arrow")
        self._draw_header()

    def _animate_header_hover(self):
        if self._hover_index is None or not self.win.winfo_exists():
            return
        self._hover_step = min(5, getattr(self, "_hover_step", 0) + 1)
        self._hover_t = self._hover_step / 5.0
        self._draw_header()
        if self._hover_step < 5:
            self._hover_after = self.win.after(16, self._animate_header_hover)
        else:
            self._hover_after = None

    def _on_header_click(self, event):
        index = self._header_index_at(event.x)
        if index is None:
            return
        key = self._keys[index]
        if self.sort_key == key:
            self.sort_desc = not self.sort_desc
        else:
            self.sort_key = key
            self.sort_desc = False
        self.tree.yview_moveto(0)
        self._render_rows()

    def _rounded_rect(self, canvas, x0, y0, x1, y1, radius, **kwargs):
        radius = max(0, min(radius, (x1 - x0) / 2, (y1 - y0) / 2))
        points = (
            x0 + radius, y0, x1 - radius, y0, x1, y0, x1, y0 + radius,
            x1, y1 - radius, x1, y1, x1 - radius, y1, x0 + radius, y1,
            x0, y1, x0, y1 - radius, x0, y0 + radius, x0, y0,
        )
        return canvas.create_polygon(
            points, smooth=True, splinesteps=18, **kwargs
        )

    def _draw_header(self):
        canvas = self.header_canvas
        canvas.delete("all")
        width = max(1, int(canvas.winfo_width()))
        height = max(1, int(canvas.winfo_height()))
        if not self._keys:
            canvas.create_line(
                0, height - 1, width, height - 1, fill=core.COLOR_LINE
            )
            return
        self._col_widths = self._actual_column_widths()
        x = -self._header_offset
        for index, key in enumerate(self._keys):
            column_width = self._col_widths.get(key, 0)
            if column_width <= 0:
                continue
            x0, x1 = x, x + column_width
            if x1 > 0 and x0 < width:
                hovered = index == self._hover_index
                if hovered:
                    grow = 1.0 + 2.0 * self._hover_t
                    self._rounded_rect(
                        canvas,
                        x0 + 3 - grow, 3 - grow,
                        x1 - 3 + grow, height - 3 + grow,
                        5, fill="", outline=core.COLOR_CYAN, width=1,
                    )
                elif x1 < width:
                    canvas.create_line(
                        x1 - 0.5, 8, x1 - 0.5, height - 8,
                        fill=core.COLOR_LINE,
                    )
                canvas.create_text(
                    x0 + 10, height / 2, text=self._labels.get(key, ""),
                    anchor="w", fill=core.COLOR_CYAN,
                    font=(core.FONT_UI, 9, "bold"),
                )
                if key == self.sort_key:
                    cy = height / 2
                    if self.sort_desc:
                        points = (
                            x1 - 15, cy - 3, x1 - 7, cy - 3, x1 - 11, cy + 4,
                        )
                    else:
                        points = (
                            x1 - 15, cy + 3, x1 - 7, cy + 3, x1 - 11, cy - 4,
                        )
                    canvas.create_polygon(
                        points, fill=core.COLOR_CYAN, outline=""
                    )
            x = x1
        canvas.create_line(
            0, height - 1, width, height - 1, fill=core.COLOR_LINE
        )

    def change_color(self, color):
        try:
            self.app.module_store.set_module_color(self.module_id, color)
            for swatch in self.color_swatches:
                swatch.set_selected(swatch.color.lower() == color.lower())
            self.app.refresh_matrix()
        except Exception as exc:
            messagebox.showerror("颜色修改失败", str(exc), parent=self.win)

    def import_update(self):
        path = filedialog.askopenfilename(
            parent=self.win, title="选择模块 Excel",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not path:
            return
        module = self.app.module_store.get_module(self.module_id)
        try:
            self.app.module_store.import_excel(
                path, module["name"], module["description"], module["color"],
                module_id=self.module_id,
            )
            self.refresh()
            self.app.refresh_matrix()
        except Exception as exc:
            logging.exception("导入模块失败")
            messagebox.showerror("导入失败", str(exc), parent=self.win)

    def export_current(self):
        module = self.app.module_store.get_module(self.module_id)
        path = filedialog.asksaveasfilename(
            parent=self.win, title="导出模块数据",
            defaultextension=".xlsx", initialfile=f"{module['name']}.xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not path:
            return
        try:
            count = self.app.module_store.export_module(
                self.module_id, path, self.app.store.mapping
            )
            messagebox.showinfo("导出完成", f"已导出 {count:,} 个 SKU", parent=self.win)
        except Exception as exc:
            logging.exception("导出模块失败")
            messagebox.showerror("导出失败", str(exc), parent=self.win)

    def version_history(self):
        VersionWindow(self.app, self.module_id, self.refresh)

    def toggle(self):
        try:
            self.app.module_store.toggle_module(self.module_id)
            self.refresh()
            self.app.refresh_matrix()
        except Exception as exc:
            messagebox.showerror("操作失败", str(exc), parent=self.win)

    def delete(self):
        self.app._delete_module(self.module_id)
        module = self.app.module_store.get_module(self.module_id)
        if not module or module["status"] == "deleted":
            self.win.destroy()


class VersionWindow:
    def __init__(self, app, module_id, refresh_callback=None):
        self.app = app
        self.module_id = module_id
        self.refresh_callback = refresh_callback
        self.win = HudToplevel(app.root, "版本历史", 640, 494)
        self.listbox = tk.Listbox(
            self.win, bg=core.COLOR_PANEL, fg=core.COLOR_TEXT, bd=0,
            highlightthickness=1, highlightbackground=core.COLOR_LINE,
            selectbackground=core.COLOR_DIM_CYAN, font=(core.FONT_MONO, 9),
        )
        self.listbox.pack(fill="both", expand=True, padx=18, pady=(18, 8))
        buttons = tk.Frame(self.win, bg=core.COLOR_BG)
        buttons.pack(fill="x", padx=18, pady=(0, 16))
        app._button(buttons, "撤回最近一次导入", self.undo, accent=True).pack(side="left")
        app._button(buttons, "关闭", self.win.destroy).pack(side="left", padx=(8, 0))
        self.refresh()

    def refresh(self):
        self.listbox.delete(0, "end")
        for version in self.app.module_store.versions(self.module_id):
            self.listbox.insert(
                "end",
                f"v{version['version_no']}  {version['created_at']}  "
                f"{version['row_count']:,} 行  {version['status']}",
            )

    def undo(self):
        if not messagebox.askyesno(
            "撤回导入",
            "撤回最近一次导入，并恢复到上一个版本？",
            parent=self.win,
        ):
            return
        try:
            previous = self.app.module_store.undo_latest(self.module_id)
        except Exception as exc:
            messagebox.showerror("撤回失败", str(exc), parent=self.win)
            return
        self.refresh()
        if self.refresh_callback:
            self.refresh_callback()
        self.app.refresh_matrix()
        messagebox.showinfo(
            "撤回完成",
            "已恢复到上一个版本" if previous else "已撤回首次导入，模块已移除",
            parent=self.win,
        )


class ColorSwatch(tk.Canvas):
    def __init__(self, parent, color, command, size=46, bg=None):
        super().__init__(
            parent, width=size, height=size, bg=bg or core.COLOR_BG,
            highlightthickness=0, bd=0, cursor="hand2",
        )
        self.color = color
        self.command = command
        self.selected = False
        self.size = int(size)
        self.bind("<Button-1>", lambda _event: self.command(self.color))
        self.redraw()

    def set_selected(self, selected):
        self.selected = bool(selected)
        self.redraw()

    def redraw(self):
        self.delete("all")
        size = self.size
        if self.selected:
            self.create_rectangle(
                1, 1, size - 1, size - 1, outline=self.color, width=1
            )
            inset = max(3, size // 8)
            self.create_rectangle(
                inset, inset, size - inset, size - inset,
                fill=hex_blend(self.color, core.COLOR_BG, 0.22),
                outline=self.color, width=2,
            )
            center = max(2, size // 6)
            self.create_rectangle(
                center, center, size - center, size - center,
                fill=self.color, outline="",
            )
        else:
            inset = max(3, size // 5)
            self.create_rectangle(
                inset, inset, size - inset, size - inset,
                fill=self.color, outline="",
            )


class NewModuleWindow:
    def __init__(self, app, path):
        self.app = app
        self.path = path
        self.win = HudToplevel(app.root, "新建模块", 680, 504)
        self.name_var = tk.StringVar()
        self.desc_var = tk.StringVar()
        self.file_var = tk.StringVar(value=path)
        self.color = MODULE_COLORS[0]
        self.swatches = []

        tk.Label(
            self.win, text="NEW MODULE / 新建模块", bg=core.COLOR_BG,
            fg=core.COLOR_CYAN, font=(core.FONT_MONO, 8),
        ).pack(anchor="w", padx=22, pady=(18, 4))
        tk.Label(
            self.win, text="Excel 已经选好，现在给这个模块命名。",
            bg=core.COLOR_BG, fg=core.COLOR_TEXT, font=(core.FONT_UI, 16),
        ).pack(anchor="w", padx=22)
        tk.Label(
            self.win, text="模块名只用于 STag 内部，不会写回 Excel。",
            bg=core.COLOR_BG, fg=core.COLOR_MUTED, font=(core.FONT_UI, 9),
        ).pack(anchor="w", padx=22, pady=(5, 0))

        file_card = tk.Frame(
            self.win, bg=core.COLOR_PANEL,
            highlightthickness=1, highlightbackground=core.COLOR_LINE,
        )
        file_card.pack(fill="x", padx=22, pady=(16, 10))
        file_inner = tk.Frame(file_card, bg=core.COLOR_PANEL)
        file_inner.pack(fill="x", padx=14, pady=11)
        tk.Label(
            file_inner, text="EXCEL SOURCE", bg=core.COLOR_PANEL,
            fg=core.COLOR_DIM_CYAN, font=(core.FONT_MONO, 8),
        ).pack(anchor="w")
        tk.Label(
            file_inner, text=os.path.basename(path), bg=core.COLOR_PANEL,
            fg=core.COLOR_TEXT, font=(core.FONT_UI, 10, "bold"),
        ).pack(anchor="w", pady=(3, 0))
        tk.Label(
            file_inner, text=path, bg=core.COLOR_PANEL, fg=core.COLOR_MUTED,
            font=(core.FONT_MONO, 8), anchor="w", justify="left",
            wraplength=560,
        ).pack(anchor="w", pady=(2, 0))
        self.app._button(file_inner, "重新选择", self.choose_file, small=True).pack(
            anchor="e", pady=(6, 0)
        )

        form = tk.Frame(self.win, bg=core.COLOR_BG)
        form.pack(fill="x", padx=22, pady=(4, 0))
        tk.Label(
            form, text="模块名称", bg=core.COLOR_BG, fg=core.COLOR_MUTED,
            font=(core.FONT_UI, 9),
        ).grid(row=0, column=0, sticky="w", pady=6)
        self.name_entry = self.app._entry(form, self.name_var, width=40)
        self.name_entry.grid(row=0, column=1, sticky="ew", pady=6, padx=(12, 0))
        tk.Label(
            form, text="说明", bg=core.COLOR_BG, fg=core.COLOR_MUTED,
            font=(core.FONT_UI, 9),
        ).grid(row=1, column=0, sticky="w", pady=6)
        self.app._entry(form, self.desc_var, width=40).grid(
            row=1, column=1, sticky="ew", pady=6, padx=(12, 0)
        )
        form.grid_columnconfigure(1, weight=1)

        color_row = tk.Frame(self.win, bg=core.COLOR_BG)
        color_row.pack(fill="x", padx=22, pady=(8, 0))
        tk.Label(
            color_row, text="模块颜色", bg=core.COLOR_BG, fg=core.COLOR_MUTED,
            font=(core.FONT_UI, 9),
        ).pack(side="left", padx=(0, 8))
        for color in MODULE_COLORS:
            swatch = ColorSwatch(color_row, color, self.set_color)
            swatch.pack(side="left", padx=1)
            self.swatches.append(swatch)
        self.swatches[0].set_selected(True)

        buttons = tk.Frame(self.win, bg=core.COLOR_BG)
        buttons.pack(fill="x", padx=22, pady=20)
        self.app._button(buttons, "导入为模块", self.do_import, accent=True).pack(side="left")
        self.app._button(buttons, "取消", self.win.destroy).pack(side="left", padx=(8, 0))
        self.status = tk.Label(
            self.win, text="", bg=core.COLOR_BG, fg=core.COLOR_MUTED,
            font=(core.FONT_UI, 9),
        )
        self.status.pack(anchor="w", padx=22)
        self.name_entry.focus_set()

    def set_color(self, color):
        self.color = color
        for swatch in self.swatches:
            swatch.set_selected(swatch.color == color)

    def choose_file(self):
        path = filedialog.askopenfilename(
            parent=self.win, title="选择模块 Excel",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.path = path
            self.file_var.set(path)
            self.win.destroy()
            NewModuleWindow(self.app, path)

    def do_import(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning("缺少名称", "给这个模块起一个名字", parent=self.win)
            self.name_entry.focus_set()
            return
        try:
            self.status.configure(text="正在导入…", fg=core.COLOR_CYAN)
            self.win.update_idletasks()
            module_id = self.app.module_store.import_excel(
                self.path, name, self.desc_var.get().strip(), self.color
            )
            self.app.refresh_matrix()
            self.win.destroy()
            ModuleWindow(self.app, module_id)
        except Exception as exc:
            logging.exception("新建模块失败")
            self.status.configure(text=str(exc), fg=core.COLOR_MAGENTA)
            messagebox.showerror("导入失败", str(exc), parent=self.win)


class BatchWindow:
    def __init__(self, app):
        self.app = app
        self.win = HudToplevel(app.root, "批量查码", 1180, 754)

        tk.Label(
            self.win, text="BATCH QUERY / 批量查码", bg=core.COLOR_BG,
            fg=core.COLOR_CYAN, font=(core.FONT_MONO, 8),
        ).pack(anchor="w", padx=20, pady=(16, 4))
        tk.Label(
            self.win, text="每行一个 SKU 编码，不要逗号；自动去重，按输入顺序返回。",
            bg=core.COLOR_BG, fg=core.COLOR_MUTED, font=(core.FONT_UI, 9),
        ).pack(anchor="w", padx=20)
        self.input = tk.Text(
            self.win, height=7, bg=core.COLOR_INPUT, fg=core.COLOR_TEXT,
            insertbackground=core.COLOR_CYAN, bd=0, highlightthickness=1,
            highlightbackground=core.COLOR_LINE_BRIGHT, font=(core.FONT_MONO, 10),
            padx=10, pady=8,
        )
        self.input.pack(fill="x", padx=20, pady=(12, 8))
        self.input.insert("1.0", "101072403\n101072404\n101072401")

        actions = tk.Frame(self.win, bg=core.COLOR_BG)
        actions.pack(fill="x", padx=20)
        app._button(actions, "开始查询", self.run, accent=True).pack(side="left")
        app._button(actions, "清空", self.clear).pack(side="left", padx=(8, 0))
        self.summary = tk.Label(
            actions, text="等待查询…", bg=core.COLOR_BG, fg=core.COLOR_MUTED,
            font=(core.FONT_UI, 9),
        )
        self.summary.pack(side="left", padx=(14, 0))

        table_frame = tk.Frame(self.win, bg=core.COLOR_BG)
        table_frame.pack(fill="both", expand=True, padx=20, pady=(10, 18))
        self.tree = ttk.Treeview(table_frame, style="Stag.Treeview", show="headings")
        yscroll = ttk.Scrollbar(
            table_frame, orient="vertical", command=self.tree.yview,
            style="Stag.Vertical.TScrollbar",
        )
        xscroll = ttk.Scrollbar(
            table_frame, orient="horizontal", command=self.tree.xview,
            style="Stag.Horizontal.TScrollbar",
        )
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.run()

    def parse_codes(self):
        raw = self.input.get("1.0", "end")
        codes = []
        seen = set()
        for part in re.split(r"\s+", raw):
            code = part.strip()
            if not code or code in seen:
                continue
            seen.add(code)
            codes.append(code)
            if len(codes) >= 200:
                break
        return codes

    def clear(self):
        self.input.delete("1.0", "end")
        self.tree.delete(*self.tree.get_children())
        self.summary.configure(text="等待查询…")

    def run(self):
        codes = self.parse_codes()
        if not codes:
            self.summary.configure(text="请输入至少一个编码")
            return
        modules = [
            module for module in self.app.module_store.list_modules()
            if module["status"] == "active"
        ]
        module_columns = []
        for module in modules:
            for column in self.app.module_store.columns(module["id"]):
                module_columns.append((module, column))

        base = ("code", "name", "style", "stock", "owner", "assistant")
        keys = list(base) + [f"m{module['id']}_{column['id']}" for module, column in module_columns]
        self.tree["columns"] = tuple(keys)
        for key, label in (
            ("code", "编码"), ("name", "名称"), ("style", "款式"),
            ("stock", "库存"), ("owner", "负责人"), ("assistant", "助理"),
        ):
            self.tree.heading(key, text=label, anchor="w")
            self.tree.column(key, width=110 if key != "name" else 250, anchor="w")
        for index, (module, column) in enumerate(module_columns):
            key = keys[len(base) + index]
            self.tree.heading(key, text=f"{module['name']} · {column['name']}", anchor="w")
            self.tree.column(key, width=110, anchor="w")

        module_values = self.app.module_store.lookup_many(codes)
        self.tree.delete(*self.tree.get_children())
        found = 0
        for code in codes:
            entry = self.app.store.mapping.get(code)
            if entry:
                found += 1
                name = entry["name"]
                style = "、".join(tag for tag, _ in core.detect_styles(name))
                stock = entry.get("stock") or {}
                stock_text = " / ".join(f"{key} {value}" for key, value in stock.items())
                ops = self.app.ops_store.lookup(code) or {}
                owner = " / ".join(ops.get("owner") or [])
                assistant = " / ".join(ops.get("assistant") or [])
                values = [code, name, style, stock_text, owner, assistant]
            else:
                values = [code, "未找到", "—", "—", "—", "—"]
            by_module = {
                item["module_id"]: item for item in module_values.get(code, [])
            }
            for module, column in module_columns:
                item = by_module.get(module["id"])
                value = ""
                if item:
                    for module_column in item["columns"]:
                        if module_column["name"] == column["name"]:
                            value = module_column["value"]
                            break
                values.append(value)
            self.tree.insert("", "end", values=values)
        self.summary.configure(
            text=f"共 {len(codes)} 个编码 · 找到 {found} 个 · 未找到 {len(codes) - found} 个 · 可左右滑动查看模块列"
        )


# ---------------------------------------------------------------- startup

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
            winreg.QueryValueEx(key, RUN_NAME)
            return True
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


def _show_existing_window():
    """把已经在托盘/后台的 STag 主窗口重新拉到前台。"""
    user32 = ctypes.windll.user32
    user32.FindWindowW.restype = ctypes.c_void_p
    user32.FindWindowW.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p]
    for _ in range(50):
        hwnd = user32.FindWindowW(None, APP_TITLE)
        if hwnd:
            user32.ShowWindow(hwnd, 9)  # SW_RESTORE
            user32.ShowWindow(hwnd, 5)  # SW_SHOW
            user32.BringWindowToTop(hwnd)
            user32.SetForegroundWindow(hwnd)
            return True
        time.sleep(0.1)
    return False


def _message_box(text, title, flags=0x40):
    try:
        user32 = ctypes.windll.user32
        user32.MessageBoxW.argtypes = [
            ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_uint
        ]
        user32.MessageBoxW(None, text, title, flags)
    except Exception:
        pass


def single_instance_guard():
    kernel32 = ctypes.windll.kernel32
    handle = kernel32.CreateMutexW(None, False, "STAG_APP")
    if kernel32.GetLastError() == 183:
        try:
            os.makedirs(core.CACHE_DIR, exist_ok=True)
            with open(STAG_SIGNAL_PATH, "w", encoding="utf-8") as handle_file:
                handle_file.write("show")
        except Exception:
            pass
        if not _show_existing_window():
            _message_box(
                "STag 已经在后台运行，但没有找到主窗口。\n\n"
                "请打开任务管理器，结束 STag.exe 后重新双击。\n"
                "如果托盘里有 STag 图标，也可以右键 → 打开查询页。",
                "STag",
            )
        return None
    return handle


def main():
    setup_logging()
    logging.info("STag 启动")
    guard = single_instance_guard()
    if guard is None:
        logging.info("已有 STag 实例在运行，退出本次启动")
        return 0
    try:
        app = StagApp()
        if "--tray" in sys.argv:
            app.root.withdraw()
        else:
            app.show_query()
        if "--settings" in sys.argv:
            app.show_page("settings")
        app.run()
    except Exception as exc:
        logging.exception("STag 异常退出")
        _message_box(
            f"STag 启动失败：{exc}\n\n"
            f"请把日志发给发版的人：\n{LOG_PATH}\n"
            f"{os.path.join(core.CACHE_DIR, 'stag.log')}",
            "STag 启动失败",
            0x10,
        )
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
