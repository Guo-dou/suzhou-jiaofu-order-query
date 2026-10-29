# -*- coding: utf-8 -*-
"""
苏州交复文件管理查询
====================
苹果风格界面的订单文件夹快速查询工具。
输入客户订单号（3DP格式）、客户名称或患者姓名，快速定位并打开对应的订单文件夹。

功能特点:
  - 数据路径全部可自选: 数据根目录 + 工作跟进表文件, 保存到同目录 config.json
  - 首次运行或路径无效时, 自动弹出路径选择窗口
  - 输入订单号(完整/部分/序号)、客户名或患者名, 实时匹配, 回车/双击一键打开文件夹
  - 启动异常自动写入 error.log, 便于排查
  - by 阿果
"""
import os
import re
import sys
import json
import ctypes
import time
import traceback
import threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter import font as tkfont
from collections import Counter

# ============ 默认路径（可被 config.json 覆盖；仅为示例占位，实际以用户选择为准） ============
DEFAULT_ROOT = r"D:\公司数据\苏州交复"
DEFAULT_XLSX = os.path.join(DEFAULT_ROOT, "苏州交复工作跟进表.xlsx")

APP_DIR = os.path.dirname(os.path.abspath(sys.argv[0] if sys.argv and sys.argv[0] else __file__))
CONFIG_PATH = os.path.join(APP_DIR, "config.json")
ERROR_LOG = os.path.join(APP_DIR, "error.log")

PRODUCT_DIRS = ["01 侧弯订单", "02 鞋垫订单", "03 头枕订单", "04 服务订单", "05 医学模型"]

# 订单编号前缀 -> (产品类型, 根目录下的相对子路径)
PREFIX_MAP = {
    "3DP-101": ("侧弯支具", r"01 侧弯订单\交复侧弯订单管理"),
    "3DP-104": ("固定支具", r"01 侧弯订单\固定支具订单管理"),
    "3DP-102": ("矫形鞋垫", r"02 鞋垫订单\订单管理"),
    "3DP-103": ("头枕",     r"03 头枕订单"),
    "3DP-105": ("医学模型", r"05 医学模型"),
}
# 售后件可能多一层: 01侧弯/交复侧弯订单管理/0售后件/客户/订单
AFTERSALES_DIRS = {
    "3DP-101": r"01 侧弯订单\交复侧弯订单管理\0售后件",
    "3DP-104": r"01 侧弯订单\固定支具订单管理\0售后件",
    "3DP-102": r"02 鞋垫订单\订单管理\0售后件",
}

PAT3 = re.compile(r"3DP-\d{3}-\d{8}-\d{3}")

# Excel 订单跟踪表列索引
C_ORDER, C_TIME, C_CUST, C_TYPE, C_PROD, C_DATA = 0, 1, 2, 3, 4, 5
C_STATUS, C_MACHINE, C_MATERIAL, C_EXPRESS = 36, 17, 18, 35

# ============ 苹果风格配色 ============
BG        = "#F5F5F7"   # 主背景: 浅灰
CARD      = "#FFFFFF"   # 卡片: 白
INPUT_BG  = "#F2F2F7"   # 输入框底
ACCENT    = "#0071E3"   # 苹果蓝
ACCENT_HV = "#0077ED"   # 悬停亮
ACCENT_PR = "#005FC5"   # 按下暗
TEXT      = "#1D1D1F"
SUB       = "#6E6E73"
LINE      = "#E8E8ED"
SCROLL    = "#D1D1D6"


def sval(v):
    """单元格值安全转字符串"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def resource_path(rel):
    """打包为 exe 后也能找到内嵌资源"""
    try:
        base = sys._MEIPASS
    except Exception:
        base = APP_DIR
    return os.path.join(base, rel)


def enable_dpi_awareness():
    """Windows 高分屏 DPI 感知: 必须在创建 Tk 前调用, 避免字体/界面被系统放大而模糊"""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # 系统级 DPI 感知
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


# ============ 配置 ============
def _write_error_log(msg):
    """把异常信息追加写入 error.log（程序目录）, 便于排查"""
    try:
        with open(ERROR_LOG, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]\n{msg}\n")
    except Exception:
        pass


def _global_excepthook(tp, val, tb):
    """任何未捕获异常都落盘, noconsole 模式下也能排查"""
    _write_error_log(f"未捕获异常 {tp.__name__}: {val}\n{''.join(traceback.format_tb(tb))}")


sys.excepthook = _global_excepthook


def load_config():
    """读取配置, 返回 (root, xlsx, 完整配置字典); 无配置或路径无效返回 (None, None, {})"""
    try:
        if os.path.isfile(CONFIG_PATH):
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            root = cfg.get("root", "")
            xlsx = cfg.get("xlsx", "")
            if root and os.path.isdir(root) and xlsx and os.path.isfile(xlsx):
                return root, xlsx, cfg
    except Exception:
        pass
    return None, None, {}


def save_config(root, xlsx, col_widths=None):
    try:
        data = {"root": root, "xlsx": xlsx}
        if col_widths:
            data["col_widths"] = {str(k): int(v) for k, v in col_widths.items()}
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


# ============ 数据层 ============
def load_excel_orders(xlsx):
    """读取订单跟踪表, 返回 (信息dict, 按时间倒序列表); 表结构不符时给出明确中文错误"""
    import openpyxl
    info = {}
    try:
        wb = openpyxl.load_workbook(xlsx, data_only=True, read_only=True)
    except Exception as e:
        raise RuntimeError(f"无法打开跟进表文件：{e}") from e
    try:
        if "订单跟踪" not in wb.sheetnames:
            raise RuntimeError(
                f"跟进表中未找到“订单跟踪”工作表。现有工作表：{'、'.join(wb.sheetnames[:8]) or '（无）'}")
        ws = wb["订单跟踪"]
        if ws.max_column < C_EXPRESS + 1:
            raise RuntimeError(
                f"跟进表列数不足（应至少 {C_EXPRESS + 1} 列，实际最大 {ws.max_column} 列），"
                f"请使用苏州交复标准《订单跟踪》表结构。")
        for r in ws.iter_rows(min_row=2, values_only=True):
            oid = r[C_ORDER]
            if not oid:
                continue
            oid = str(oid).strip()
            if not PAT3.match(oid):
                continue
            info[oid] = {
                "oid": oid,
                "time": r[C_TIME],
                "cust": sval(r[C_CUST]),
                "otype": sval(r[C_TYPE]),
                "prod": sval(r[C_PROD]),
                "data": sval(r[C_DATA]),
                "status": sval(r[C_STATUS]),
                "machine": sval(r[C_MACHINE]),
                "material": sval(r[C_MATERIAL]),
                "express": sval(r[C_EXPRESS]),
            }
    finally:
        wb.close()
    ordered = sorted(info.values(), key=lambda x: str(x["time"] or ""), reverse=True)
    return info, ordered


def scan_disk_orders(root):
    """全量递归扫描产品目录, 返回 {3DP订单号: 完整路径}"""
    idx = {}
    for d in PRODUCT_DIRS:
        base = os.path.join(root, d)
        if not os.path.isdir(base):
            continue
        for root2, dirs, _files in os.walk(base):
            for name in dirs:
                if PAT3.match(name):
                    idx[name] = os.path.join(root2, name)
    return idx


def expected_path(oid, excel_info, root):
    """根据订单号前缀 + Excel客户名, 构造预期文件夹路径; 不存在返回 None"""
    prefix = oid[:7]
    map_item = PREFIX_MAP.get(prefix)
    if not map_item:
        return None
    _prod, sub = map_item
    cust = excel_info.get(oid, {}).get("cust")
    if not cust:
        return None
    p = os.path.join(root, sub, cust, oid)
    if os.path.isdir(p):
        return p
    aft = AFTERSALES_DIRS.get(prefix)
    if aft:
        p2 = os.path.join(root, aft, cust, oid)
        if os.path.isdir(p2):
            return p2
    return None


# ============ 苹果风格组件 ============
class RoundButton(tk.Canvas):
    """圆角按钮: 蓝底白字, 悬停/按下变色"""
    def __init__(self, master, text, command, bg=ACCENT, fg="white",
                 radius=15, font=None, padx=20, pady=8, bold=True, **kw):
        self._bg, self._fg = bg, fg
        self._radius, self._command = radius, command
        self._padx, self._pady = padx, pady
        self._font = font or (("Microsoft YaHei", 11, "bold") if bold else ("Microsoft YaHei", 10))
        parent_bg = master.cget("bg") if isinstance(master, tk.Widget) else BG
        super().__init__(master, highlightthickness=0, bd=0, bg=parent_bg, **kw)
        self._font_obj = tkfont.Font(root=self, family=self._font[0], size=self._font[1], weight=self._font[2] if len(self._font) > 2 else "normal")
        self._tw = self._font_obj.measure(text)
        self._th = self._font_obj.metrics("linespace")
        self.configure(width=self._tw + self._padx * 2, height=self._th + self._pady * 2)
        self._text = text
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)
        self._draw(self._bg)

    def _draw(self, color):
        self.delete("all")
        w, h = int(self["width"]), int(self["height"])
        self.create_oval(0, 0, h, h, fill=color, outline="")
        self.create_oval(w - h, 0, w, h, fill=color, outline="")
        self.create_rectangle(h / 2, 0, w - h / 2, h, fill=color, outline="")
        self.create_text(w / 2, h / 2, text=self._text, fill=self._fg,
                         font=self._font_obj)

    def _on_enter(self, _e):
        self._draw(ACCENT_HV if self._bg == ACCENT else self._bg)

    def _on_leave(self, _e):
        self._draw(self._bg)

    def _on_press(self, _e):
        self._draw(ACCENT_PR if self._bg == ACCENT else self._bg)

    def _on_release(self, _e):
        self._draw(ACCENT_HV if self._bg == ACCENT else self._bg)
        if self._command:
            self._command()


class GhostButton(RoundButton):
    """浅灰圆角按钮: 次要用"""
    def __init__(self, master, text, command, padx=16, pady=6, bold=False, **kw):
        super().__init__(master, text, command, bg="#F2F2F7", fg=TEXT,
                         padx=padx, pady=pady, bold=bold, **kw)

    def _draw(self, color):
        super()._draw(color)
        self.tag_lower("all")


class Card(tk.Canvas):
    """白色圆角卡片容器, 通过 content 获取内部 Frame; height 可固定卡片高度(内容place布局不贡献请求尺寸)"""
    def __init__(self, master, radius=18, pad=0, bg=CARD, height=None, **kw):
        self._radius, self._pad = radius, pad
        if height:
            kw["height"] = height
        super().__init__(master, highlightthickness=0, bd=0, bg=BG, **kw)
        self.content = tk.Frame(self, bg=bg)
        self.bind("<Configure>", self._redraw)

    def _redraw(self, _e=None):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 10 or h < 10:
            return
        self.create_rounded_rect(1, 1, w - 1, h - 1, self._radius, fill=CARD, outline="")
        self.content.place(x=self._pad, y=self._pad,
                           width=max(0, w - self._pad * 2),
                           height=max(0, h - self._pad * 2))

    def create_rounded_rect(self, x1, y1, x2, y2, r, **kw):
        pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
               x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
               x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
        return self.create_polygon(pts, smooth=True, **kw)


class SearchBox(tk.Canvas):
    """圆角搜索输入框"""
    def __init__(self, master, textvariable=None, radius=16, height=52, **kw):
        super().__init__(master, highlightthickness=0, bd=0, bg=CARD, height=height, **kw)
        self._radius = radius
        self.entry = tk.Entry(self, textvariable=textvariable, bg=INPUT_BG,
                              relief=tk.FLAT, highlightthickness=0,
                              font=("Microsoft YaHei", 13), fg=TEXT,
                              insertbackground=TEXT)
        self.bind("<Configure>", self._redraw)
        self._redraw()

    def _redraw(self, _e=None):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 10 or h < 10:
            return
        pts = [self._radius, 0, w - self._radius, 0, w, 0, w, self._radius,
               w, h - self._radius, w, h, w - self._radius, h,
               self._radius, h, 0, h, 0, h - self._radius, 0, self._radius, 0, 0]
        self.create_polygon(pts, smooth=True, fill=INPUT_BG, outline="")
        self.entry.place(x=18, y=0, width=max(0, w - 36), height=h)
        self.entry.configure(background=INPUT_BG)


# ============ GUI ============
class App:
    def __init__(self, root):
        self.root = root
        self.excel_info = {}
        self.excel_ordered = []
        self.disk_idx = {}
        self.disk_ready = False
        self.after_id = None
        self.selected_oid = None

        # 按真实 DPI 校正 Tk 缩放, 保证字体清晰锐利
        try:
            dpi = root.winfo_fpixels("1i")
            root.tk.call("tk", "scaling", dpi / 72.0)
        except Exception:
            pass

        root.title("苏州交复文件管理查询")
        # 窗口尺寸按屏幕与 DPI 自适应, 避免高缩放(125%/150%)下窗口超出屏幕
        try:
            scale = root.winfo_fpixels("1i") / 96.0
            sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
            w = min(1270, int(sw / scale * 0.98))
            h = min(780, int(sh / scale * 0.92))
        except Exception:
            w, h = 1150, 760
        root.geometry(f"{w}x{h}")
        root.minsize(960, 620)
        root.configure(bg=BG)
        icon = resource_path("icon.ico")
        if os.path.isfile(icon):
            try:
                root.iconbitmap(icon)
            except Exception:
                pass

        self._build_ui()

        # 加载配置: 无效则引导选择
        self.data_root, self.xlsx_path, _cfg = load_config()
        self.saved_col_widths = {str(k): int(v) for k, v in (_cfg.get("col_widths") or {}).items()}
        if not self.data_root or not self.xlsx_path:
            if not self._ask_paths_first_run():
                return
        self._load_data()
        self.status.config(text=f"数据目录：{self.data_root}\n跟进表：{self.xlsx_path}")

    # ---------- 界面 ----------
    def _build_ui(self):
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Apple.Treeview", background=CARD, fieldbackground=CARD,
                        foreground=TEXT, font=("Microsoft YaHei", 10),
                        rowheight=30, borderwidth=0)
        style.configure("Apple.Treeview.Heading", background="#F2F2F7", foreground=SUB,
                        font=("Microsoft YaHei", 10, "bold"), relief="flat",
                        borderwidth=0, padding=(6, 6))
        style.map("Apple.Treeview",
                  background=[("selected", ACCENT)],
                  foreground=[("selected", "white")])
        style.configure("Apple.Vertical.TScrollbar", background=SCROLL,
                        troughcolor=CARD, bordercolor=CARD,
                        arrowcolor=SUB, relief="flat", width=10)
        style.configure("Apple.Horizontal.TScrollbar", background=SCROLL,
                        troughcolor=CARD, bordercolor=CARD,
                        arrowcolor=SUB, relief="flat", height=12)
        # 去掉滚动条两端的箭头图标（仅保留滑轨与滑块, 简洁苹果风）
        try:
            style.layout("Apple.Vertical.TScrollbar",
                         [("Vertical.Scrollbar.trough", {"sticky": "ns", "children": [
                             ("Vertical.Scrollbar.thumb", {"sticky": "nswe", "expand": "1"})]})])
            style.layout("Apple.Horizontal.TScrollbar",
                         [("Horizontal.Scrollbar.trough", {"sticky": "ew", "children": [
                             ("Horizontal.Scrollbar.thumb", {"sticky": "nswe", "expand": "1"})]})])
        except tk.TclError:
            pass

        root = self.root

        # ---- 顶栏: 标题 + 操作按钮 + 搜索框（合并为一张卡, 紧凑） ----
        header = Card(root, radius=18, pad=14, height=126)
        header.pack(fill=tk.X, padx=16, pady=(12, 0))
        hc = header.content
        hc.columnconfigure(0, weight=1)
        tk.Label(hc, text="苏州交复文件管理查询", font=("Microsoft YaHei", 15, "bold"),
                 bg=CARD, fg=TEXT).grid(row=0, column=0, sticky="w")
        ops = tk.Frame(hc, bg=CARD)
        ops.grid(row=0, column=1, sticky="e", padx=(12, 0))
        self.btn_refresh = GhostButton(ops, "刷新数据", self._load_data, pady=5)
        self.btn_refresh.pack(side=tk.LEFT, padx=(0, 6))
        self.btn_settings = GhostButton(ops, "设置路径…", self._open_settings, pady=5)
        self.btn_settings.pack(side=tk.LEFT)
        sbox = SearchBox(hc, height=42)
        sbox.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        self.entry = sbox.entry
        self.entry.bind("<KeyRelease>", self._on_key)
        self.entry.bind("<Return>", self._on_enter)

        # ---- 中部: 订单列表（弹性占满） ----
        mid = tk.Frame(root, bg=BG)
        mid.pack(fill=tk.BOTH, expand=True, padx=12, pady=(10, 0))

        list_card = Card(mid, radius=18, pad=10)
        list_card.pack(fill=tk.BOTH, expand=True)
        lc = list_card.content
        self.tree_cols = ("oid", "otype", "cust", "prod", "data", "status", "time")
        self.tree_heads = ("订单编号", "订单类型", "客户", "产品类型", "数据名称", "当前状态", "接单时间")
        self.tree = ttk.Treeview(lc, columns=self.tree_cols, show="headings",
                                 selectmode="browse", style="Apple.Treeview")
        for c, h in zip(self.tree_cols, self.tree_heads):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=110, minwidth=50, stretch=False)
        vsb = ttk.Scrollbar(lc, orient="vertical", command=self.tree.yview,
                            style="Apple.Vertical.TScrollbar")
        self.hsb = ttk.Scrollbar(lc, orient="horizontal", command=self.tree.xview,
                                 style="Apple.Horizontal.TScrollbar")
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=self.hsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind("<Configure>", self._maybe_hsb)
        self.tree.bind("<Double-1>", lambda e: self.open_selected())
        self.tree.bind("<Return>", lambda e: self.open_selected())
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<ButtonRelease-1>", self._maybe_save_colwidths)

        # ---- 底部: 订单详情（横贯全宽, 紧凑） ----
        detail_card = Card(root, radius=18, pad=14, height=138)
        detail_card.pack(fill=tk.X, padx=16, pady=(10, 0))
        dc = detail_card.content
        for i in (1, 3, 5, 7):
            dc.columnconfigure(i, weight=1)
        tk.Label(dc, text="订单详情", font=("Microsoft YaHei", 10, "bold"),
                 bg=CARD, fg=SUB).grid(row=0, column=0, sticky="w", padx=(0, 12))
        self.d_oid = tk.Label(dc, text="—", font=("Microsoft YaHei", 13, "bold"),
                              bg=CARD, fg=TEXT, anchor="w")
        self.d_oid.grid(row=0, column=1, columnspan=5, sticky="w")
        self.btn_open = RoundButton(dc, "打开文件夹", self.open_selected,
                                    font=("Microsoft YaHei", 11, "bold"),
                                    padx=22, pady=6)
        self.btn_open.grid(row=0, column=8, rowspan=3, sticky="e", padx=(14, 0))

        self.field_vars = {}
        fields = [("客户", "cust"), ("产品", "prod"), ("数据", "data"), ("状态", "status"),
                  ("机器", "machine"), ("批次", "material"), ("快递", "express"), ("时间", "time")]
        for i, (lab, key) in enumerate(fields):
            r, c = divmod(i, 4)
            col = c * 2
            tk.Label(dc, text=lab, font=("Microsoft YaHei", 10), bg=CARD,
                     fg=SUB).grid(row=1 + r, column=col, sticky="nw", pady=2)
            var = tk.StringVar(value="—")
            tk.Label(dc, textvariable=var, font=("Microsoft YaHei", 10), bg=CARD,
                     fg=TEXT, anchor="w", justify="left", wraplength=110).grid(
                         row=1 + r, column=col + 1, sticky="nw", padx=(8, 14), pady=2)
            self.field_vars[key] = var

        self.d_path = tk.Label(dc, text="", font=("Microsoft YaHei", 9), bg=CARD,
                               fg=SUB, anchor="w", justify="left")
        self.d_path.grid(row=3, column=0, columnspan=8, sticky="ew", pady=(6, 0))

        # ---- 底部: 状态 + 署名 ----
        foot = tk.Frame(root, bg=BG)
        foot.pack(fill=tk.X, padx=16, pady=(8, 8))
        self.status = tk.Label(foot, text="", font=("Microsoft YaHei", 9),
                               fg=SUB, anchor="w", justify="left", bg=BG)
        self.status.pack(side=tk.LEFT, fill=tk.X)
        tk.Label(foot, text="by 阿果", font=("Microsoft YaHei", 10, "bold"),
                 fg=SUB, bg=BG).pack(side=tk.RIGHT)

        self.root.after(150, lambda: self.entry.focus_set())

    # ---------- 路径选择 ----------
    def _ask_paths_first_run(self):
        """首次运行/路径失效时, 引导选择"""
        msg = ("这是首次运行，或上次保存的路径已失效。\n\n"
               "请选择：\n"
               "① 数据根目录（存放 01~05 产品类型文件夹的目录，如 D:\\公司数据\\苏州交复）\n"
               "② 工作跟进表文件（.xlsx）")
        messagebox.showinfo("设置数据路径", msg)
        return self._open_settings(first=True)

    def _open_settings(self, first=False):
        dlg = tk.Toplevel(self.root)
        dlg.title("设置数据路径")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.configure(bg=BG)

        frm = tk.Frame(dlg, bg=BG, padx=20, pady=16)
        frm.pack(fill=tk.BOTH, expand=True)

        var_root = tk.StringVar(value=self.data_root or "")
        var_xlsx = tk.StringVar(value=self.xlsx_path or "")

        def browse_root():
            p = filedialog.askdirectory(title="选择数据根目录", initialdir=var_root.get() or DEFAULT_ROOT)
            if p:
                var_root.set(p)

        def browse_xlsx():
            p = filedialog.askopenfilename(title="选择工作跟进表文件", initialdir=var_xlsx.get() or DEFAULT_ROOT,
                                           filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")])
            if p:
                var_xlsx.set(p)

        tk.Label(frm, text="① 数据根目录（含 01~05 产品类型文件夹）：", font=("Microsoft YaHei", 10),
                 bg=BG, fg=TEXT).grid(row=0, column=0, sticky="w", pady=(0, 4))
        row1 = tk.Frame(frm, bg=BG)
        row1.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        row1.columnconfigure(0, weight=1)
        tk.Entry(row1, textvariable=var_root, font=("Microsoft YaHei", 10),
                 relief=tk.FLAT, highlightthickness=1, highlightbackground=LINE,
                 highlightcolor=ACCENT).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6), ipady=4)
        tk.Button(row1, text="浏览…", command=browse_root, relief=tk.FLAT,
                  bg="#F2F2F7", fg=TEXT, font=("Microsoft YaHei", 10), cursor="hand2",
                  activebackground=LINE, bd=0).pack(side=tk.RIGHT)

        tk.Label(frm, text="② 工作跟进表文件（.xlsx）：", font=("Microsoft YaHei", 10),
                 bg=BG, fg=TEXT).grid(row=2, column=0, sticky="w", pady=(0, 4))
        row2 = tk.Frame(frm, bg=BG)
        row2.grid(row=3, column=0, sticky="ew", pady=(0, 14))
        row2.columnconfigure(0, weight=1)
        tk.Entry(row2, textvariable=var_xlsx, font=("Microsoft YaHei", 10),
                 relief=tk.FLAT, highlightthickness=1, highlightbackground=LINE,
                 highlightcolor=ACCENT).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6), ipady=4)
        tk.Button(row2, text="浏览…", command=browse_xlsx, relief=tk.FLAT,
                  bg="#F2F2F7", fg=TEXT, font=("Microsoft YaHei", 10), cursor="hand2",
                  activebackground=LINE, bd=0).pack(side=tk.RIGHT)

        result = {"ok": False}

        def on_ok():
            r = var_root.get().strip()
            x = var_xlsx.get().strip()
            if not r or not os.path.isdir(r):
                messagebox.showwarning("路径无效", "数据根目录不存在，请重新选择。", parent=dlg)
                return
            if not x or not os.path.isfile(x):
                messagebox.showwarning("路径无效", "工作跟进表文件不存在，请重新选择。", parent=dlg)
                return
            result["ok"] = True
            self.data_root, self.xlsx_path = r, x
            save_config(r, x)
            dlg.destroy()

        btns = tk.Frame(frm, bg=BG)
        btns.grid(row=4, column=0, sticky="e")
        RoundButton(btns, "保存并加载", on_ok, padx=22, pady=6,
                    font=("Microsoft YaHei", 10, "bold")).pack(side=tk.LEFT, padx=(0, 8))
        GhostButton(btns, "取消", dlg.destroy, padx=16, pady=6).pack(side=tk.LEFT)

        dlg.update_idletasks()
        w, h = dlg.winfo_reqwidth() + 20, dlg.winfo_reqheight() + 20
        x = (dlg.winfo_screenwidth() - w) // 2
        y = (dlg.winfo_screenheight() - h) // 3
        dlg.geometry(f"{w}x{h}+{x}+{y}")
        self.root.wait_window(dlg)
        if result["ok"]:
            self._load_data()
        return result["ok"]

    # ---------- 数据加载 ----------
    def _load_data(self):
        if not self.data_root or not self.xlsx_path:
            return
        self.status.config(text="正在加载数据…")
        self.root.update_idletasks()

        def worker():
            try:
                info, ordered = load_excel_orders(self.xlsx_path)
                idx = scan_disk_orders(self.data_root)
                self.root.after(0, lambda: self._on_data_loaded(info, ordered, idx))
            except Exception as e:
                self.root.after(0, lambda: self._on_load_error(e))

        threading.Thread(target=worker, daemon=True).start()

    def _on_load_error(self, e):
        messagebox.showerror("加载失败", f"无法读取跟进表或数据目录：\n{e}")
        self.status.config(text="加载失败，请检查 设置路径…")

    def _on_data_loaded(self, info, ordered, idx):
        self.excel_info = info
        self.excel_ordered = ordered
        self.disk_idx = idx
        n_excel = len(info)
        self.disk_ready = True
        if self.saved_col_widths:
            # 有手动调整过的列宽记录 → 沿用
            for c in self.tree_cols:
                w = self.saved_col_widths.get(c)
                if w:
                    self.tree.column(c, width=int(w), minwidth=50, stretch=False)
            self.root.after(100, self._maybe_hsb)
        else:
            # 首次使用 → 窗口映射、tk scaling 稳定后按最长内容计算列宽
            self.root.after(300, self._autosize_columns)
        self._refresh_list()
        self.root.after(0, lambda: self.status.config(
            text=f"数据目录：{self.data_root}\n跟进表：{self.xlsx_path}\n"
                 f"Excel 订单 {n_excel} 条   |   磁盘 3DP 订单文件夹 {len(idx)} 个   |   输入订单号/客户/患者名后回车或双击打开"))

    def _autosize_columns(self):
        """订单列表列宽按该列最长内容自适应（含表头与内边距）"""
        if not hasattr(self, "tree") or not self.tree.winfo_exists():
            return
        try:  # 以当前真实 DPI 校准, 保证测量与实际渲染一致
            dpi = self.root.winfo_fpixels("1i")
            self.root.tk.call("tk", "scaling", dpi / 72.0)
        except Exception:
            pass
        f = tkfont.Font(root=self.root, family="Microsoft YaHei", size=10)
        fh = tkfont.Font(root=self.root, family="Microsoft YaHei", size=10, weight="bold")
        maxw = [fh.measure(h) + 12 for h in self.tree_heads]
        rows = self.excel_ordered or []
        for x in rows:
            for i, c in enumerate(self.tree_cols):
                v = str(x.get(c, "") or "")
                if c == "time":  # 列表展示截取日期部分, 按展示值测量
                    v = v[:10]
                w = f.measure(v) + 12
                if w > maxw[i]:
                    maxw[i] = w
        widths = [max(60, min(w, 245 if c == "data" else 340)) for c, w in zip(self.tree_cols, maxw)]
        # 若总宽不足可视区, 把剩余空间均匀摊给各列, 列表铺满不留白
        try:
            self.root.update_idletasks()
            avail = self.tree.winfo_width()
            if avail > sum(widths) and avail > 0:
                extra = (avail - sum(widths)) // len(widths)
                widths = [w + extra for w in widths]
        except Exception:
            pass
        for c, w in zip(self.tree_cols, widths):
            self.tree.column(c, width=w, minwidth=60, stretch=False)

    def _maybe_save_colwidths(self, evt):
        """手动拖动列分隔线调整列宽后, 记住宽度到 config.json（下次启动沿用）"""
        try:
            # 仅表头区域(顶部约50px)的点击/拖动才涉及列宽调整
            if evt.y > 50:
                return
            widths = {c: int(self.tree.column(c, "width")) for c in self.tree_cols}
            self.saved_col_widths = widths
            save_config(self.data_root, self.xlsx_path, widths)
        except Exception:
            pass

    def _maybe_hsb(self, _e=None):
        """仅当列总宽超出可视区时才显示水平滚动条, 平时隐藏不占空间"""
        try:
            need = sum(self.tree.column(c, "width") for c in self.tree_cols) > self.tree.winfo_width()
            if need and not self.hsb.winfo_manager():
                self.hsb.pack(side=tk.BOTTOM, fill=tk.X)
            elif not need and self.hsb.winfo_manager():
                self.hsb.pack_forget()
        except Exception:
            pass

    # ---------- 搜索 ----------
    def _on_key(self, _evt):
        if self.after_id:
            self.root.after_cancel(self.after_id)
        self.after_id = self.root.after(150, self._refresh_list)

    def _on_enter(self, _evt):
        self._refresh_list()
        self.open_selected()

    def _match(self, kw):
        kw = kw.strip().lower()
        if not kw:
            return self.excel_ordered[:200]
        # 含ASCII字母/数字 -> 按订单号模糊匹配（支持 完整号/日期段/序号 等）
        has_ascii = any(c.isascii() and c.isalnum() for c in kw)
        if has_ascii:
            hits = [x for x in self.excel_ordered if kw in x["oid"].lower()]
            if hits:
                return hits[:200]
            # 磁盘有但Excel没有（新文件夹）→ 按磁盘索引兜底
            disk_hits = [oid for oid in self.disk_idx if kw in oid.lower()]
            return [{"oid": oid, "otype": "", "cust": "", "prod": PREFIX_MAP.get(oid[:7], ("", ""))[0],
                     "data": "", "status": "（磁盘未登记）", "time": ""} for oid in sorted(disk_hits)[:200]]
        # 否则按客户名 或 患者名(数据名称) 匹配
        hits = [x for x in self.excel_ordered
                if kw in x["cust"].lower() or kw in x["data"].lower()]
        return hits[:200]

    def _refresh_list(self):
        self.tree.delete(*self.tree.get_children())
        kw = self.entry.get()
        rows = self._match(kw)
        if not rows:
            # 无匹配时清空选中, 避免回车误操作上一次选中的订单
            self.selected_oid = None
            self.d_oid.config(text="—")
            for var in self.field_vars.values():
                var.set("—")
            self.d_path.config(text="未找到匹配的订单。\n可输入：订单号（完整/部分）、客户名称，或患者姓名。")
            return
        for x in rows:
            self.tree.insert("", "end", values=(
                x["oid"], x.get("otype", ""), x["cust"], x["prod"], x["data"], x["status"],
                str(x["time"])[:10] if x.get("time") else ""))
        first = self.tree.get_children()[0]
        self.tree.selection_set(first)
        self.tree.focus(first)
        self._show_detail(first)

    def _on_tree_select(self, _evt):
        sel = self.tree.selection()
        if sel:
            self._show_detail(sel[0])

    def _show_detail(self, iid):
        vals = self.tree.item(iid, "values")
        oid = vals[0]
        self.selected_oid = oid
        x = self.excel_info.get(oid, {})
        path = self.disk_idx.get(oid)
        if path is None:
            path = expected_path(oid, self.excel_info, self.data_root)
        self.d_oid.config(text=oid)
        for key, var in self.field_vars.items():
            v = x.get(key)
            if key == "time" and v:
                v = str(v)[:10]
            var.set(v or "—")
        if path:
            self.d_path.config(text=f"文件夹：{path}", fg=SUB)
        else:
            self.d_path.config(text="文件夹：磁盘上未找到（可能尚未创建）", fg="#C0392B")

    # ---------- 打开 ----------
    def open_selected(self):
        if not self.selected_oid:
            messagebox.showinfo("未选择订单", "请先在列表中选中一个订单，再点击“打开文件夹”。")
            return
        oid = self.selected_oid
        path = self.disk_idx.get(oid)
        if path is None:
            path = expected_path(oid, self.excel_info, self.data_root)
        if path and os.path.isdir(path):
            try:
                os.startfile(path)
                self.status.config(text=f"已打开：{path}")
            except Exception as e:
                messagebox.showerror("打开失败", f"无法打开文件夹：\n{path}\n{e}")
        else:
            x = self.excel_info.get(oid, {})
            sub = PREFIX_MAP.get(oid[:7], (None, None))[1]
            guess = os.path.join(self.data_root, sub, x.get("cust", "?"), oid) if sub else None
            msg = f"订单 {oid} 的文件夹在磁盘上未找到。\n\n预期位置：\n{guess}\n\n可能原因：文件夹尚未创建，或客户名与表格不一致。"
            if x.get("cust"):
                msg += f"\n\n提示：可到 “{self.data_root}\\{sub}” 下按客户 “{x.get('cust')}” 查找。"
            messagebox.showinfo("文件夹未找到", msg)


def main():
    enable_dpi_awareness()  # 必须在创建窗口前声明 DPI 感知, 避免模糊
    try:
        tkroot = tk.Tk()
        App(tkroot)
        tkroot.mainloop()
    except Exception:
        _write_error_log(f"启动异常：\n{traceback.format_exc()}")
        try:  # 尝试弹窗告知, 弹不出至少日志已落盘
            import tkinter.messagebox as _mb
            _mb.showerror("程序启动异常",
                          "程序运行出错，请将程序目录下的 error.log 文件发送给开发者排查。")
        except Exception:
            pass
        raise


# ============ 自测模式 ============
def selftest():
    print("=" * 66)
    print("自测：配置读取")
    root, xlsx, _ = load_config()
    if not root:
        root, xlsx = DEFAULT_ROOT, DEFAULT_XLSX
        print("  (无配置文件, 使用默认路径)")
    print(f"  根目录: {root}")
    print(f"  跟进表: {xlsx}")

    print("=" * 66)
    print("自测：读取 Excel 订单跟踪表")
    info, ordered = load_excel_orders(xlsx)
    print(f"  3DP 订单总数: {len(info)}")
    print(f"  前缀分布: {dict(Counter(o[:7] for o in info))}")

    print("=" * 66)
    print("自测：全量扫描磁盘 3DP 订单文件夹")
    idx = scan_disk_orders(root)
    print(f"  磁盘 3DP 订单文件夹: {len(idx)}")
    print(f"  前缀分布: {dict(Counter(o[:7] for o in idx))}")

    print("=" * 66)
    print("自测：命中核对")
    hit = sum(1 for o in info if o in idx)
    miss = [o for o in info if o not in idx]
    print(f"  Excel订单 直接命中磁盘文件夹: {hit}/{len(info)}")
    print(f"  未命中 {len(miss)} 条: {miss[:10]}")

    print("=" * 66)
    print("自测：搜索逻辑")
    for kw in ["3DP-101-20260401-001", "20260401-001", "001", "测试客户甲", "测试客户乙", "3dp-105", "示例患者甲", "示例患者乙"]:
        hits = [x for x in ordered
                if kw.lower() in x["oid"].lower() or kw.lower() in x["cust"].lower() or kw.lower() in x["data"].lower()]
        print(f"  关键词 [{kw}] -> {len(hits)} 条, 明细:")
        for h in hits[:5]:
            print(f"      {h['oid']}  客户={h['cust']}  数据={h['data']}  磁盘文件夹={'有' if h['oid'] in idx else '无'}")
    print("自测完成 ✓")


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        selftest()
    else:
        main()
