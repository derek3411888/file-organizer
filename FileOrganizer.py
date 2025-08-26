# -*- coding: utf-8 -*-
import os, sys, shutil, ctypes, time, threading, hashlib, stat, datetime, re, json, tempfile, urllib.request, urllib.error, subprocess

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# ========= 基本資訊（請依你的 GitHub 倉庫調整） =========
APP_NAME = "FileOrganizer"
APP_VERSION = "1.0.1"   # 發版時改這個
MANIFEST_URL = "https://raw.githubusercontent.com/derek3411888/file-organizer/refs/heads/main/manifest.json"
UPDATE_INFO_URL = MANIFEST_URL   # ← 新增：讓下面函式用到的名稱一致

DEBUG_UPDATE = True  # 開發期先設 True（會顯示黑窗與暫停）；穩定後改 False
UPDATE_LOG = os.path.join(tempfile.gettempdir(), "FileOrganizer_update.log")



# 自動更新相關
AUTO_CHECK_ON_START = True   # 啟動時自動檢查
HTTP_TIMEOUT = 20            # 下載逾時秒數



# ====== 可用代碼說明 ======
CUSTOM_HELP = r"""
可用代碼（可自由組合，大小寫固定）：
  {檔名}                     原始檔名，含副檔名（例：photo.jpg）
  {檔名不含副檔名}           不含副檔名（例：photo）
  {副檔名}                   原始副檔名「含點」（例：.jpg）
  {副檔名不含點}             副檔名「不含點」（例：jpg）
  {副檔名大寫}               含點，大寫（例：.JPG）
  {副檔名大寫不含點}         不含點，大寫（例：JPG）
  {副檔名小寫}               含點，小寫（例：.jpg）
  {副檔名小寫不含點}         不含點，小寫（例：jpg）
  {父資料夾}                 來源檔案上一層資料夾名稱
  {根目錄}                   目前所處的根目錄名稱

日期時間（可加 strftime 格式）：
  {建立日期}                 預設 YYYY-MM-DD
  {修改日期}                 預設 YYYY-MM-DD
  {建立日期:%Y%m%d}          例：20240825
  {修改日期:%Y%m%d-%H%M%S}   例：20240825-153045
  {ctime:%Y-%m-%d_%H-%M-%S}  建立時間自訂格式
  {mtime:%Y-%m-%d_%H-%M-%S}  修改時間自訂格式

規則：
  • 預設一律保留原始副檔名。
  • 只有勾選「變更副檔名」且填寫新副檔名時，才會改副檔名。
  • 「整個檔名改為客製」也會自動補上正確副檔名（遵守上面規則）。
"""

CODE_DEFS = [
    ("{檔名}", "原始檔名，含副檔名", "{檔名}"),
    ("{檔名不含副檔名}", "不含副檔名的檔名", "{檔名不含副檔名}"),
    ("{副檔名}", "原始副檔名，含點", "{副檔名}"),
    ("{副檔名不含點}", "副檔名，不含點", "{副檔名不含點}"),
    ("{副檔名大寫}", "副檔名含點，轉大寫", "{副檔名大寫}"),
    ("{副檔名大寫不含點}", "副檔名不含點，轉大寫", "{副檔名大寫不含點}"),
    ("{副檔名小寫}", "副檔名含點，轉小寫", "{副檔名小寫}"),
    ("{副檔名小寫不含點}", "副檔名不含點，轉小寫", "{副檔名小寫不含點}"),
    ("{父資料夾}", "來源檔案上一層資料夾名稱", "{父資料夾}"),
    ("{根目錄}", "Tree 清單的根目錄名稱", "{根目錄}"),
    ("{建立日期}", "建立日期，預設 YYYY-MM-DD", "{建立日期}"),
    ("{修改日期}", "修改日期，預設 YYYY-MM-DD", "{修改日期}"),
    ("{建立日期:%Y%m%d}", "建立日期自訂格式", "{建立日期:%Y%m%d}"),
    ("{修改日期:%Y%m%d-%H%M%S}", "修改日期自訂格式", "{修改日期:%Y%m%d-%H%M%S}"),
    ("{ctime:%Y-%m-%d_%H-%M-%S}", "建立時間自訂格式", "{ctime:%Y-%m-%d_%H-%M-%S}"),
    ("{mtime:%Y-%m-%d_%H-%M-%S}", "修改時間自訂格式", "{mtime:%Y-%m-%d_%H-%M-%S}"),
]

# ---------- 提權（pythonw 無黑框） ----------
def ensure_admin():
    try:
        if ctypes.windll.shell32.IsUserAnAdmin():
            return
    except Exception:
        pass
    params = " ".join(f'"{a}"' for a in sys.argv)
    pyw = sys.executable.replace("python.exe", "pythonw.exe")
    ctypes.windll.shell32.ShellExecuteW(None, "runas", pyw, params, None, 1)
    sys.exit()

# ---------- 常量 ----------
IGNORABLE_FILES = {"thumbs.db", "desktop.ini", ".ds_store"}
CATEGORY_MAP = {
    "影片": {"mp4","mkv","avi","mov","wmv","flv","m4v","ts","webm"},
    "音訊": {"mp3","wav","flac","m4a","aac","ogg","wma","aiff","alac"},
    "影像": {"jpg","jpeg","png","gif","bmp","tiff","tif","webp","heic","raw","cr2","nef","arw"},
    "PDF": {"pdf"},
    "文件": {"txt","md","rtf","doc","docx","odt","epub"},
    "表格": {"xls","xlsx","csv","ods"},
    "簡報": {"ppt","pptx","odp","key"},
    "壓縮": {"zip","rar","7z","tar","gz","bz2"},
    "程式碼": {"py","pyw","js","ts","java","c","cpp","cs","go","rs","rb","php","sh","bat","ps1","html","css","json","xml","yaml","yml"},
    "字型": {"ttf","otf","woff","woff2"},
}
PROTECTED_EXTS = {'.exe','.dll','.sys','.msi','.lnk','.bat','.cmd','.ps1'}

# ---------- 基礎工具 ----------
def unique_path(dest_path):
    if not os.path.exists(dest_path): return dest_path
    base, ext = os.path.splitext(dest_path); i = 1
    while True:
        cand = f"{base} ({i}){ext}"
        if not os.path.exists(cand): return cand
        i += 1

def remove_readonly(p):
    try: os.chmod(p, stat.S_IWRITE | stat.S_IREAD)
    except: pass

def md5_8(path, chunk_mb=4):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(chunk_mb*1024*1024), b""): h.update(chunk)
    return h.hexdigest()[:8]

def full_md5(path, chunk_mb=4):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(chunk_mb*1024*1024), b""): h.update(chunk)
    return h.hexdigest()

def ext_key(path_or_name):
    e = os.path.splitext(path_or_name)[1].lower()
    return (e if e else "._noext").lstrip(".")

def prune_empty_dirs(root, dry_run=False, progress_cb=None):
    removed_total = 0
    while True:
        removed = 0
        for folder, _, files in os.walk(root, topdown=False):
            if os.path.normcase(folder) == os.path.normcase(root): continue
            for name in list(files):
                if name.lower() in IGNORABLE_FILES:
                    fp = os.path.join(folder, name)
                    try:
                        if not dry_run: remove_readonly(fp); os.remove(fp)
                    except: pass
            try:
                if not os.listdir(folder):
                    if not dry_run: remove_readonly(folder); os.rmdir(folder)
                    removed += 1; removed_total += 1
                    if progress_cb: progress_cb(removed_total, 0, f"移除空資料夾：{folder}")
            except: pass
        if removed == 0: break
    return removed_total

def gather_all_files(root, include_root):
    out = []
    for folder, _, files in os.walk(root):
        if not include_root and os.path.normcase(folder) == os.path.normcase(root):
            continue
        for name in files:
            out.append(os.path.join(folder, name))
    return out

def scan_exts(root):
    exts = set()
    for folder, _, files in os.walk(root):
        for name in files:
            exts.add(ext_key(name))
    if "_noext" in exts: exts.remove("_noext")
    return exts or {"_noext"}

def categorize_exts(exts):
    cats = {k: [] for k in CATEGORY_MAP}
    cats["其他"] = []; cats["無副檔名"] = []
    for e in sorted(exts):
        if e == "_noext": cats["無副檔名"].append(e); continue
        placed = False
        for cat, s in CATEGORY_MAP.items():
            if e in s: cats[cat].append(e); placed = True; break
        if not placed: cats["其他"].append(e)
    return {k: v for k, v in cats.items() if v}

# ---------- 代碼渲染 ----------
def render_tokens(src_path, template, root_name=None):
    st = os.stat(src_path)
    ctime = datetime.datetime.fromtimestamp(st.st_ctime)
    mtime = datetime.datetime.fromtimestamp(st.st_mtime)
    parent = os.path.basename(os.path.dirname(src_path))
    base = os.path.basename(src_path)
    stem, ext = os.path.splitext(base)

    ext_with_dot = ext
    ext_nodot = ext[1:] if ext.startswith(".") else ext
    ext_up_with_dot = ext_with_dot.upper()
    ext_up_nodot = ext_nodot.upper()
    ext_lo_with_dot = ext_with_dot.lower()
    ext_lo_nodot = ext_nodot.lower()

    out = (template.replace("{檔名}", base)
                    .replace("{檔名不含副檔名}", stem)
                    .replace("{副檔名}", ext_with_dot)
                    .replace("{副檔名不含點}", ext_nodot)
                    .replace("{副檔名大寫}", ext_up_with_dot)
                    .replace("{副檔名大寫不含點}", ext_up_nodot)
                    .replace("{副檔名小寫}", ext_lo_with_dot)
                    .replace("{副檔名小寫不含點}", ext_lo_nodot)
                    .replace("{父資料夾}", parent)
                    .replace("{根目錄}", root_name or "")
                    .replace("{建立日期}", ctime.strftime("%Y-%m-%d"))
                    .replace("{修改日期}", mtime.strftime("%Y-%m-%d")))
    def repl_dt(m):
        key, fmt = m.group(1), m.group(2) or "%Y-%m-%d_%H-%M-%S"
        dt = ctime if key in ("建立日期","ctime") else mtime
        return dt.strftime(fmt)
    out = re.sub(r"\{(建立日期|修改日期|ctime|mtime):([^}]+)\}", repl_dt, out)
    return out

# ---------- 目的路徑決策（保留或改副檔名） ----------
def _apply_final_ext(target_name, orig_ext, change_ext, new_ext_input):
    final_ext = orig_ext
    if change_ext:
        ne = (new_ext_input or "").strip()
        if ne:
            if not ne.startswith("."): ne = "." + ne
            final_ext = ne
    base, _ = os.path.splitext(target_name)
    return base + final_ext

def decide_dest(src_path, dest_dir, dup_mode, custom_mode, custom_text, change_ext=False, new_ext_input="", root_name=None):
    name = os.path.basename(src_path)
    stem, ext = os.path.splitext(name)

    if custom_mode == "prefix":
        piece = render_tokens(src_path, custom_text, root_name=root_name)
        target_name = f"{piece}{stem}{ext}"
    elif custom_mode == "suffix":
        piece = render_tokens(src_path, custom_text, root_name=root_name)
        target_name = f"{stem}{piece}{ext}"
    elif custom_mode == "replace":
        body = render_tokens(src_path, custom_text, root_name=root_name)
        target_name = _apply_final_ext(body, ext, change_ext, new_ext_input)
    else:
        target_name = _apply_final_ext(name, ext, change_ext, new_ext_input)

    if custom_mode in ("prefix", "suffix"):
        target_name = _apply_final_ext(target_name, ext, change_ext, new_ext_input)

    target = os.path.join(dest_dir, target_name)

    if dup_mode == "skip": return None if os.path.exists(target) else target
    if dup_mode == "overwrite": return ("overwrite", target)
    if dup_mode == "index": return unique_path(target)
    if dup_mode == "datetime":
        stamp = time.strftime("%Y%m%d-%H%M%S")
        b0, e0 = os.path.splitext(target_name)
        return unique_path(os.path.join(dest_dir, f"{b0}_{stamp}{e0}"))
    if dup_mode == "hash8":
        try: h8 = md5_8(src_path)
        except: h8 = "unknown"
        b0, e0 = os.path.splitext(target_name)
        return unique_path(os.path.join(dest_dir, f"{b0}_{h8}{e0}"))
    return unique_path(target)

def organize(root, include_root, exts_filter, dup_mode, custom_mode, custom_text,
             change_ext, new_ext_input, dry_run, progress_cb, verbose_cb, *,
             flatten_all=False):
    all_files = gather_all_files(root, include_root=include_root)
    if exts_filter != "ALL":
        all_files = [p for p in all_files if (ext_key(p) in exts_filter or ("_noext" in exts_filter and ext_key(p)=="_noext"))]
    total = len(all_files); done = 0
    progress_cb(root, done, total, f"發現 {total} 檔")
    root_name = os.path.basename(root) or root
    for src in all_files:
        e = ext_key(src)
        if flatten_all:
            dest_dir = root
        else:
            dest_dir = os.path.join(root, e if e != "_noext" else "_noext")
        try:
            if not dry_run and not os.path.exists(dest_dir): os.makedirs(dest_dir, exist_ok=True)
            decision = decide_dest(src, dest_dir, dup_mode, custom_mode, custom_text, change_ext, new_ext_input, root_name=root_name)
            if decision is None:
                verbose_cb(f"[SKIP] 已存在同名 → {src}")
            else:
                do_overwrite = False; dest = decision
                if isinstance(decision, tuple) and decision[0] == "overwrite":
                    do_overwrite = True; dest = decision[1]
                if os.path.normcase(src) == os.path.normcase(dest):
                    verbose_cb(f"[SKIP] 來源=目的 → {src}")
                else:
                    if not dry_run:
                        if do_overwrite and os.path.exists(dest):
                            try: remove_readonly(dest); os.remove(dest)
                            except: dest = unique_path(dest)
                        shutil.move(src, dest)
                    verbose_cb(f"[OK] {src}  →  {dest}")
        except Exception as e:
            verbose_cb(f"[FAIL] {src} → {e}")
        done += 1
        progress_cb(root, done, total, f"搬移中… {done}/{total}")
    progress_cb(root, done, total, "刪除空資料夾…")
    prune_empty_dirs(root, dry_run=dry_run)
    progress_cb(root, total, total, "完成")

# ---------- Windows 回收桶刪除 ----------
FO_DELETE = 3
FOF_ALLOWUNDO = 0x0040
FOF_NOCONFIRMATION = 0x0010
FOF_SILENT = 0x0004
FOF_NOERRORUI = 0x0400
class SHFILEOPSTRUCTW(ctypes.Structure):
    _fields_ = [
        ("hwnd", ctypes.c_void_p),
        ("wFunc", ctypes.c_uint),
        ("pFrom", ctypes.c_wchar_p),
        ("pTo", ctypes.c_wchar_p),
        ("fFlags", ctypes.c_uint),
        ("fAnyOperationsAborted", ctypes.c_bool),
        ("hNameMappings", ctypes.c_void_p),
        ("lpszProgressTitle", ctypes.c_wchar_p),
    ]
def _paths_to_double_null(paths):
    return "\x00".join(paths) + "\x00\x00"
def move_to_recycle_bin_one(p):
    op = SHFILEOPSTRUCTW()
    op.hwnd = None; op.wFunc = FO_DELETE
    op.pFrom = _paths_to_double_null([p]); op.pTo = None
    op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT
    rc = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
    return rc == 0 and not op.fAnyOperationsAborted

# ---------- 簡易勾選清單（已移除「拖移框選」） ----------
class SimpleCheckList(ttk.Frame):
    def __init__(self, master, **kw):
        super().__init__(master)
        self.lb = tk.Listbox(self, selectmode="extended")
        self.lb.pack(fill="both", expand=True)
        self.items = []
        self.checked = set()
        # 單擊切換、空白鍵切換、滾輪捲動
        self.lb.bind("<Button-1>", self.on_click)
        self.lb.bind("<space>", self.toggle_active)
        self.lb.bind("<MouseWheel>", lambda ev: self.lb.yview_scroll(-1 if ev.delta>0 else 1, "units"))

    def set_items(self, items, checked_set=None):
        self.items = list(items)
        self.checked = set(checked_set or [])
        self.lb.delete(0, "end")
        for i,s in enumerate(self.items):
            self.lb.insert("end", self._line(i))
        self.lb.update_idletasks()

    def get_checked_items(self):
        return [self.items[i] for i in sorted(self.checked)]

    def check_all(self, flag):
        if flag: self.checked = set(range(len(self.items)))
        else: self.checked.clear()
        self._render()

    def _line(self, i):
        return ("☑ " if i in self.checked else "☐ ") + self.items[i]

    def _render(self):
        self.lb.delete(0, "end")
        for i in range(len(self.items)):
            self.lb.insert("end", self._line(i))

    def toggle_index(self, i, to=None):
        if i < 0 or i >= len(self.items): return
        if to is None: to = not (i in self.checked)
        if to: self.checked.add(i)
        else: self.checked.discard(i)
        self.lb.delete(i); self.lb.insert(i, self._line(i))

    def toggle_active(self, ev):
        i = self.lb.index("active")
        self.toggle_index(i)
        return "break"

    def on_click(self, ev):
        i = self.lb.nearest(ev.y)
        self.toggle_index(i)
        return "break"

# ---------- Tooltip ----------
class TreeTooltip:
    def __init__(self, tree, get_text_by_iid):
        self.tree = tree; self.get_text = get_text_by_iid
        self.tip = None; self.after_id = None
        tree.bind("<Motion>", self.on_motion); tree.bind("<Leave>", self.hide)
    def on_motion(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid: self.hide(); return
        text = self.get_text(iid)
        if not text: self.hide(); return
        if self.after_id: self.tree.after_cancel(self.after_id)
        self.after_id = self.tree.after(400, lambda: self.show(event.x_root+12, event.y_root+12, text))
    def show(self, x, y, text):
        self.hide()
        self.tip = tk.Toplevel(self.tree); self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        ttk.Label(self.tip, text=text, background="#ffffe0", relief="solid", borderwidth=1).pack(ipadx=6, ipady=3)
    def hide(self, *_):
        if self.after_id: self.tree.after_cancel(self.after_id); self.after_id = None
        if self.tip: self.tip.destroy(); self.tip = None

# ---------- 根目錄設定（改成「點擊/空白鍵」＋按鈕全選） ----------
class RootConfigDialog(tk.Toplevel):
    def __init__(self, master, root_path, cfg):
        super().__init__(master)
        self.title(os.path.basename(root_path) or root_path)
        self.resizable(True, True)
        self.cfg = cfg
        pad = 8

        frm1 = ttk.Frame(self, padding=pad); frm1.pack(fill="x")
        self.var_inc = tk.BooleanVar(value=self.cfg.get('include_root', False))
        ttk.Checkbutton(frm1, text="包含根目錄檔案", variable=self.var_inc).pack(anchor="w")

        self.var_flat = tk.BooleanVar(value=self.cfg.get('flatten_all', False))
        ttk.Checkbutton(frm1, text="忽略分類，全部放到根目錄", variable=self.var_flat).pack(anchor="w", pady=(4,0))

        frm2 = ttk.Frame(self, padding=(pad,0,pad,pad)); frm2.pack(fill="both", expand=True)
        ttk.Label(frm2, text="選擇要整理的檔案類型（分頁；單擊切換；空白鍵切換）：").pack(anchor="w")

        self.var_all = tk.BooleanVar(value=(self.cfg.get('exts_selected',"ALL")=="ALL"))
        ttk.Checkbutton(frm2, text="全部類型", variable=self.var_all, command=self.toggle_all).pack(anchor="w", pady=(0,4))

        cats = categorize_exts(self.cfg['exts_all'])
        self.nb = ttk.Notebook(frm2); self.nb.pack(fill="both", expand=True)
        self.checklists = {}
        self.all_exts_sorted = {}

        for cat, items in cats.items():
            tab = ttk.Frame(self.nb); self.nb.add(tab, text=cat)
            cbar = ttk.Frame(tab); cbar.pack(fill="x")
            cl = SimpleCheckList(tab); cl.pack(fill="both", expand=True)
            items_sorted = sorted(items); self.all_exts_sorted[cat] = items_sorted
            checked = set()
            if isinstance(self.cfg.get('exts_selected'), set):
                chosen = self.cfg['exts_selected']; checked = {i for i,x in enumerate(items_sorted) if x in chosen}
            cl.set_items(items_sorted, checked)
            ttk.Button(cbar, text="全選本頁", command=lambda C=cl: C.check_all(True)).pack(side="left", padx=4, pady=2)
            ttk.Button(cbar, text="全不選本頁", command=lambda C=cl: C.check_all(False)).pack(side="left", padx=4, pady=2)
            self.checklists[cat] = cl

        self.toggle_all()
        self.geometry("720x640")

        frm_btn = ttk.Frame(self, padding=pad); frm_btn.pack(fill="x")
        ttk.Button(frm_btn, text="確定", command=self.ok).pack(side="right", padx=4)
        ttk.Button(frm_btn, text="取消", command=self.destroy).pack(side="right")

    def toggle_all(self):
        state = "disabled" if self.var_all.get() else "normal"
        for cl in self.checklists.values():
            cl.lb.config(state=state)

    def ok(self):
        self.cfg['include_root'] = self.var_inc.get()
        self.cfg['flatten_all'] = self.var_flat.get()
        if self.var_all.get():
            self.cfg['exts_selected'] = "ALL"
        else:
            sel = set()
            for cl in self.checklists.values():
                for x in cl.get_checked_items(): sel.add(x)
            self.cfg['exts_selected'] = sel
        self.destroy()

# ---------- 批次設定（所有根目錄） ----------
class BulkConfigDialog(tk.Toplevel):
    def __init__(self, master, paths, cfg_map):
        super().__init__(master)
        self.title("設定所有根目錄")
        self.resizable(True, True)
        self.cfg_map = cfg_map; self.paths = paths
        pad = 8

        all_exts = set()
        for p in paths: all_exts |= cfg_map[p]['exts_all']

        frm1 = ttk.Frame(self, padding=pad); frm1.pack(fill="x")
        self.var_inc = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm1, text="包含根目錄檔案（套用到全部）", variable=self.var_inc).pack(anchor="w")
        self.var_flat = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm1, text="忽略分類，全部放到根目錄（套用到全部）", variable=self.var_flat).pack(anchor="w", pady=(4,0))

        frm2 = ttk.Frame(self, padding=(pad,0,pad,pad)); frm2.pack(fill="both", expand=True)
        ttk.Label(frm2, text="選擇要整理的檔案類型（留空=全部；單擊切換；空白鍵切換）：").pack(anchor="w")

        cats = categorize_exts(all_exts)
        self.nb = ttk.Notebook(frm2); self.nb.pack(fill="both", expand=True)
        self.checklists = {}; self.all_exts_sorted = {}

        for cat, items in cats.items():
            tab = ttk.Frame(self.nb); self.nb.add(tab, text=cat)
            cbar = ttk.Frame(tab); cbar.pack(fill="x")
            cl = SimpleCheckList(tab); cl.pack(fill="both", expand=True)
            items_sorted = sorted(items); self.all_exts_sorted[cat] = items_sorted
            cl.set_items(items_sorted, checked_set=set())
            ttk.Button(cbar, text="全選本頁", command=lambda C=cl: C.check_all(True)).pack(side="left", padx=4, pady=2)
            ttk.Button(cbar, text="全不選本頁", command=lambda C=cl: C.check_all(False)).pack(side="left", padx=4, pady=2)
            self.checklists[cat] = cl

        frm_btn = ttk.Frame(self, padding=pad); frm_btn.pack(fill="x")
        ttk.Button(frm_btn, text="確定", command=self.ok).pack(side="right", padx=4)
        ttk.Button(frm_btn, text="取消", command=self.destroy).pack(side="right")
        self.geometry("720x640")

    def ok(self):
        sel = set()
        for _, cl in self.checklists.items():
            for x in cl.get_checked_items(): sel.add(x)
        for p in self.paths:
            self.cfg_map[p]['include_root'] = self.var_inc.get()
            self.cfg_map[p]['flatten_all'] = self.var_flat.get()
            self.cfg_map[p]['exts_selected'] = sel if sel else "ALL"
        self.destroy()

# ---------- 代碼清單面板 ----------
class CodePalette(tk.Toplevel):
    def __init__(self, master, insert_cb):
        super().__init__(master)
        self.title("代碼清單")
        self.geometry("820x560")
        self.resizable(True, True)
        self.insert_cb = insert_cb

        top = ttk.Frame(self, padding=8); top.pack(fill="x")
        ttk.Label(top, text="搜尋：").pack(side="left")
        self.var_q = tk.StringVar()
        ent = ttk.Entry(top, textvariable=self.var_q, width=28); ent.pack(side="left", padx=6)
        ttk.Button(top, text="清除", command=lambda: (self.var_q.set(""), self.refresh())).pack(side="left")
        ttk.Button(top, text="使用說明…", command=self.show_help).pack(side="right")

        cols = ("token","desc","ex")
        self.tv = ttk.Treeview(self, columns=cols, show="headings")
        self.tv.heading("token", text="代碼")
        self.tv.heading("desc", text="說明")
        self.tv.heading("ex", text="範例")
        self.tv.column("token", width=220, anchor="w")
        self.tv.column("desc", width=340, anchor="w")
        self.tv.column("ex", width=220, anchor="w")
        self.tv.pack(fill="both", expand=True, padx=8, pady=(0,8))

        btns = ttk.Frame(self, padding=8); btns.pack(fill="x")
        ttk.Button(btns, text="插入", command=self.insert_selected).pack(side="left")
        ttk.Button(btns, text="關閉", command=self.destroy).pack(side="right")

        self.tv.bind("<Double-1>", lambda e: self.insert_selected())
        self.var_q.trace_add("write", lambda *_: self.refresh())
        self.refresh()

    def refresh(self):
        q = self.var_q.get().strip().lower()
        self.tv.delete(*self.tv.get_children())
        for t, d, ex in CODE_DEFS:
            if not q or (q in t.lower() or q in d.lower()):
                self.tv.insert("", "end", values=(t, d, ex))

    def insert_selected(self):
        sel = self.tv.selection()
        if not sel: return
        token = self.tv.item(sel[0], "values")[0]
        self.insert_cb(token)

    def show_help(self):
        win = tk.Toplevel(self)
        win.title("客製化重新命名 — 使用說明")
        win.geometry("820x560")
        txt = tk.Text(win, wrap="word")
        txt.insert("1.0", CUSTOM_HELP)
        txt.config(state="disabled")
        txt.pack(fill="both", expand=True)

# ---------- 重複檔清理對話框 ----------
DUP_HELP = (
    "比對原理：\n"
    "1) 先以檔案大小分組，不同大小不可能相同。\n"
    "2) 同一大小再計算內容 MD5 雜湊，MD5 相同視為重複。\n"
    "3) 依『保留策略』決定每組保留哪一份，其餘列為刪除候選。\n\n"
    "注意：請先『預覽』確認，再執行刪除。預設送回收桶，可還原。"
)

class DuplicateCleanerDialog(tk.Toplevel):
    def __init__(self, master, root_cfg, log_cb=None):
        super().__init__(master)
        self.title("重複檔清理")
        self.geometry("1100x680")
        self.resizable(True, True)
        self.root_cfg = root_cfg
        self.log_cb = log_cb or (lambda s: None)
        pad = 8

        top = ttk.Frame(self, padding=pad); top.pack(fill="x")
        ttk.Button(top, text="比對原理", command=lambda: messagebox.showinfo("說明", DUP_HELP)).pack(side="left")
        ttk.Label(top, text="保留策略：").pack(side="left", padx=(12,4))
        self.keep_map = {
            "保留最新": "newest",
            "保留最舊": "oldest",
            "優先保留含關鍵字": "keyword",
            "優先保留路徑前綴": "prefix",
        }
        self.combo_keep = ttk.Combobox(top, state="readonly", width=18, values=list(self.keep_map.keys()))
        self.combo_keep.set("保留最新"); self.combo_keep.pack(side="left")
        ttk.Label(top, text="關鍵字：").pack(side="left", padx=(12,4))
        self.entry_keyword = ttk.Entry(top, width=22); self.entry_keyword.pack(side="left")
        ttk.Label(top, text="路徑前綴：").pack(side="left", padx=(12,4))
        self.entry_prefix = ttk.Entry(top, width=30); self.entry_prefix.pack(side="left")

        mid = ttk.Frame(self, padding=(pad,0,pad,pad)); mid.pack(fill="x")
        ttk.Button(mid, text="掃描重複", command=self.scan_dups).pack(side="left")
        ttk.Button(mid, text="全選候選", command=self.select_all_delete).pack(side="left", padx=6)
        ttk.Button(mid, text="全不選", command=self.unselect_all).pack(side="left")
        ttk.Button(mid, text="送回收桶刪除", command=lambda: self.delete_selected(mode="recycle")).pack(side="right")
        ttk.Button(mid, text="永久刪除", command=lambda: self.delete_selected(mode="permanent")).pack(side="right", padx=6)

        cols = ("keep","group","size","mtime","path")
        self.tv = ttk.Treeview(self, columns=cols, show="headings", selectmode="extended")
        for c, w, a, t in [
            ("keep", 80, "center", "保留?"),
            ("group", 160, "w", "群組(MD5)"),
            ("size", 100, "e", "大小"),
            ("mtime", 160, "center", "修改時間"),
            ("path", 560, "w", "路徑"),
        ]:
            self.tv.heading(c, text=t); self.tv.column(c, width=w, anchor=a)
        self.tv.pack(fill="both", expand=True, padx=8, pady=(0,8))

        bottom = ttk.Frame(self, padding=pad); bottom.pack(fill="x")
        self.pb = ttk.Progressbar(bottom, orient="horizontal", mode="determinate")
        self.pb.pack(fill="x", expand=True, side="left", padx=(0,8))
        self.lbl = ttk.Label(bottom, text="待命"); self.lbl.pack(side="left")

        self.items = []  # list of dict: {path,size,mtime,md5,group,keep}
        self.group_map = {}  # md5 -> list indexes

        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="此列設為保留", command=self.context_keep)
        self.menu.add_command(label="此組只保留此列", command=self.context_keep_only)
        self.tv.bind("<Button-3>", self.popup)

    def popup(self, e):
        iid = self.tv.identify_row(e.y)
        if iid:
            self.tv.selection_set(iid)
            self.menu.post(e.x_root, e.y_root)

    def context_keep(self):
        for iid in self.tv.selection():
            idx = int(self.tv.set(iid, "keep"))
            self.items[idx]["keep"] = True
        self._refresh_rows()

    def context_keep_only(self):
        for iid in self.tv.selection():
            idx = int(self.tv.set(iid, "keep"))
            md5 = self.items[idx]["group"]
            for j in self.group_map.get(md5, []):
                self.items[j]["keep"] = False
            self.items[idx]["keep"] = True
        self._refresh_rows()

    def _fmt_size(self, n):
        for unit in ("B","KB","MB","GB","TB"):
            if n < 1024 or unit=="TB": return f"{n:.1f}{unit}"
            n /= 1024

    def _refresh_rows(self):
        self.tv.delete(*self.tv.get_children())
        for i, it in enumerate(self.items):
            self.tv.insert("", "end", values=(i, it["group"], self._fmt_size(it["size"]),
                                              time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(it["mtime"])),
                                              it["path"]),
                           tags=("keep" if it["keep"] else "del",))
        self.tv.tag_configure("keep", background="#f0fff0")
        self.tv.tag_configure("del", background="#fff0f0")

    def _update_status(self, percent, msg):
        self.pb.config(value=percent)
        self.lbl.config(text=msg)
        self.update_idletasks()

    def _collect_files(self):
        roots = list(self.root_cfg.keys())
        include_map = {r: self.root_cfg[r]['include_root'] for r in roots}
        exts_map = {r: self.root_cfg[r]['exts_selected'] for r in roots}
        files = []
        for r in roots:
            arr = gather_all_files(r, include_root=include_map[r])
            if exts_map[r] != "ALL":
                arr = [p for p in arr if (ext_key(p) in exts_map[r] or ("_noext" in exts_map[r] and ext_key(p)=="_noext"))]
            files.extend(arr)
        return files

    def _group_by_size(self, paths):
        size_map = {}
        for p in paths:
            try:
                st = os.stat(p)
            except OSError:
                continue
            size_map.setdefault(st.st_size, []).append((p, st.st_mtime))
        return {sz: lst for sz, lst in size_map.items() if len(lst) > 1}

    def _hash_groups(self, size_groups):
        items = []
        self.group_map.clear()
        sz_keys = list(size_groups.keys())
        total = len(sz_keys) if sz_keys else 1
        for i, sz in enumerate(sz_keys, 1):
            lst = size_groups[sz]
            hash_map = {}
            for p, m in lst:
                try:
                    h = full_md5(p)
                except Exception:
                    continue
                hash_map.setdefault(h, []).append((p, m))
            for h, glist in hash_map.items():
                if len(glist) < 2:
                    continue
                for p, m in glist:
                    items.append({"path": p, "size": sz, "mtime": m, "group": h, "keep": False})
            self._update_status(int(i/total*50), f"計算雜湊 {i}/{total}")
        return items

    def _apply_keep_strategy(self):
        strategy = self.keep_map[self.combo_keep.get()]
        kw = self.entry_keyword.get().strip().lower()
        px = self.entry_prefix.get().strip().lower()
        self.group_map.clear()
        for idx, it in enumerate(self.items):
            self.group_map.setdefault(it["group"], []).append(idx)
        for md5, idxs in self.group_map.items():
            pick = None
            if strategy == "prefix" and px:
                for j in idxs:
                    if self.items[j]["path"].lower().startswith(px):
                        pick = j; break
            if pick is None and strategy == "keyword" and kw:
                for j in idxs:
                    if kw in os.path.basename(self.items[j]["path"]).lower():
                        pick = j; break
            if pick is None:
                if strategy == "newest":
                    pick = max(idxs, key=lambda j: self.items[j]["mtime"])
                elif strategy == "oldest":
                    pick = min(idxs, key=lambda j: self.items[j]["mtime"])
                else:
                    pick = max(idxs, key=lambda j: self.items[j]["mtime"])
            for j in idxs:
                self.items[j]["keep"] = (j == pick)

    def scan_dups(self):
        def worker():
            self.items = []
            self._update_status(0, "蒐集檔案…")
            files = self._collect_files()
            if not files:
                self._update_status(0, "沒有可掃描的檔案"); return
            self._update_status(5, "依大小分組…")
            size_groups = self._group_by_size(files)
            if not size_groups:
                self._update_status(0, "沒有偵測到重複大小的檔案"); return
            self._update_status(10, "計算 MD5…")
            self.items = self._hash_groups(size_groups)
            if not self.items:
                self._update_status(0, "沒有偵測到雜湊相同的檔案"); return
            self._update_status(60, "套用保留策略…")
            self._apply_keep_strategy()
            self._update_status(80, "建立清單…")
            self._refresh_rows()
            del_cnt = sum(1 for it in self.items if not it["keep"])
            self._update_status(100, f"完成。重複組：{len(self.group_map)}，刪除候選：{del_cnt}")
        threading.Thread(target=worker, daemon=True).start()

    def select_all_delete(self):
        for it in self.items:
            it["keep"] = False
        self._refresh_rows()

    def unselect_all(self):
        for it in self.items:
            it["keep"] = True
        self._refresh_rows()

    def delete_selected(self, mode="recycle"):
        targets = [it["path"] for it in self.items if not it["keep"]]
        if not targets:
            messagebox.showinfo("提示", "沒有勾選要刪除的項目。"); return
        if any(os.path.splitext(p)[1].lower() in PROTECTED_EXTS for p in targets):
            if not messagebox.askyesno("警告", "清單包含可執行或系統副檔名。仍要繼續？"):
                return
        if mode == "permanent":
            if not messagebox.askyesno("永久刪除確認", f"確定永久刪除 {len(targets)} 個檔案？此操作不可還原。"):
                return
        def worker():
            ok = 0; fail = 0
            total = len(targets)
            for i, p in enumerate(targets, 1):
                try:
                    if mode == "recycle":
                        if move_to_recycle_bin_one(p): ok += 1
                        else: fail += 1
                    else:
                        st = os.stat(p)
                        os.chmod(p, stat.S_IWRITE | stat.S_IREAD)
                        os.remove(p) if not stat.S_ISDIR(st.st_mode) else os.rmdir(p)
                        ok += 1
                except Exception:
                    fail += 1
                self.pb.config(value=int(i/total*100))
                self.lbl.config(text=f"刪除中 {i}/{total}")
            self.pb.config(value=100)
            self.lbl.config(text=f"刪除完成：成功 {ok}，失敗 {fail}")
            messagebox.showinfo("刪除結果", f"成功 {ok}，失敗 {fail}")
            remain = []
            for it in self.items:
                if not os.path.exists(it["path"]):
                    continue
                remain.append(it)
            self.items = remain
            self.group_map.clear()
            for idx, it in enumerate(self.items):
                self.group_map.setdefault(it["group"], []).append(idx)
            self._refresh_rows()
        threading.Thread(target=worker, daemon=True).start()

# ---------- 一般規則刪除對話框（含空資料夾清理） ----------
class DeleteToolDialog(tk.Toplevel):
    def __init__(self, master, root_cfg, log_cb=None):
        super().__init__(master)
        self.title("一般刪除工具")
        self.geometry("1100x720")
        self.resizable(True, True)
        self.root_cfg = root_cfg
        self.log_cb = log_cb or (lambda s: None)
        pad = 8

        nb = ttk.Notebook(self); nb.pack(fill="both", expand=True)

        # Tab1: 規則刪除
        tab_rule = ttk.Frame(nb); nb.add(tab_rule, text="規則刪除")
        frm = ttk.Frame(tab_rule, padding=pad); frm.pack(fill="x")
        left = ttk.LabelFrame(tab_rule, text="條件", padding=pad)
        left.pack(fill="x", padx=8, pady=(0,8))
        self.var_min = tk.StringVar(value="0")
        self.var_max = tk.StringVar(value="0")
        self.var_days = tk.StringVar(value="0")
        self.var_inc_kw = tk.StringVar(value="")
        self.var_exc_kw = tk.StringVar(value="")
        self.var_allow_prot = tk.BooleanVar(value=False)
        ttk.Label(left, text="最小大小(MB, 0=不設限)：").grid(row=0, column=0, sticky="w")
        ttk.Entry(left, textvariable=self.var_min, width=8).grid(row=0, column=1, sticky="w", padx=6)
        ttk.Label(left, text="最大大小(MB, 0=不設限)：").grid(row=0, column=2, sticky="w")
        ttk.Entry(left, textvariable=self.var_max, width=8).grid(row=0, column=3, sticky="w", padx=6)
        ttk.Label(left, text="早於 N 天未修改(0=不設限)：").grid(row=0, column=4, sticky="w")
        ttk.Entry(left, textvariable=self.var_days, width=8).grid(row=0, column=5, sticky="w", padx=6)
        ttk.Label(left, text="名稱包含(, 分隔)：").grid(row=1, column=0, sticky="w", pady=(6,0))
        ttk.Entry(left, textvariable=self.var_inc_kw, width=36).grid(row=1, column=1, columnspan=2, sticky="w", padx=6, pady=(6,0))
        ttk.Label(left, text="名稱排除(, 分隔)：").grid(row=1, column=3, sticky="w", pady=(6,0))
        ttk.Entry(left, textvariable=self.var_exc_kw, width=36).grid(row=1, column=4, columnspan=2, sticky="w", padx=6, pady=(6,0))
        ttk.Checkbutton(left, text="允許包含受保護副檔名(.exe/.dll…)", variable=self.var_allow_prot).grid(row=2, column=0, columnspan=6, sticky="w", pady=(6,0))

        btns = ttk.Frame(tab_rule, padding=(8,0,8,0)); btns.pack(fill="x")
        ttk.Button(btns, text="預覽候選", command=self.preview_rule).pack(side="left")
        ttk.Button(btns, text="全選候選", command=lambda: self._set_all_rule(False)).pack(side="left", padx=6)
        ttk.Button(btns, text="全不選", command=lambda: self._set_all_rule(True)).pack(side="left")
        ttk.Button(btns, text="送回收桶刪除", command=lambda: self.exec_rule("recycle")).pack(side="right")
        ttk.Button(btns, text="永久刪除", command=lambda: self.exec_rule("permanent")).pack(side="right", padx=6)

        cols = ("keep","size","mtime","path","reason")
        self.tv_rule = ttk.Treeview(tab_rule, columns=cols, show="headings", selectmode="extended")
        for c, w, a, t in [
            ("keep", 70, "center", "刪除?"),
            ("size", 120, "e", "大小"),
            ("mtime", 160, "center", "修改時間"),
            ("path", 600, "w", "路徑"),
            ("reason", 120, "w", "原因"),
        ]:
            self.tv_rule.heading(c, text=t); self.tv_rule.column(c, width=w, anchor=a)
        self.tv_rule.pack(fill="both", expand=True, padx=8, pady=(4,8))

        # Tab2: 空資料夾
        tab_empty = ttk.Frame(nb); nb.add(tab_empty, text="空資料夾清理")
        ttk.Label(tab_empty, text="清理所有根目錄內的空資料夾（會自動刪常見殘留檔：thumbs.db、desktop.ini…）").pack(anchor="w", padx=8, pady=8)
        ttk.Button(tab_empty, text="開始清理", command=self.clean_empty_dirs).pack(anchor="w", padx=8)
        self.pb_empty = ttk.Progressbar(tab_empty, orient="horizontal", mode="determinate"); self.pb_empty.pack(fill="x", padx=8, pady=(8,8))
        self.lbl_empty = ttk.Label(tab_empty, text="待命"); self.lbl_empty.pack(anchor="w", padx=8, pady=(0,8))

        bottom = ttk.Frame(self, padding=pad); bottom.pack(fill="x")
        self.pb = ttk.Progressbar(bottom, orient="horizontal", mode="determinate")
        self.pb.pack(fill="x", expand=True, side="left", padx=(0,8))
        self.lbl = ttk.Label(bottom, text="待命"); self.lbl.pack(side="left")

        self.rule_items = []  # list of dict: {path,size,mtime,reason,delete(bool)}

    def _fmt_size(self, n):
        for unit in ("B","KB","MB","GB","TB"):
            if n < 1024 or unit=="TB": return f"{n:.1f}{unit}"
            n /= 1024

    def _collect_files_by_cfg(self):
        roots = list(self.root_cfg.keys())
        include_map = {r: self.root_cfg[r]['include_root'] for r in roots}
        exts_map = {r: self.root_cfg[r]['exts_selected'] for r in roots}
        files = []
        for r in roots:
            arr = gather_all_files(r, include_root=include_map[r])
            if exts_map[r] != "ALL":
                arr = [p for p in arr if (ext_key(p) in exts_map[r] or ("_noext" in exts_map[r] and ext_key(p)=="_noext"))]
            files.extend(arr)
        return files

    def preview_rule(self):
        try:
            min_mb = float(self.var_min.get() or 0)
            max_mb = float(self.var_max.get() or 0)
            days = float(self.var_days.get() or 0)
        except ValueError:
            messagebox.showwarning("格式錯誤", "大小或天數需為數字"); return
        inc = [x.strip() for x in self.var_inc_kw.get().split(",") if x.strip()]
        exc = [x.strip() for x in self.var_exc_kw.get().split(",") if x.strip()]
        allow_prot = self.var_allow_prot.get()

        def worker():
            self.rule_items = []
            files = self._collect_files_by_cfg()
            total = len(files) if files else 1
            now = time.time()
            kept = 0
            self.pb.config(value=0)
            self.lbl.config(text="依規則掃描…")
            for i, p in enumerate(files, 1):
                try:
                    st = os.stat(p)
                except OSError:
                    continue
                name = os.path.basename(p)
                ext = os.path.splitext(name)[1].lower()
                if not allow_prot and ext in PROTECTED_EXTS:
                    pass
                else:
                    sz_mb = st.st_size / (1024*1024)
                    old_ok = (days==0 or (now - st.st_mtime) >= days*86400)
                    min_ok = (min_mb==0 or sz_mb >= min_mb)
                    max_ok = (max_mb==0 or sz_mb <= max_mb)
                    inc_ok = (not inc or any(k.lower() in name.lower() for k in inc))
                    exc_ok = (not exc or not any(k.lower() in name.lower() for k in exc))
                    if old_ok and min_ok and max_ok and inc_ok and exc_ok:
                        self.rule_items.append({
                            "path": p, "size": st.st_size, "mtime": st.st_mtime,
                            "reason": "rule", "delete": True
                        })
                        kept += 1
                if i % 50 == 0 or i == total:
                    self.pb.config(value=int(i/total*100))
                    self.lbl.config(text=f"掃描中 {i}/{total}  已選 {kept}")
                    self.update_idletasks()
            self._render_rule_tv()
            self.lbl.config(text=f"完成。候選 {len(self.rule_items)}")
        threading.Thread(target=worker, daemon=True).start()

    def _render_rule_tv(self):
        self.tv_rule.delete(*self.tv_rule.get_children())
        for idx, it in enumerate(self.rule_items):
            self.tv_rule.insert("", "end",
                                values=("是" if it["delete"] else "否",
                                        self._fmt_size(it["size"]),
                                        time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(it["mtime"])),
                                        it["path"], it["reason"]),
                                tags=("del" if it["delete"] else "keep",))
        self.tv_rule.tag_configure("del", background="#fff0f0")
        self.tv_rule.tag_configure("keep", background="#f0fff0")

    def _set_all_rule(self, keep_all):
        for it in self.rule_items:
            it["delete"] = (not keep_all)
        self._render_rule_tv()

    def exec_rule(self, mode):
        targets = [it["path"] for it in self.rule_items if it["delete"]]
        if not targets:
            messagebox.showinfo("提示", "沒有要刪除的項目。"); return
        if any(os.path.splitext(p)[1].lower() in PROTECTED_EXTS for p in targets):
            if not messagebox.askyesno("警告", "清單包含可執行或系統副檔名。仍要繼續？"):
                return
        if mode == "permanent":
            if not messagebox.askyesno("永久刪除確認", f"確定永久刪除 {len(targets)} 個檔案？此操作不可還原。"):
                return
        def worker():
            ok, fail = 0, 0
            total = len(targets)
            for i, p in enumerate(targets, 1):
                try:
                    if mode == "recycle":
                        if move_to_recycle_bin_one(p): ok += 1
                        else: fail += 1
                    else:
                        st = os.stat(p)
                        os.chmod(p, stat.S_IWRITE | stat.S_IREAD)
                        os.remove(p) if not stat.S_ISDIR(st.st_mode) else os.rmdir(p)
                        ok += 1
                except Exception:
                    fail += 1
                self.pb.config(value=int(i/total*100))
                self.lbl.config(text=f"刪除中 {i}/{total}")
                self.update_idletasks()
            self.pb.config(value=100); self.lbl.config(text=f"完成：成功 {ok}，失敗 {fail}")
            messagebox.showinfo("刪除結果", f"成功 {ok}，失敗 {fail}")
            remain = []
            for it in self.rule_items:
                if not os.path.exists(it["path"]):
                    continue
                remain.append(it)
            self.rule_items = remain
            self._render_rule_tv()
        threading.Thread(target=worker, daemon=True).start()

    def clean_empty_dirs(self):
        roots = list(self.root_cfg.keys())
        if not roots:
            messagebox.showinfo("提示","請先加入根目錄。"); return
        def worker():
            total_removed = 0
            for r in roots:
                self.lbl_empty.config(text=f"清理中：{r}")
                total_removed += prune_empty_dirs(r, dry_run=False,
                                                  progress_cb=lambda cur, tot, msg: (self.pb_empty.config(value=(cur%100)), self.lbl_empty.config(text=msg)))
            self.pb_empty.config(value=100)
            self.lbl_empty.config(text=f"完成。移除空資料夾 {total_removed} 個")
            messagebox.showinfo("空資料夾清理", f"總共移除 {total_removed} 個空資料夾")
        threading.Thread(target=worker, daemon=True).start()

# ========= 自動更新（GitHub Raw + Release） =========
# ========= 自動更新（GitHub Raw + Release） =========
def _ver_tuple(v: str):
    return tuple(int(x) for x in re.findall(r"\d+", v)[:3] or [0])

def _http_get(url, timeout=HTTP_TIMEOUT):
    req = urllib.request.Request(url, headers={"User-Agent": f"{APP_NAME}/{APP_VERSION}"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()

def _download_to_tmp(url, suffix):
    data = _http_get(url)
    fd, p = tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    with open(p, "wb") as f:
        f.write(data)
    return p

def _write_bat_and_run(args):
    """
    args: [NEW_PATH, TARGET_EXE_OR_PY, PYEXE_OR_EMPTY, PYSRC_OR_EMPTY]
    會：1) 等舊程式退出→2) 重試刪除→3) 改名/覆蓋→4) 依參數重啟
    全程把訊息寫到 UPDATE_LOG（%TEMP%/FileOrganizer_update.log）
    """
    fd, bat = tempfile.mkstemp(suffix=".bat")
    os.close(fd)

    # 用 ^ 續行，所有關鍵步驟寫入 log；DEBUG_UPDATE 時最後加 PAUSE
    script = rf"""@echo off
setlocal ENABLEDELAYEDEXPANSION
chcp 65001 >nul
set "LOG={UPDATE_LOG}"
echo === [{APP_NAME}] Self-Update started %date% %time% ===> "%LOG%"
echo BAT path: "%~f0" >> "%LOG%"

set "NEW=%~1"
set "TARGET=%~2"
set "PYEXE=%~3"
set "PYSRC=%~4"

echo NEW="%NEW%" >> "%LOG%"
echo TARGET="%TARGET%" >> "%LOG%"
echo PYEXE="%PYEXE%" >> "%LOG%"
echo PYSRC="%PYSRC%" >> "%LOG%"

rem 等待舊檔解鎖
set RETRIES=30
:WAIT_UNLOCK
  2>nul (>>"%TARGET%" echo test) && (echo target free >> "%LOG%") && goto DO_REPLACE
  echo target locked, retry... >> "%LOG%"
  timeout /t 1 /nobreak >nul
  set /a RETRIES-=1
  if %RETRIES% GTR 0 goto WAIT_UNLOCK

echo ERROR: target still locked after retries >> "%LOG%"
goto END

:DO_REPLACE
echo Try delete old target >> "%LOG%"
del /f /q "%TARGET%" >> "%LOG%" 2>&1
if exist "%TARGET%" (
  echo delete failed, will overwrite by move/copy >> "%LOG%"
)

echo Move new to target >> "%LOG%"
move /Y "%NEW%" "%TARGET%" >> "%LOG%" 2>&1
if errorlevel 1 (
  echo move failed, try copy /Y then del >> "%LOG%"
  copy /Y "%NEW%" "%TARGET%" >> "%LOG%" 2>&1
  if errorlevel 1 (
    echo ERROR: copy failed too >> "%LOG%"
    goto END
  )
)

echo Restart phase >> "%LOG%"
if "%PYEXE%"=="" (
  start "" "%TARGET%" >> "%LOG%" 2>&1
) else (
  start "" "%PYEXE%" "%PYSRC%" >> "%LOG%" 2>&1
)

:END
echo === Self-Update done %date% %time% ===>> "%LOG%"
"""

    if DEBUG_UPDATE:
        script += "\r\npause\r\n"

    with open(bat, "w", encoding="utf-8") as f:
        f.write(script)

    try:
        # 用 cmd /c 執行 bat；DEBUG_UPDATE=True 時會留下黑窗
        subprocess.Popen(["cmd.exe", "/c", bat] + args, creationflags=0)
    except Exception:
        os.system('start "" cmd /c "{}" {}'.format(bat, " ".join(f'"{a}"' for a in args)))




def check_for_updates(silent=False, parent=None):
    # 判斷是否為打包執行（Nuitka/pyinstaller 皆可）
    IS_FROZEN = bool(getattr(sys, "frozen", False) or "__compiled__" in globals())

    # 讀取 manifest.json（支援 exe_url / url / py_url）
    try:
        info_raw = _http_get(MANIFEST_URL).decode("utf-8", "replace")
        info     = json.loads(info_raw)
        latest   = str(info.get("version", "0.0.0"))
        notes    = str(info.get("notes", "") or "")
        exe_url  = info.get("exe_url") or info.get("url", "")  # 兼容 url 欄位
        py_url   = info.get("py_url", "")
    except Exception as e:
        if not silent:
            messagebox.showwarning("更新", f"取得更新資訊失敗：{e}")
        return

    # 已是最新
    if _ver_tuple(latest) <= _ver_tuple(APP_VERSION):
        if not silent:
            messagebox.showinfo("更新", f"目前已是最新版本（{APP_VERSION}）。")
        return

    # 詢問是否更新
    msg = (
        f"偵測到新版本 {latest}（目前 {APP_VERSION}）。\n\n"
        f"更新內容：\n{notes or '—'}\n\n要立即更新嗎？"
    )
    if not messagebox.askyesno("有可用更新", msg, parent=parent):
        return

    # 下載與交棒給更新腳本
    try:
        if IS_FROZEN:
            if not exe_url:
                messagebox.showwarning("更新", "manifest.json 未提供 exe_url/url。")
                return
            new_path = _download_to_tmp(exe_url, ".exe")
            target   = os.path.abspath(sys.executable)  # 目前執行中的 .exe
            # 直接覆蓋並重啟 exe（第三、四參數留空）
            _write_bat_and_run([new_path, target, "", ""])
        else:
            if not py_url:
                messagebox.showwarning("更新", "manifest.json 未提供 py_url（原始碼更新網址）。")
                return
            new_path = _download_to_tmp(py_url, ".py")
            target   = os.path.abspath(__file__)
            # 優先用 pythonw 重啟，沒有則回退 python
            candidate = sys.executable.replace("python.exe", "pythonw.exe")
            pythonw   = candidate if os.path.exists(candidate) else sys.executable
            _write_bat_and_run([new_path, target, pythonw, target])

        # 立刻退出，讓外部批次/PS 腳本接手覆蓋與重啟
        sys.exit(0)

    except urllib.error.URLError as e:
        messagebox.showerror("更新", f"下載失敗：{e}")
    except Exception as e:
        messagebox.showerror("更新", f"更新失敗：{e}")


# ---------- 主程式 ----------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1280x860"); self.minsize(1220, 820)

        # === 功能表 ===
        m = tk.Menu(self)
        self.config(menu=m)
        m_help = tk.Menu(m, tearoff=0)
        m.add_cascade(label="說明", menu=m_help)
        m_help.add_command(label="檢查更新…", command=lambda: check_for_updates(silent=True, parent=self))
        m_help.add_command(label="關於", command=lambda: messagebox.showinfo("關於", f"{APP_NAME}\n版本：{APP_VERSION}"))

        top = ttk.Frame(self, padding=8); top.pack(fill="x")
        ttk.Label(top, text="貼上根目錄路徑（可多行；以 ; 或換行分隔）：").grid(row=0, column=0, sticky="w")
        self.txt = tk.Text(top, height=4); self.txt.grid(row=1, column=0, sticky="ew", pady=(4,8))
        top.grid_columnconfigure(0, weight=1)

        row2 = ttk.Frame(top); row2.grid(row=2, column=0, sticky="ew")
        ttk.Button(row2, text="加入到清單", command=self.add_paths).pack(side="left", padx=5)
        ttk.Button(row2, text="瀏覽新增…", command=self.browse_path).pack(side="left", padx=5)

        mid = ttk.Frame(self, padding=(8,0,8,8)); mid.pack(fill="both", expand=True)
        cols = ("rootname","types","include_root")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("rootname", text="根目錄")
        self.tree.heading("types", text="已選類型")
        self.tree.heading("include_root", text="包含根目錄檔案")
        self.tree.column("rootname", width=380, anchor="w")
        self.tree.column("types", width=720, anchor="w")
        self.tree.column("include_root", width=160, anchor="center")
        self.tree.pack(fill="both", expand=True)

        self.tree_ids = {}; self.path_by_iid = {}
        TreeTooltip(self.tree, lambda iid: self.path_by_iid.get(iid, ""))

        ops = ttk.Frame(mid); ops.pack(fill="x")
        ttk.Button(ops, text="設定所選根目錄…", command=self.config_selected).pack(side="left")
        ttk.Button(ops, text="設定所有根目錄…", command=self.config_all).pack(side="left", padx=6)
        ttk.Button(ops, text="移除所選", command=self.remove_selected).pack(side="left", padx=6)
        ttk.Button(ops, text="清空清單", command=self.clear_list).pack(side="left", padx=6)
        ttk.Button(ops, text="一般刪除…", command=self.open_delete_tool).pack(side="right", padx=6)
        ttk.Button(ops, text="重複檔清理…", command=self.open_dup_cleaner).pack(side="right")

        opts = ttk.LabelFrame(self, text="重新命名選項", padding=8); opts.pack(fill="x", padx=8, pady=(4,8))
        ttk.Label(opts, text="重複檔名處理：").grid(row=0, column=0, sticky="w")
        self.dup_map = {"跳過同名":"skip","覆蓋同名":"overwrite","自動加序號":"index","加時間後綴":"datetime","加MD5前8碼":"hash8"}
        self.combo_dup = ttk.Combobox(opts, state="readonly", width=18, values=list(self.dup_map.keys()))
        self.combo_dup.set("自動加序號"); self.combo_dup.grid(row=0, column=1, sticky="w", padx=6)

        ttk.Label(opts, text="客製化重新命名：").grid(row=0, column=2, sticky="e")
        self.custom_map = {"無":"none","加前綴":"prefix","加後綴":"suffix","整個檔名改為客製":"replace"}
        self.combo_custom = ttk.Combobox(opts, state="readonly", width=18, values=list(self.custom_map.keys()))
        self.combo_custom.set("無"); self.combo_custom.grid(row=0, column=3, sticky="w", padx=6)

        ttk.Label(opts, text="內容：").grid(row=0, column=4, sticky="e")
        self.entry_custom = ttk.Entry(opts, width=44); self.entry_custom.grid(row=0, column=5, sticky="w", padx=6)

        codes_combo_values = [t for t,_,_ in CODE_DEFS]
        ttk.Label(opts, text="插入代碼：").grid(row=0, column=6, sticky="e")
        self.combo_codes = ttk.Combobox(opts, state="readonly", width=28, values=codes_combo_values)
        self.combo_codes.set("{檔名不含副檔名}")
        ttk.Button(opts, text="插入", command=lambda: self.entry_custom.insert("insert", self.combo_codes.get())).grid(row=0, column=7, sticky="w", padx=6)
        ttk.Button(opts, text="代碼清單…", command=self.open_code_palette).grid(row=0, column=8, sticky="w")
        ttk.Button(opts, text="使用說明…", command=self.show_custom_help).grid(row=0, column=9, sticky="w", padx=6)

        self.var_change_ext = tk.BooleanVar(value=False)
        ttk.Checkbutton(opts, text="變更副檔名", variable=self.var_change_ext,
                        command=lambda: self.entry_new_ext.config(state=("normal" if self.var_change_ext.get() else "disabled"))).grid(row=1, column=0, sticky="w", pady=(8,0))
        ttk.Label(opts, text="新副檔名：").grid(row=1, column=1, sticky="e", pady=(8,0))
        self.entry_new_ext = ttk.Entry(opts, width=16, state="disabled")
        self.entry_new_ext.grid(row=1, column=2, sticky="w", padx=6, pady=(8,0))
        ttk.Label(opts, text="例：jpg 或 .jpg").grid(row=1, column=3, sticky="w", pady=(8,0))

        logbar = ttk.Frame(self, padding=(8,0,8,0)); logbar.pack(fill="x")
        self.var_verbose = tk.BooleanVar(value=False)
        ttk.Checkbutton(logbar, text="顯示詳細日誌", variable=self.var_verbose, command=self.toggle_log).pack(side="left")
        self.log_frame = ttk.Frame(self, padding=(8,4,8,8))
        self.log_text = tk.Text(self.log_frame, height=10); self.log_text.pack(fill="both", expand=True)

        bottom = ttk.Frame(self, padding=8); bottom.pack(fill="x")
        self.btn_preview = ttk.Button(bottom, text="預覽前 10 筆改名結果", command=self.preview_rename); self.btn_preview.pack(side="left")
        self.btn_run = ttk.Button(bottom, text="開始執行", command=self.run_all); self.btn_run.pack(side="left", padx=8)
        ttk.Button(bottom, text="離開", command=self.safe_exit).pack(side="right")
        self.pb = ttk.Progressbar(bottom, orient="horizontal", mode="determinate"); self.pb.pack(fill="x", expand=True, padx=8)
        self.lbl_status = ttk.Label(bottom, text="待命"); self.lbl_status.pack(fill="x")

        self.running = False
        self.root_cfg = {}
        self._start_ts = None

        # 啟動後自動檢查更新（避免卡 UI，延遲少許）
        if AUTO_CHECK_ON_START:
            self.after(1200, lambda: check_for_updates(silent=True, parent=self))

    # 開啟工具
    def open_code_palette(self):
        CodePalette(self, insert_cb=lambda token: self.entry_custom.insert("insert", token))
    def open_dup_cleaner(self):
        if not self.root_cfg:
            messagebox.showinfo("提示","請先加入根目錄。"); return
        dlg = DuplicateCleanerDialog(self, self.root_cfg, log_cb=self.log)
        dlg.grab_set(); self.wait_window(dlg)
    def open_delete_tool(self):
        if not self.root_cfg:
            messagebox.showinfo("提示","請先加入根目錄。"); return
        dlg = DeleteToolDialog(self, self.root_cfg, log_cb=self.log)
        dlg.grab_set(); self.wait_window(dlg)

    # 使用說明視窗
    def show_custom_help(self):
        win = tk.Toplevel(self)
        win.title("客製化重新命名 — 使用說明")
        win.geometry("820x560")
        win.resizable(True, True)
        txt = tk.Text(win, wrap="word")
        txt.insert("1.0", CUSTOM_HELP)
        txt.config(state="disabled")
        txt.pack(fill="both", expand=True)

    # 日誌：捲動鎖定
    def _log_at_bottom(self):
        try:
            _, bottom = self.log_text.yview()
            return bottom >= 0.999
        except Exception:
            return True
    def log(self, msg):
        if not self.var_verbose.get(): return
        def _a():
            autoscroll = self._log_at_bottom()
            self.log_text.insert("end", msg+"\n")
            if autoscroll: self.log_text.see("end")
        self.after(0, _a)
    def toggle_log(self):
        if self.var_verbose.get(): self.log_frame.pack(fill="both", expand=True)
        else: self.log_frame.forget()

    # 路徑處理
    def parse_pasted(self, s):
        raw = s.replace("\r","\n").replace(";", "\n").split("\n")
        return [p.strip().strip('"') for p in raw if p.strip()]

    def add_paths(self):
        paths = self.parse_pasted(self.txt.get("1.0","end")); self.txt.delete("1.0","end")
        for p in paths:
            if not os.path.isdir(p): continue
            if p in self.root_cfg: continue
            exts = scan_exts(p)
            self.root_cfg[p] = {'exts_all': exts, 'exts_selected': "ALL", 'include_root': False, 'flatten_all': False}
            iid = self.tree.insert("", "end",
                                   values=(os.path.basename(p) or p, self._types_label(p), "否"))
            self.tree_ids[p] = iid; self.path_by_iid[iid] = p
        self._refresh_tree_labels()

    def browse_path(self):
        p = filedialog.askdirectory(title="選取根目錄")
        if not p or not os.path.isdir(p) or p in self.root_cfg: return
        exts = scan_exts(p)
        self.root_cfg[p] = {'exts_all': exts, 'exts_selected': "ALL", 'include_root': False, 'flatten_all': False}
        iid = self.tree.insert("", "end",
                               values=(os.path.basename(p) or p, self._types_label(p), "否"))
        self.tree_ids[p] = iid; self.path_by_iid[iid] = p
        self._refresh_tree_labels()

    def _types_label(self, p):
        cfg = self.root_cfg[p]
        if cfg['exts_selected']=="ALL": return "全部"
        if not cfg['exts_selected']: return "（未選）"
        sample = sorted(cfg['exts_selected']); s = ", ".join(sample[:12])
        if len(sample)>12: s += f" …共{len(sample)}種"
        return s

    def _refresh_tree_labels(self):
        for p, iid in self.tree_ids.items():
            inc = "是" if self.root_cfg[p]['include_root'] else "否"
            self.tree.item(iid, values=(os.path.basename(p) or p, self._types_label(p), inc))

    def _selected_root(self):
        sel = self.tree.selection()
        if not sel: return None
        return self.path_by_iid.get(sel[0])

    def config_selected(self):
        p = self._selected_root()
        if not p:
            messagebox.showinfo("提示","請先選取一個根目錄。"); return
        dlg = RootConfigDialog(self, p, self.root_cfg[p].copy())
        self.wait_window(dlg)
        self.root_cfg[p].update(dlg.cfg)
        self._refresh_tree_labels()

    def config_all(self):
        roots = list(self.root_cfg.keys())
        if not roots:
            messagebox.showinfo("提示","清單為空。"); return
        dlg = BulkConfigDialog(self, roots, self.root_cfg)
        self.wait_window(dlg)
        self._refresh_tree_labels()

    def remove_selected(self):
        p = self._selected_root()
        if not p: return
        iid = self.tree_ids[p]
        self.tree.delete(iid); del self.tree_ids[p]; del self.root_cfg[p]; del self.path_by_iid[iid]

    def clear_list(self):
        self.tree.delete(*self.tree.get_children())
        self.root_cfg.clear(); self.tree_ids = {}; self.path_by_iid = {}

    # 預覽前 10 筆改名結果
    def preview_rename(self):
        p = self._selected_root()
        if not p:
            messagebox.showinfo("提示","請先在清單中選取一個根目錄。"); return
        if not os.path.isdir(p):
            messagebox.showwarning("提示","選取的項目不是資料夾。"); return
        dup_mode = self.dup_map[self.combo_dup.get()]
        custom_mode = self.custom_map[self.combo_custom.get()]
        custom_text = self.entry_custom.get().strip()
        change_ext = self.var_change_ext.get()
        new_ext = self.entry_new_ext.get().strip()
        cfg = self.root_cfg[p]
        files = gather_all_files(p, include_root=cfg['include_root'])
        if cfg['exts_selected'] != "ALL":
            files = [x for x in files if (ext_key(x) in cfg['exts_selected'] or ("_noext" in cfg['exts_selected'] and ext_key(x)=="_noext"))]
        files = files[:10]
        win = tk.Toplevel(self); win.title(f"預覽改名（前 10 筆） - {os.path.basename(p) or p}"); win.geometry("1000x440")
        cols = ("src","dest"); tv = ttk.Treeview(win, columns=cols, show="headings")
        tv.heading("src", text="來源檔案"); tv.heading("dest", text="預計目的檔名（含分類資料夾）")
        tv.column("src", width=480, anchor="w"); tv.column("dest", width=480, anchor="w"); tv.pack(fill="both", expand=True)
        root_name = os.path.basename(p) or p
        for src in files:
            e = ext_key(src)
            dest_dir = p if cfg.get('flatten_all', False) else os.path.join(p, e if e != "_noext" else "_noext")
            decision = decide_dest(src, dest_dir, dup_mode, custom_mode, custom_text, change_ext, new_ext, root_name=root_name)
            dest = "(跳過：同名已存在)" if decision is None else (decision[1] if isinstance(decision, tuple) else decision)
            tv.insert("", "end", values=(src, dest))
        ttk.Label(win, text="提示：依目前設定推演；實際結果仍可能受現場同名檔案影響。").pack(anchor="w", padx=8, pady=6)
        ttk.Button(win, text="關閉", command=win.destroy).pack(pady=4)

    # ====== 小工具：格式化時長 ======
    @staticmethod
    def _fmt_dur(sec):
        sec = int(max(0, sec))
        h, r = divmod(sec, 3600)
        m, s = divmod(r, 60)
        if h: return f"{h}h{m:02d}m{s:02d}s"
        if m: return f"{m}m{s:02d}s"
        return f"{s}s"

    # 執行整理
    def run_all(self):
        if self.running: return
        roots = list(self.root_cfg.keys())
        if not roots:
            messagebox.showwarning("提示","清單為空。請先加入根目錄。"); return
        dup_mode = self.dup_map[self.combo_dup.get()]
        custom_mode = self.custom_map[self.combo_custom.get()]
        custom_text = self.entry_custom.get().strip()
        change_ext = self.var_change_ext.get()
        new_ext = self.entry_new_ext.get().strip()
        if custom_mode!="none" and not custom_text:
            if not messagebox.askyesno("確認","未填客製內容，確定以空字串執行？"): return
        if change_ext and not new_ext:
            if not messagebox.askyesno("確認","已勾選變更副檔名，但未填新副檔名。要改為沿用原始副檔名嗎？"):
                return

        self.running = True; self.btn_run.config(state="disabled"); self.btn_preview.config(state="disabled")
        self.pb.config(value=0, maximum=100); self.lbl_status.config(text="準備中…")
        dry_run = False
        self._start_ts = time.time()

        def worker():
            totals = {}
            for r in roots:
                cfg = self.root_cfg[r]
                files = gather_all_files(r, include_root=cfg['include_root'])
                if cfg['exts_selected']!="ALL":
                    files = [p for p in files if (ext_key(p) in cfg['exts_selected'] or ("_noext" in cfg['exts_selected'] and ext_key(p)=="_noext"))]
                totals[r] = len(files)
            total_all = sum(totals.values()) or 1
            root_done = {r: 0 for r in roots}

            def progress_cb(root, done, total, msg):
                done_sum = sum((done if rr==root else min(root_done.get(rr,0), totals[rr])) for rr in totals)
                percent = 100.0 * (done_sum/total_all)
                now = time.time()
                elapsed = now - self._start_ts if self._start_ts else 0.0
                eta_txt = "—"; remain_txt = "—"
                if elapsed >= 2 and done_sum > 0:
                    rate = done_sum / elapsed
                    remain = max(0.0, total_all - done_sum) / max(rate, 1e-9)
                    eta = time.localtime(now + remain)
                    eta_txt = time.strftime("%H:%M:%S", eta)
                    remain_txt = self._fmt_dur(remain)
                elapsed_txt = self._fmt_dur(elapsed)
                def ui():
                    self.pb.config(value=percent)
                    self.lbl_status.config(
                        text=f"[{os.path.basename(root) or root}] {msg}｜整體 {percent:5.1f}%｜ETA {eta_txt}｜剩餘 {remain_txt}｜已耗時 {elapsed_txt}"
                    )
                    root_done[root] = done
                self.after(0, ui)

            for r in roots:
                if not os.path.isdir(r):
                    self.log(f"[SKIP] 非資料夾：{r}"); continue
                cfg = self.root_cfg[r]
                try:
                    organize(
                        r,
                        cfg['include_root'],
                        cfg['exts_selected'],
                        dup_mode,
                        custom_mode,
                        custom_text,
                        change_ext,
                        new_ext,
                        dry_run,
                        progress_cb,
                        self.log,
                        flatten_all=cfg.get('flatten_all', False)
                    )
                except Exception as e:
                    self.log(f"[ERROR] {r} → {e}")

            def finish():
                total_elapsed = self._fmt_dur(time.time() - self._start_ts) if self._start_ts else "0s"
                self.pb.config(value=100)
                self.lbl_status.config(text=f"全部完成｜總耗時 {total_elapsed}")
                self.btn_run.config(state="normal"); self.btn_preview.config(state="normal")
                self.running = False
                messagebox.showinfo("完成", f"全部處理完成。\n總耗時：{total_elapsed}")
            self.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    def safe_exit(self):
        if self.running and not messagebox.askyesno("離開","工作仍在進行，確定要離開？"): return
        self.destroy()

if __name__ == "__main__":
    ensure_admin()
    app = App()
    app.mainloop()
