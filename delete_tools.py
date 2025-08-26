# -*- coding: utf-8 -*-
import os, stat, time, hashlib, ctypes, ctypes.wintypes as wt
from dataclasses import dataclass
from typing import Iterable, List, Callable, Dict, Tuple, Optional

# --------- 型別 ---------
@dataclass
class DeleteItem:
    path: str
    size: int
    mtime: float
    reason: str
    group_id: Optional[str] = None  # 重複組ID

ProgressCb = Callable[[int, int, str], None]   # current, total, msg
FilterCfg = Dict[str, object]                  # 規則設定

# --------- 工具 ---------
PROTECTED_EXTS = {'.exe','.dll','.sys','.msi','.lnk','.bat','.cmd','.ps1'}
IGNORABLE_FILES = {'thumbs.db','desktop.ini','.ds_store'}

def _is_protected(path: str) -> bool:
    name = os.path.basename(path).lower()
    if name in IGNORABLE_FILES:  # 這些可刪，不算保護
        return False
    _, ext = os.path.splitext(name)
    return ext.lower() in PROTECTED_EXTS

def _md5_16(path: str, chunk_mb: int = 4) -> str:
    h = hashlib.md5()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(chunk_mb*1024*1024), b''):
            h.update(chunk)
    return h.hexdigest()

# --------- 掃描候選 ---------
def collect_candidates(
    roots: Iterable[str],
    include_root_map: Dict[str, bool],
    exts_filter_map: Dict[str, object],
    rules: FilterCfg
) -> List[DeleteItem]:
    """
    rules 支援：
      mode: "filter" | "dupes"
      min_size_mb: float 或 0
      max_size_mb: float 或 0
      older_than_days: int 或 0
      name_include: List[str]
      name_exclude: List[str]
      allow_protected_exts: bool
    exts_filter_map[root] = "ALL" 或 set({'jpg','png', ...})
    """
    items: List[DeleteItem] = []
    now = time.time()

    for root in roots:
        if not os.path.isdir(root):
            continue
        for folder, _, files in os.walk(root):
            if not include_root_map.get(root, False) and os.path.normcase(folder) == os.path.normcase(root):
                # 不含根目錄直屬檔
                pass
            for name in files:
                p = os.path.join(folder, name)
                try:
                    st = os.stat(p)
                except OSError:
                    continue

                # 類型過濾
                extkey = (os.path.splitext(name)[1].lower() or "._noext").lstrip(".")
                flt = exts_filter_map.get(root, "ALL")
                if flt != "ALL":
                    if not (extkey in flt or ("_noext" in flt and extkey == "_noext")):
                        continue

                # 規則過濾
                if not rules.get("allow_protected_exts", False) and _is_protected(p):
                    continue
                sz_mb = st.st_size / (1024*1024)
                if rules.get("min_size_mb", 0) and sz_mb < float(rules["min_size_mb"]):
                    continue
                if rules.get("max_size_mb", 0) and sz_mb > float(rules["max_size_mb"]):
                    continue
                if rules.get("older_than_days", 0):
                    if now - st.st_mtime < float(rules["older_than_days"])*86400:
                        continue
                inc = rules.get("name_include", [])
                exc = rules.get("name_exclude", [])
                low = name.lower()
                if inc and not any(k.lower() in low for k in inc):
                    continue
                if exc and any(k.lower() in low for k in exc):
                    continue

                items.append(DeleteItem(path=p, size=st.st_size, mtime=st.st_mtime, reason="rule"))
    return items

# --------- 重複檔偵測 ---------
def group_duplicates(paths: Iterable[str], progress_cb: Optional[ProgressCb] = None) -> List[DeleteItem]:
    """
    先以檔案大小分桶，再算 MD5。回傳每組中除了第一個以外的候選（建議保留最早或最晚一份可由呼叫端決定）。
    """
    size_map: Dict[int, List[str]] = {}
    for p in paths:
        try:
            st = os.stat(p)
        except OSError:
            continue
        size_map.setdefault(st.st_size, []).append(p)

    cand: List[DeleteItem] = []
    total = sum(1 for v in size_map.values() if len(v) > 1)
    done = 0
    for sz, group in size_map.items():
        if len(group) < 2:
            continue
        # 計算 MD5
        hash_map: Dict[str, List[str]] = {}
        for p in group:
            try:
                h = _md5_16(p)
            except OSError:
                continue
            hash_map.setdefault(h, []).append(p)

        for h, g in hash_map.items():
            if len(g) < 2:
                continue
            # 以 mtime 最新者保留，其餘列為刪除候選
            g_sorted = sorted(g, key=lambda x: os.stat(x).st_mtime, reverse=True)
            keeper = g_sorted[0]
            for q in g_sorted[1:]:
                st = os.stat(q)
                cand.append(DeleteItem(path=q, size=st.st_size, mtime=st.st_mtime, reason=f"dup:{h}", group_id=h))
        done += 1
        if progress_cb:
            progress_cb(done, total, f"重複分組 {done}/{total}")
    return cand

# --------- 回收桶刪除（Windows） ---------
# 使用 SHFileOperationW，允許復原
FO_DELETE = 3
FOF_ALLOWUNDO = 0x0040
FOF_NOCONFIRMATION = 0x0010
FOF_SILENT = 0x0004
FOF_NOERRORUI = 0x0400

class SHFILEOPSTRUCTW(ctypes.Structure):
    _fields_ = [
        ("hwnd", wt.HWND),
        ("wFunc", wt.UINT),
        ("pFrom", wt.LPCWSTR),
        ("pTo", wt.LPCWSTR),
        ("fFlags", wt.UINT),
        ("fAnyOperationsAborted", wt.BOOL),
        ("hNameMappings", wt.LPVOID),
        ("lpszProgressTitle", wt.LPCWSTR),
    ]

def _paths_to_double_null(paths: List[str]) -> str:
    return "\x00".join(paths) + "\x00\x00"

def move_to_recycle_bin(paths: List[str], progress_cb: Optional[ProgressCb] = None) -> Tuple[int,int]:
    """
    送至回收桶；回傳 (成功數, 失敗數)
    """
    ok = 0; fail = 0
    total = len(paths)
    for i, chunk in enumerate(paths):
        # 一個一個送，避免路徑過長時整批失敗；也便於回報
        p = _paths_to_double_null([chunk])
        op = SHFILEOPSTRUCTW()
        op.hwnd = None
        op.wFunc = FO_DELETE
        op.pFrom = p
        op.pTo = None
        op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT
        rc = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
        if rc == 0 and not op.fAnyOperationsAborted:
            ok += 1
        else:
            fail += 1
        if progress_cb:
            progress_cb(i+1, total, f"回收桶刪除 {i+1}/{total}")
    return ok, fail

# --------- 永久刪除 ---------
def delete_permanently(paths: List[str], progress_cb: Optional[ProgressCb] = None) -> Tuple[int,int]:
    ok = 0; fail = 0
    total = len(paths)
    for i, p in enumerate(paths):
        try:
            st = os.stat(p)
            if stat.S_ISDIR(st.st_mode):
                # 呼叫端應傳入檔案；保守起見不遞迴刪資料夾
                os.rmdir(p)
            else:
                os.chmod(p, stat.S_IWRITE | stat.S_IREAD)
                os.remove(p)
            ok += 1
        except Exception:
            fail += 1
        if progress_cb:
            progress_cb(i+1, total, f"永久刪除 {i+1}/{total}")
    return ok, fail

# --------- 刪空資料夾 ---------
def prune_empty_dirs(root: str, progress_cb: Optional[ProgressCb] = None) -> int:
    removed = 0
    for folder, dirs, files in os.walk(root, topdown=False):
        try:
            if not os.listdir(folder):
                os.rmdir(folder)
                removed += 1
                if progress_cb:
                    progress_cb(removed, 0, f"移除空資料夾：{folder}")
        except Exception:
            pass
    return removed
