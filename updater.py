# -*- coding: utf-8 -*-
import os, sys, json, hashlib, tempfile, zipfile, shutil, urllib.request, time, subprocess

def _ver_tuple(v):
    # "1.2.10" -> (1,2,10)
    return tuple(int(x) for x in v.strip().split("."))

def _http_get(url, timeout=30):
    with urllib.request.urlopen(url, timeout=timeout) as r:
        return r.read()

def _sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024*1024), b""):
            h.update(chunk)
    return h.hexdigest()

def _extract_zip(zip_path, dst_dir):
    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(dst_dir)

def _is_frozen():
    return getattr(sys, "frozen", False)

def _app_dir():
    return os.path.dirname(sys.executable if _is_frozen() else os.path.abspath(sys.argv[0]))

def _restart(exe_or_py):
    if _is_frozen():
        # exe：直接啟動自己
        subprocess.Popen([exe_or_py], close_fds=True, shell=False)
    else:
        # .py：用同個 python 重新啟動
        subprocess.Popen([sys.executable, exe_or_py], close_fds=True, shell=False)

def _run_replacer_and_exit(src_dir, dst_dir, relaunch_target):
    """
    針對 exe（檔案鎖定）使用：開一個 bat 等本程式結束後再覆蓋，最後重啟。
    """
    pid = os.getpid()
    bat_path = os.path.join(tempfile.gettempdir(), f"update_{pid}.bat")
    # 以 xcopy 覆蓋全部檔案；等 PID 消失再動手；最後刪掉自己
    bat = rf"""@echo off
set SRC="{src_dir}"
set DST="{dst_dir}"
:wait
tasklist /FI "PID eq {pid}" | find "{pid}" >nul && (timeout /t 1 >nul & goto wait)
xcopy /E /I /Y "%SRC%\*" "%DST%" >nul
start "" "{relaunch_target}"
del "%~f0" & exit
"""
    with open(bat_path, "w", encoding="utf-8") as f:
        f.write(bat)
    # 啟動 bat 並退出
    subprocess.Popen([bat_path], shell=True, close_fds=True)
    os._exit(0)

def auto_update(app_name, current_version, manifest_url):
    """
    回傳 (changed: bool, message: str)
    """
    # 1) 抓取 manifest
    data = _http_get(manifest_url)
    mani = json.loads(data.decode("utf-8"))

    latest = mani["version"]
    if _ver_tuple(latest) <= _ver_tuple(current_version):
        return (False, f"{app_name} 已是最新版本（{current_version}）")

    asset_url = mani["url"]
    expected_sha256 = mani.get("sha256", "")
    notes = mani.get("notes", "")

    # 2) 下載壓縮包
    tmp = tempfile.mkdtemp(prefix="upd_")
    pkg_path = os.path.join(tmp, "pkg.zip")
    with open(pkg_path, "wb") as f:
        f.write(_http_get(asset_url))

    # 3) 校驗
    if expected_sha256:
        got = _sha256(pkg_path)
        if got.lower() != expected_sha256.lower():
            raise RuntimeError(f"SHA256 不符：{got} ≠ {expected_sha256}")

    # 4) 解壓到臨時資料夾
    new_dir = os.path.join(tmp, "new")
    os.makedirs(new_dir, exist_ok=True)
    _extract_zip(pkg_path, new_dir)

    # 5) 覆蓋安裝
    dst = _app_dir()

    if _is_frozen():
        # exe：用 bat 等程式關閉後再覆蓋，最後自動重啟
        target = sys.executable
        _run_replacer_and_exit(new_dir, dst, target)
        return (True, f"正在更新至 {latest}，程式將自動重啟…")
    else:
        # .py：可直接覆蓋
        for name in os.listdir(new_dir):
            s = os.path.join(new_dir, name)
            d = os.path.join(dst, name)
            if os.path.isdir(s):
                if os.path.exists(d): shutil.rmtree(d, ignore_errors=True)
                shutil.copytree(s, d)
            else:
                shutil.copy2(s, d)
        return (True, f"已更新至 {latest}！{notes or ''}".strip())
