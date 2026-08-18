# -*- coding: utf-8 -*-
"""平台启动器：以完全脱离当前会话的独立进程方式启动 Flask 服务
用法: python start_server.py
"""
import os
import subprocess
import sys
import time

BASE = os.path.dirname(os.path.abspath(__file__))
PY = r"E:\python\python.exe"
APP = os.path.join(BASE, "app.py")
LOG = os.path.join(BASE, "server.log")

DETACHED_PROCESS = 0x00000008
CREATE_NEW_PROCESS_GROUP = 0x00000200

# 若已有实例在跑，先停掉
import ctypes
try:
    import wmi  # 可选，失败就跳过旧进程清理
except ImportError:
    wmi = None

logf = open(LOG, "a", encoding="utf-8", buffering=1)
p = subprocess.Popen(
    [PY, APP],
    cwd=BASE,
    stdout=logf,
    stderr=subprocess.STDOUT,
    creationflags=DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP,
    close_fds=True,
)
print(f"[launcher] started pid={p.pid}")
time.sleep(3)
# 健康检查
import urllib.request
try:
    r = urllib.request.urlopen("http://127.0.0.1:5050/api/templates", timeout=8)
    print("[launcher] health check OK, HTTP", r.status)
except Exception as e:
    print("[launcher] health check FAILED:", e)
    print("--- server.log 尾部 ---")
    try:
        with open(LOG, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
        print("".join(lines[-15:]))
    except Exception:
        pass
