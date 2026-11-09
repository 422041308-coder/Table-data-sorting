# -*- coding: utf-8 -*-
"""Windows 原生"选择文件夹"对话框（ctypes，零依赖）。

用途：Web 界面里点「浏览…」时，由本地服务弹出系统文件夹选择框，
选完把绝对路径返回给页面。只在 Windows 上可用；取消返回 None。
"""
from __future__ import annotations

import ctypes
import sys
from ctypes import wintypes


class BROWSEINFOW(ctypes.Structure):
    _fields_ = [
        ("hwndOwner", wintypes.HWND),
        ("pidlRoot", ctypes.c_void_p),
        ("pszDisplayName", wintypes.LPWSTR),
        ("lpszTitle", wintypes.LPCWSTR),
        ("ulFlags", ctypes.c_uint),
        ("lpfn", ctypes.c_void_p),
        ("lParam", ctypes.c_void_p),
        ("iImage", ctypes.c_int),
    ]


BIF_RETURNONLYFSDIRS = 0x00000001
BIF_EDITBOX = 0x00000010
BIF_NEWDIALOGSTYLE = 0x00000040
BIF_SHAREABLE = 0x00008000


def pick_folder(title: str = "选择文件夹", initial: str | None = None) -> str | None:
    """弹出系统文件夹选择框，返回绝对路径；用户取消返回 None。"""
    if not sys.platform.startswith("win"):
        raise RuntimeError("当前系统不支持原生文件夹选择框")
    ole32 = ctypes.windll.ole32
    shell32 = ctypes.windll.shell32

    ole32.CoInitialize(None)
    try:
        display = ctypes.create_unicode_buffer(1024)
        if initial:
            display.value = str(initial)[:1000]
        bi = BROWSEINFOW()
        bi.hwndOwner = None
        bi.pidlRoot = None
        bi.pszDisplayName = ctypes.cast(display, wintypes.LPWSTR)
        bi.lpszTitle = str(title or "选择文件夹")
        bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_EDITBOX | BIF_NEWDIALOGSTYLE | BIF_SHAREABLE

        pidl = shell32.SHBrowseForFolderW(ctypes.byref(bi))
        if not pidl:
            return None
        try:
            buf = ctypes.create_unicode_buffer(1024)
            if shell32.SHGetPathFromIDListW(pidl, buf):
                return buf.value or None
            return None
        finally:
            ole32.CoTaskMemFree(pidl)
    finally:
        ole32.CoUninitialize()
