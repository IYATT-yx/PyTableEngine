'''
file: SelectionPickerUI.py
description: 通用 Excel 交互选取器组件
author: IYATT-yx
copyright:  Copyright (c) 2026 IYATT-yx.
            Licensed under the MIT License. See LICENSE file in the project root for full license information.
'''
import atexit
import ctypes
from ctypes import wintypes
import tkinter as tk
from tkinter import messagebox, ttk

user32 = ctypes.windll.user32
GetAsyncKeyState = user32.GetAsyncKeyState
GetAsyncKeyState.argtypes = [ctypes.c_int]
GetAsyncKeyState.restype = wintypes.SHORT

VK_RETURN = 0x0D  # 回车键


class SelectionPickerUI:
    """通用交互区域选取器面板"""

    def __init__(self, appComHandle, logger=None, title="区域选取器", promptTip=None):
        self._App = appComHandle
        self._Logger = logger
        self._Title = title
        self._PromptTip = promptTip or "快捷操作提示：\n • 在 Excel 选好区域后，按【Enter 回车】快速提取选区\n • 提取完成后，点击确定提交\n • 按【Esc】取消退出"

        self._SelectedTargets = []
        self._IsConfirmed = False
        self._IsListening = False
        self._LastEnterState = False  # 用于防按键抖动

        # 初始化主窗口
        self._Root = tk.Tk()
        self._Root.title(self._Title)
        self._Root.geometry("460x370")
        self._Root.attributes("-topmost", True)
        self._Root.resizable(False, False)

        # 注册子进程退出兜底
        atexit.register(self._StopListening)

        self._BuildUi()
        self._BindEvents()

    def _BuildUi(self):
        """构建 GUI 界面布局"""
        lblTip = ttk.Label(
            self._Root,
            text=self._PromptTip,
            justify="left",
            wraplength=440,
            font=("Microsoft YaHei", 9)
        )
        lblTip.pack(padx=10, pady=8, anchor="w")

        frameList = ttk.Frame(self._Root)
        frameList.pack(padx=10, pady=2, fill="both", expand=True)

        self._Listbox = tk.Listbox(
            frameList,
            selectmode=tk.EXTENDED,
            height=9,
            font=("Consolas", 10)
        )
        self._Listbox.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(frameList, orient="vertical", command=self._Listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self._Listbox.config(yscrollcommand=scrollbar.set)

        frameListBtns = ttk.Frame(self._Root)
        frameListBtns.pack(padx=10, pady=6, fill="x")

        btnAdd = ttk.Button(frameListBtns, text="⚡ 提取当前选区 (Enter)", command=self._AddCurrentSelection)
        btnAdd.pack(side="left", padx=2)

        btnDel = ttk.Button(frameListBtns, text="🗑️ 移除选中 (Del)", command=self._DeleteSelectedItems)
        btnDel.pack(side="left", padx=2)

        btnClear = ttk.Button(frameListBtns, text="🧹 清空", command=self._ClearAllItems)
        btnClear.pack(side="right", padx=2)

        frameBottom = ttk.Frame(self._Root)
        frameBottom.pack(padx=10, pady=8, fill="x")

        btnConfirm = ttk.Button(frameBottom, text="🚀 确定提交", command=self._OnConfirm)
        btnConfirm.pack(side="right", padx=5)

        btnCancel = ttk.Button(frameBottom, text="取消 (Esc)", command=self._OnCancel)
        btnCancel.pack(side="right", padx=5)

    def _BindEvents(self):
        """绑定快捷键与事件"""
        self._Root.bind("<Escape>", lambda e: self._OnCancel())
        self._Listbox.bind("<Delete>", lambda e: self._DeleteSelectedItems())

    def _StartListening(self):
        """启动安全轮询检测"""
        self._IsListening = True
        self._PollEnterKey()

    def _StopListening(self):
        """停止轮询检测"""
        self._IsListening = False

    def _PollEnterKey(self):
        """主线程安全：利用 GetAsyncKeyState 进行精准按键检测"""
        if not self._IsListening or not self._Root:
            return

        try:
            state = GetAsyncKeyState(VK_RETURN)
            isPressed = bool(state & 0x8000)

            # 边沿触发检测：按下 Enter 瞬间触发一次提取
            if isPressed and not self._LastEnterState:
                self._AddCurrentSelection()

            self._LastEnterState = isPressed

        except Exception as e:
            if self._Logger:
                self._Logger.error(f"按键检测发生异常: {e}")

        if self._IsListening and self._Root:
            self._Root.after(30, self._PollEnterKey)

    def _AddCurrentSelection(self):
        """安全提取 Excel 当前 Selection 区域，并重置焦点到左上角单元格"""
        try:
            sel = self._App.Selection
            if not sel:
                return

            sheetName = sel.Worksheet.Name
            address = sel.Address
            displayStr = f"Sheet: {sheetName} | Range: {address}"

            # 防重复插入列表
            for sName, aStr, _ in self._SelectedTargets:
                if sName == sheetName and aStr == address:
                    return

            # 1. 存入已选列表
            self._SelectedTargets.append((sheetName, address, sel))
            self._Listbox.insert(tk.END, displayStr)
            self._Listbox.see(tk.END)

            # 2. 核心：取消大面积框选，定位到左上角第1格，保住视野焦点不飘移
            sel.Cells(1, 1).Select()

        except Exception as e:
            if self._Logger:
                self._Logger.error(f"提取选区发生异常: {e}")

    def _DeleteSelectedItems(self):
        """删除 Listbox 选中项"""
        selectedIndices = list(self._Listbox.curselection())
        if not selectedIndices:
            return

        for index in reversed(selectedIndices):
            self._Listbox.delete(index)
            self._SelectedTargets.pop(index)

    def _ClearAllItems(self):
        """清空暂存列表"""
        self._Listbox.delete(0, tk.END)
        self._SelectedTargets.clear()

    def _OnConfirm(self):
        """确认提交"""
        if not self._SelectedTargets:
            self._AddCurrentSelection()
            if not self._SelectedTargets:
                messagebox.showwarning("提示", "请先在 Excel 中选择有效区域并提取！", parent=self._Root)
                return

        self._IsConfirmed = True
        self._CloseWindow()

    def _OnCancel(self):
        """取消操作"""
        self._IsConfirmed = False
        self._CloseWindow()

    def _CloseWindow(self):
        """正常关闭窗口并收尾"""
        self._StopListening()
        try:
            self._Root.destroy()
        except Exception:
            pass

    def show(self):
        """显示面板"""
        try:
            self._StartListening()
            self._Root.protocol("WM_DELETE_WINDOW", self._OnCancel)
            self._Root.mainloop()
        except Exception as e:
            if self._Logger:
                self._Logger.error(f"选取器运行中发生未捕获异常: {e}")
            raise e
        finally:
            self._StopListening()

        return self._IsConfirmed, self._SelectedTargets