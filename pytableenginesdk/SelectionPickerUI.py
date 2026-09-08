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
        self.mApp = appComHandle
        self.mLogger = logger
        self.mTitle = title
        self.mPromptTip = promptTip or "快捷操作提示：\n • 在 Excel 选好区域后，按【Enter 回车】快速提取选区\n • 提取完成后，点击确定提交\n • 按【Esc】取消退出"

        self.mSelectedTargets = []
        self.mIsConfirmed = False
        self.mIsListening = False
        self.mLastEnterState = False  # 用于防按键抖动

        # 初始化主窗口
        self.mRoot = tk.Tk()
        self.mRoot.title(self.mTitle)
        self.mRoot.geometry("460x370")
        self.mRoot.attributes("-topmost", True)
        self.mRoot.resizable(False, False)

        # 注册子进程退出兜底
        atexit.register(self.mStopListening)

        self.mBuildUi()
        self.mBindEvents()

    def mBuildUi(self):
        """构建 GUI 界面布局"""
        lblTip = ttk.Label(
            self.mRoot,
            text=self.mPromptTip,
            justify="left",
            wraplength=440,
            font=("Microsoft YaHei", 9)
        )
        lblTip.pack(padx=10, pady=8, anchor="w")

        frameList = ttk.Frame(self.mRoot)
        frameList.pack(padx=10, pady=2, fill="both", expand=True)

        self.mListbox = tk.Listbox(
            frameList,
            selectmode=tk.EXTENDED,
            height=9,
            font=("Consolas", 10)
        )
        self.mListbox.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(frameList, orient="vertical", command=self.mListbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.mListbox.config(yscrollcommand=scrollbar.set)

        frameListBtns = ttk.Frame(self.mRoot)
        frameListBtns.pack(padx=10, pady=6, fill="x")

        btnAdd = ttk.Button(frameListBtns, text="⚡ 提取当前选区 (Enter)", command=self.mAddCurrentSelection)
        btnAdd.pack(side="left", padx=2)

        btnDel = ttk.Button(frameListBtns, text="🗑️ 移除选中 (Del)", command=self.mDeleteSelectedItems)
        btnDel.pack(side="left", padx=2)

        btnClear = ttk.Button(frameListBtns, text="🧹 清空", command=self.mClearAllItems)
        btnClear.pack(side="right", padx=2)

        frameBottom = ttk.Frame(self.mRoot)
        frameBottom.pack(padx=10, pady=8, fill="x")

        btnConfirm = ttk.Button(frameBottom, text="🚀 确定提交", command=self.mOnConfirm)
        btnConfirm.pack(side="right", padx=5)

        btnCancel = ttk.Button(frameBottom, text="取消 (Esc)", command=self.mOnCancel)
        btnCancel.pack(side="right", padx=5)

    def mBindEvents(self):
        """绑定快捷键与事件"""
        self.mRoot.bind("<Escape>", lambda e: self.mOnCancel())
        self.mListbox.bind("<Delete>", lambda e: self.mDeleteSelectedItems())

    def mStartListening(self):
        """启动安全轮询检测"""
        self.mIsListening = True
        self.mPollEnterKey()

    def mStopListening(self):
        """停止轮询检测"""
        self.mIsListening = False

    def mPollEnterKey(self):
        """主线程安全：利用 GetAsyncKeyState 进行精准按键检测"""
        if not self.mIsListening or not self.mRoot:
            return

        try:
            state = GetAsyncKeyState(VK_RETURN)
            isPressed = bool(state & 0x8000)

            # 边沿触发检测：按下 Enter 瞬间触发一次提取
            if isPressed and not self.mLastEnterState:
                self.mAddCurrentSelection()

            self.mLastEnterState = isPressed

        except Exception as e:
            if self.mLogger:
                self.mLogger.error(f"按键检测发生异常: {e}")

        if self.mIsListening and self.mRoot:
            self.mRoot.after(30, self.mPollEnterKey)

    def mAddCurrentSelection(self):
        """安全提取 Excel 当前 Selection 区域，并重置焦点到左上角单元格"""
        try:
            sel = self.mApp.Selection
            if not sel:
                return

            sheetName = sel.Worksheet.Name
            address = sel.Address
            displayStr = f"Sheet: {sheetName} | Range: {address}"

            # 防重复插入列表
            for sName, aStr, _ in self.mSelectedTargets:
                if sName == sheetName and aStr == address:
                    return

            # 1. 存入已选列表
            self.mSelectedTargets.append((sheetName, address, sel))
            self.mListbox.insert(tk.END, displayStr)
            self.mListbox.see(tk.END)

            # 2. 核心：取消大面积框选，定位到左上角第1格，保住视野焦点不飘移
            sel.Cells(1, 1).Select()

        except Exception as e:
            if self.mLogger:
                self.mLogger.error(f"提取选区发生异常: {e}")

    def mDeleteSelectedItems(self):
        """删除 Listbox 选中项"""
        selectedIndices = list(self.mListbox.curselection())
        if not selectedIndices:
            return

        for index in reversed(selectedIndices):
            self.mListbox.delete(index)
            self.mSelectedTargets.pop(index)

    def mClearAllItems(self):
        """清空暂存列表"""
        self.mListbox.delete(0, tk.END)
        self.mSelectedTargets.clear()

    def mOnConfirm(self):
        """确认提交"""
        if not self.mSelectedTargets:
            self.mAddCurrentSelection()
            if not self.mSelectedTargets:
                messagebox.showwarning("提示", "请先在 Excel 中选择有效区域并提取！", parent=self.mRoot)
                return

        self.mIsConfirmed = True
        self.mCloseWindow()

    def mOnCancel(self):
        """取消操作"""
        self.mIsConfirmed = False
        self.mCloseWindow()

    def mCloseWindow(self):
        """正常关闭窗口并收尾"""
        self.mStopListening()
        try:
            self.mRoot.destroy()
        except Exception:
            pass

    def show(self):
        """显示面板"""
        try:
            self.mStartListening()
            self.mRoot.protocol("WM_DELETE_WINDOW", self.mOnCancel)
            self.mRoot.mainloop()
        except Exception as e:
            if self.mLogger:
                self.mLogger.error(f"选取器运行中发生未捕获异常: {e}")
            raise e
        finally:
            self.mStopListening()

        return self.mIsConfirmed, self.mSelectedTargets