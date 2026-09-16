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
from typing import Any, Callable, Iterable, List, Optional, Tuple, Union

user32 = ctypes.windll.user32
GetAsyncKeyState = user32.GetAsyncKeyState
GetAsyncKeyState.argtypes = [ctypes.c_int]
GetAsyncKeyState.restype = wintypes.SHORT

VK_RETURN = 0x0D  # 回车键


class SelectionPickerUI:
    """通用交互区域选取器面板"""

    def __init__(
            self, 
            appComHandle: Any, 
            logger: Optional[Any] = None, 
            title: str = "区域选取器", 
            promptTip: Optional[str] = None, 
            singleMode: bool = False, 
            validator: Optional[Callable[[List[Tuple[str, str, Any]]], Union[bool, str]]]|None = None
        ) -> None:
        """
        初始化区域选取器 UI。
        
        Args:
            appComHandle (Any): Excel/WPS 的 COM 句柄。
            logger (Optional[Any]): 日志记录器实例，默认为 None。
            title (str): 弹窗的标题文本，默认为 "区域选取器"。
            promptTip (Optional[str]): 自定义界面的快捷操作提示，如果不传则使用默认提示。
            singleMode (bool): 是否为单选模式（开启后新提取的区域会覆盖旧区域），默认为 False。
            validator (Optional[Callable]): 自定义校验函数。
                接收参数: targets (List[Tuple[str, str, Any]])
                返回: True 表示校验通过；返回 str 表示校验失败并作为警告信息弹出。
        """
        self._app = appComHandle
        self._logger = logger
        self._title = title
        self._singleMode = singleMode
        self._validator = validator

        self._origMoveAfterReturn = True
        try:
            self._origMoveAfterReturn = self._app.MoveAfterReturn
        except Exception:
            pass
        
        if promptTip:
            self._promptTip = promptTip
        elif self._singleMode:
            self._promptTip = "快捷操作提示：\n • 在 Excel 选好目标后，按【Enter 回车】提取单元格（新选择将覆盖旧选择）\n • 提取完成后，点击确定提交\n • 按【Esc】取消退出"
        else:
            self._promptTip = "快捷操作提示：\n • 在 Excel 选好区域后，按【Enter 回车】快速提取选区\n • 提取完成后，点击确定提交\n • 按【Esc】取消退出"

        self._selectedTargets = []
        self._isConfirmed = False
        self._isListening = False
        self._lastEnterState = False
        self._afterId = None

        self._root = tk.Tk()
        self._root.title(self._title)
        self._root.geometry("460x370")
        self._root.attributes("-topmost", True)
        self._root.resizable(False, False)

        atexit.register(self._stopListening)

        self._buildUi()
        self._bindEvents()

    def _buildUi(self):
        """构建 GUI 界面布局"""
        lblTip = ttk.Label(
            self._root,
            text=self._promptTip,
            justify="left",
            wraplength=440,
            font=("Microsoft YaHei", 9)
        )
        lblTip.pack(padx=10, pady=8, anchor="w")

        frameList = ttk.Frame(self._root)
        frameList.pack(padx=10, pady=2, fill="both", expand=True)

        self._listbox = tk.Listbox(
            frameList,
            selectmode=tk.SINGLE if self._singleMode else tk.EXTENDED,
            height=9,
            font=("Consolas", 10)
        )
        self._listbox.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(frameList, orient="vertical", command=self._listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self._listbox.config(yscrollcommand=scrollbar.set)

        frameListBtns = ttk.Frame(self._root)
        frameListBtns.pack(padx=10, pady=6, fill="x")

        btnAdd = ttk.Button(frameListBtns, text="⚡ 提取当前选区 (Enter)", command=self._addCurrentSelection)
        btnAdd.pack(side="left", padx=2)

        btnDel = ttk.Button(frameListBtns, text="🗑️ 移除选中 (Del)", command=self._deleteSelectedItems)
        btnDel.pack(side="left", padx=2)

        btnClear = ttk.Button(frameListBtns, text="🧹 清空", command=self._clearAllItems)
        btnClear.pack(side="right", padx=2)

        frameBottom = ttk.Frame(self._root)
        frameBottom.pack(padx=10, pady=8, fill="x")

        btnConfirm = ttk.Button(frameBottom, text="🚀 确定提交", command=self._onConfirm)
        btnConfirm.pack(side="right", padx=5)

        btnCancel = ttk.Button(frameBottom, text="取消 (Esc)", command=self._onCancel)
        btnCancel.pack(side="right", padx=5)

    def _bindEvents(self):
        """绑定快捷键与事件"""
        self._root.bind("<Escape>", lambda e: self._onCancel())
        self._listbox.bind("<Delete>", lambda e: self._deleteSelectedItems())

    def _startListening(self):
        """启动安全轮询检测，并临时禁用 Enter 下移"""
        self._isListening = True
        try:
            self._app.MoveAfterReturn = False  # 禁用 Enter 跳转
        except Exception:
            pass
        self._pollEnterKey()
        
    def _stopListening(self):
        """停止轮询检测并恢复 Excel 默认属性"""
        self._isListening = False
        try:
            self._app.MoveAfterReturn = self._origMoveAfterReturn  # 恢复 Enter 跳转
        except Exception:
            pass

        if self._afterId and self._root:
            try:
                self._root.after_cancel(self._afterId)
            except Exception:
                pass
            self._afterId = None

    def _pollEnterKey(self):
        """主线程安全：利用 GetAsyncKeyState 进行精准按键检测"""
        if not self._isListening:
            return

        try:
            # 增加窗口存活校验
            if not self._root or not self._root.winfo_exists():
                self._isListening = False
                return

            state = GetAsyncKeyState(VK_RETURN)
            isPressed = bool(state & 0x8000)

            # 边沿触发检测：按下 Enter 瞬间触发一次提取
            if isPressed and not self._lastEnterState:
                self._addCurrentSelection()

            self._lastEnterState = isPressed

        except Exception as e:
            # 捕获 Tkinter 实例已销毁引发的 TclError 异常
            if isinstance(e, tk.TclError):
                self._isListening = False
                return
            if self._logger:
                self._logger.error(f"按键检测发生异常: {e}")

        # 重新注册下一个周期的定时器并保存句柄
        if self._isListening and self._root:
            try:
                if self._root.winfo_exists():
                    self._afterId = self._root.after(30, self._pollEnterKey)
            except tk.TclError:
                self._isListening = False

    def _addCurrentSelection(self):
        """安全提取 Excel 当前 Selection 区域，并重置焦点到左上角单元格"""
        try:
            sel = self._app.Selection
            if not sel:
                return

            # 开启单选模式时强制只取左上角单个单元格
            if self._singleMode:
                sel = sel.Cells(1, 1)

            sheetName = sel.Worksheet.Name
            address = sel.Address
            displayStr = f"Sheet: {sheetName} | Range: {address}"

            # 开启单选模式时覆盖原有选择，多选模式时防重复插入
            if self._singleMode:
                self._selectedTargets.clear()
                self._listbox.delete(0, tk.END)
            else:
                for sName, aStr, _ in self._selectedTargets:
                    if sName == sheetName and aStr == address:
                        return

            # 1. 存入已选列表
            self._selectedTargets.append((sheetName, address, sel))
            self._listbox.insert(tk.END, displayStr)
            self._listbox.see(tk.END)

            # 2. 取消大面积框选，定位到左上角第1格
            sel.Cells(1, 1).Select()

        except Exception as e:
            if self._logger:
                self._logger.error(f"提取选区发生异常: {e}")

    def _deleteSelectedItems(self):
        """删除 Listbox 选中项"""
        selectedIndices = list(self._listbox.curselection())
        if not selectedIndices:
            return

        for index in reversed(selectedIndices):
            self._listbox.delete(index)
            self._selectedTargets.pop(index)

    def _clearAllItems(self):
        """清空暂存列表"""
        self._listbox.delete(0, tk.END)
        self._selectedTargets.clear()

    def _onConfirm(self):
        """确认提交"""
        # 1. 确保有数据
        if not self._selectedTargets:
            if not self._selectedTargets:
                messagebox.showwarning("提示", "请先在 Excel 中选择有效区域并提取！", parent=self._root)
                return

        # 2. 触发回调校验（如果传入了 validator）
        if self._validator:
            try:
                # 将当前选中的所有目标传递给回调函数
                result = self._validator(self._selectedTargets)
                
                # 如果返回 False，说明校验失败（validator 内部可能已经自己弹窗了），拦截关闭
                if result is False:
                    return
                # 如果返回的是字符串，说明校验失败，直接用这个字符串弹出警告，拦截关闭
                elif isinstance(result, str):
                    messagebox.showwarning("校验失败", result, parent=self._root)
                    return
            except Exception as e:
                if self._logger:
                    self._logger.error(f"执行提交校验回调时发生异常: {e}")
                messagebox.showerror("校验错误", f"执行校验时发生内部错误:\n{e}", parent=self._root)
                return

        # 3. 校验通过（或没有校验函数），正常确认并退出
        self._isConfirmed = True
        self._closeWindow()

    def _onCancel(self):
        """取消操作"""
        self._isConfirmed = False
        self._closeWindow()

    def _closeWindow(self):
        """正常关闭窗口并收尾"""
        self._stopListening()
        try:
            self._root.destroy()
        except Exception:
            pass

    def show(self) -> Tuple[bool, List[Tuple[str, str, Any]]]:
        """
        显示面板并阻塞等待用户操作结束。
        
        Returns:
            Tuple[bool, List[Tuple[str, str, Any]]]:
                - 第一个元素 (bool): isConfirmed，标识用户是否点击了“确定提交”（True）还是“取消/关闭”（False）。
                - 第二个元素 (List): targets，提取到的选区数据列表。
                  列表元素结构为元组 (工作表名称, 单元格地址, Range的COM对象实例)。
        """
        try:
            self._startListening()
            self._root.protocol("WM_DELETE_WINDOW", self._onCancel)
            self._root.mainloop()
        except Exception as e:
            if self._logger:
                self._logger.error(f"选取器运行中发生未捕获异常: {e}")
            raise e
        finally:
            self._stopListening()

        return self._isConfirmed, self._selectedTargets