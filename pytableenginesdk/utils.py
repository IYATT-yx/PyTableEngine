"""
file: utils.py
description: 基础工具集
author: IYATT-yx
copyright:  Copyright (c) 2026 IYATT-yx.
            Licensed under the MIT License. See LICENSE file in the project root for full license information.
"""
import contextlib
from typing import Callable, Iterable, List, Tuple, Any
import multiprocessing

from pytableenginesdk.PluginLogger import PluginLogger
from pytableenginesdk.SelectionPickerUI import SelectionPickerUI


@contextlib.contextmanager
def comEnvironment(appComHandle: Any, disableEvents: bool = False, manualCalc: bool = False):
    """
    Excel COM 运行环境上下文管理器。
    自动冻结屏幕刷新与警告弹出，并在退出时安全恢复原有设置。

    Args:
        appComHandle (Any): Excel/WPS 的 COM 应用句柄 (appComHandle)。
        disableEvents (bool): 是否禁用 Excel 事件监听（默认 False）。
        manualCalc (bool): 是否将重算模式切换为手动重算（默认 False）。
    """
    oldScreenUpdating = appComHandle.ScreenUpdating
    oldDisplayAlerts = appComHandle.DisplayAlerts
    oldEnableEvents = getattr(appComHandle, 'EnableEvents', True)
    oldCalculation = getattr(appComHandle, 'Calculation', -4105)  # xlCalculationAutomatic

    try:
        appComHandle.ScreenUpdating = False
        appComHandle.DisplayAlerts = False
        if disableEvents:
            appComHandle.EnableEvents = False
        if manualCalc:
            appComHandle.Calculation = -4135  # xlCalculationManual
        yield
    finally:
        if manualCalc:
            appComHandle.Calculation = oldCalculation
        if disableEvents:
            appComHandle.EnableEvents = oldEnableEvents
        appComHandle.DisplayAlerts = oldDisplayAlerts
        appComHandle.ScreenUpdating = oldScreenUpdating

def forEachArea(mainRng: Any) -> Iterable[Any]:
    """
    遍历 Range 对象的 Areas 子区域迭代器。
    统一兼容单区域与多重选择（Multi-selection）区域。

    Args:
        mainRng (Any): Excel COM Range 对象。

    Yields:
        Any: 拆分后的单个连续 Range 对象。
    """
    if mainRng is None:
        return

    areasCount = getattr(mainRng.Areas, 'Count', 1)
    if areasCount > 1:
        for idx in range(1, areasCount + 1):
            yield mainRng.Areas.Item(idx)
    else:
        yield mainRng

def batchCombineRanges(appComHandle: Any, ranges: List[Any], batchSize: int = 500) -> Iterable[Any]:
    """
    将一维 Range 列表分批次通过 app.Union 组合为连续的大块复合 Range。
    降低 COM 调用的频次，提升样式批量写入性能。

    Args:
        appComHandle (Any): Excel/WPS COM 句柄。
        ranges (List[Any]): 单个 Range 或单元格列表。
        batchSize (int): 每批组合的最大数量（默认 500）。

    Yields:
        Any: 通过 Union 组合后的复合 Range 对象。
    """
    if not ranges:
        return

    for i in range(0, len(ranges), batchSize):
        batch = ranges[i:i + batchSize]
        combinedRng = batch[0]
        for r in batch[1:]:
            combinedRng = appComHandle.Union(combinedRng, r)
        yield combinedRng

def runWithSelection(appComHandle: Any, logQueue: multiprocessing.Queue[str], pluginInfo: dict[str, str], pickerTitle: str, actionFunc: Callable[[Any, List[Tuple[str, str, Any]], PluginLogger], None], disableEvents: bool = False, manualCalc: bool = False ) -> None:
    """
    包含区域选取 UI 与 COM 环境保护的标准插件执行封装函数。

    Args:
        appComHandle (Any): Excel/WPS COM 句柄。
        logQueue (Any): 日志队列句柄。
        pluginInfo (dict): 插件元信息字典。
        pickerTitle (str): 选区弹窗标题。
        actionFunc (Callable): 核心业务回调函数，签名为 actionFunc(appComHandle, targets, logger)。
        disableEvents (bool): 是否在执行期间禁用事件（默认 False）。
        manualCalc (bool): 是否在执行期间开启手动重算（默认 False）。
    """
    logger = PluginLogger(logQueue, pluginInfo)
    logger.info(f"开始执行【{pluginInfo.get('name', '插件')}】...")

    if appComHandle.Workbooks.Count == 0:
        logger.warning('当前没有打开的工作簿。')
        return

    picker = SelectionPickerUI(appComHandle, logger, title=pickerTitle)
    confirmed, targets = picker.show()

    if not confirmed or not targets:
        logger.info('用户取消了选取操作，插件退出。')
        return

    try:
        with comEnvironment(appComHandle, disableEvents=disableEvents, manualCalc=manualCalc):
            actionFunc(appComHandle, targets, logger)
    except Exception as e:
        logger.error(f'插件处理过程出现异常: {e}')
        raise e