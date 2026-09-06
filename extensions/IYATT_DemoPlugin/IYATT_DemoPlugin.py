'''
file: IYATT_DemoPlugin
description: 示例插件：跨表数据处理与批量格式化
author: IYATT-yx
copyright:  Copyright (c) 2026 IYATT-yx.
            Licensed under the MIT License. See LICENSE file in the project root for full license information.
'''
from pytableenginesdk.PluginLogger import PluginLogger

pluginInfo = {
    'name': 'DemoPlugin',
    'author': 'IYATT',
    'description': '示例插件：跨表数据处理与批量格式化',
    'version': '0.0.1',
}


def run(appComHandle, logQueue):
    '''
    插件主执行入口

    Args:
        appComHandle: 主进程传入的 COM 句柄（Excel.Application 或 WPS 实例）
        logQueue: 进程间通信日志队列
    '''
    logger = PluginLogger(logQueue, pluginInfo)
    logger.info('demoPlugin 开始执行跨表数据处理...')

    # 检查是否有打开的工作簿
    if appComHandle.Workbooks.Count == 0:
        logger.warning('当前没有打开的 Excel/WPS 工作簿，自动新建一个演示工作簿。')
        wb = appComHandle.Workbooks.Add()
    else:
        wb = appComHandle.ActiveWorkbook

    # 演示跨表操作（若只有一个 Sheet，则自动新建第二个 Sheet）
    if wb.Sheets.Count < 2:
        wsSource = wb.Sheets(1)
        wsTarget = wb.Sheets.Add(After=wsSource)
        wsSource.Name = '数据源'
        wsTarget.Name = '处理结果'
    else:
        wsSource = wb.Sheets(1)
        wsTarget = wb.Sheets(2)

    # 性能优化设置：冻结屏幕刷新
    appComHandle.ScreenUpdating = False
    try:
        # 在数据源 Sheet 制造一些示例数据
        sampleData = [
            ['项目', '数值 A', '数值 B', '计算结果'],
            ['零件 1', 100, 200, None],
            ['零件 2', 150, 250, None],
            ['零件 3', 300, 400, None],
        ]
        wsSource.Range('A1:D4').Value = sampleData

        # 一次性读取数据到 Python 内存 (整块读取)
        rawData = wsSource.Range('A1:D4').Value

        # 在 Python 内存中进行数据计算
        processedData = [list(row) for row in rawData]
        for idx in range(1, len(processedData)):
            valA = processedData[idx][1] or 0
            valB = processedData[idx][2] or 0
            processedData[idx][3] = valA + valB

        # 跨表写入：把处理好的结果一次性灌入目标 Sheet
        wsTarget.Range('A1:D4').Value = processedData

        # 批量调整格式（1 次 COM 调用）
        headerRange = wsTarget.Range('A1:D1')
        headerRange.Interior.Color = 0xFFD700  # 金黄色背景
        headerRange.Font.Bold = True

        logger.info(f'成功将 [{wsSource.Name}] 处理好的数据跨表写入到 [{wsTarget.Name}]')

    finally:
        # 恢复屏幕刷新
        appComHandle.ScreenUpdating = True