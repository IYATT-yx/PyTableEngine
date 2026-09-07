'''
file: IYATT_DemoPlugin
description: 示例插件：跨表数据处理与批量格式化（含 NumPy DLL 深度测试）
author: IYATT-yx
copyright:   Copyright (c) 2026 IYATT-yx.
            Licensed under the MIT License. See LICENSE file in the project root for full license information.
'''
import numpy as np
from pytableenginesdk.PluginLogger import PluginLogger

pluginInfo = {
    'name': 'DemoPlugin',
    'author': 'IYATT',
    'description': '示例插件：跨表数据处理与批量格式化（含 NumPy DLL 深度测试）',
    'version': '0.0.2',
}

def testNumpyDll(logger: PluginLogger):
    '''深度测试 NumPy 底层 C 扩展及动态链接库 (DLL) 是否全部正常工作'''
    try:
        # 基础版本与 C 扩展导出检查
        version = np.__version__
        logger.info(f'[NumPy 测试 1/4] 版本号导入成功: {version}')

        # 基础 C 数组与内存连续性测试
        arr = np.array([1.0, 2.0, 3.0, 4.0], dtype=np.float64)
        meanVal = np.mean(arr)
        logger.info(f'[NumPy 测试 2/4] C 数组均值计算成功 (Mean={meanVal})')

        # BLAS/LAPACK 底层 C/Fortran 动态链接库 (DLL) 矩阵乘法测试
        # (通常 DLL 加载失败、缺少 OpenBLAS/MKL/MSVC 依赖时会在这一步崩溃)
        matrixA = np.array([[1.0, 2.0], [3.0, 4.0]])
        matrixB = np.array([[5.0, 6.0], [7.0, 8.0]])
        dotResult = np.dot(matrixA, matrixB)
        logger.info(
            f'[NumPy 测试 3/4] BLAS/LAPACK DLL 矩阵乘法测试通过! 结果:\n{dotResult}'
        )

        # C 动态库随机数引擎与高级数学函数 (FFT) 测试
        randData = np.random.rand(8)
        fftResult = np.fft.fft(randData)
        logger.info('[NumPy 测试 4/4] 随机数生成器与 C-FFT 动态库调用成功！')

        return True, dotResult[0, 0]
    except Exception as e:
        logger.error(f'NumPy DLL 库运行异常 (可能缺少 C 运行时或 DLL): {e}')
        raise e


def run(appComHandle, logQueue):
    '''
    插件主执行入口

    Args:
        appComHandle: 主进程传入的 COM 句柄（Excel.Application 或 WPS 实例）
        logQueue: 进程间通信日志队列
    '''
    logger = PluginLogger(logQueue, pluginInfo)
    logger.info('demoPlugin 开始执行跨表数据处理...')

    # 执行 NumPy DLL 完整功能测试
    _, testCalcValue = testNumpyDll(logger)

    # 检查是否有打开的工作簿
    if appComHandle.Workbooks.Count == 0:
        logger.warning(
            '当前没有打开的 Excel/WPS 工作簿，自动新建一个演示工作簿。'
        )
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
        # 在数据源 Sheet 制造一些示例数据（引入 NumPy 的计算结果）
        sampleData = [
            ['项目', '数值 A', '数值 B', '计算结果'],
            ['零件 1', 100, 200, None],
            ['零件 2', 150, 250, None],
            ['零件 3 (NumPy DLL验证项)', testCalcValue, 400, None],
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

        logger.info(
            f'成功将 [{wsSource.Name}] 处理好的数据跨表写入到 [{wsTarget.Name}]'
        )

    finally:
        # 恢复屏幕刷新
        appComHandle.ScreenUpdating = True