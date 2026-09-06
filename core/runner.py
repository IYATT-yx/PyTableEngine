'''
file: runner.py
description: 插件执行管理
author: IYATT-yx
copyright:  Copyright (c) 2026 IYATT-yx.
            Licensed under the MIT License. See LICENSE file in the project root for full license information.
'''
import importlib.util
import os
import sys

import pythoncom
import win32com.client


def pluginRunnerTask(
    pluginDirPath, pluginFolderName, progId, logQueue, configFileName
):
    '''独立子进程：执行插件的具体逻辑'''
    pythoncom.CoInitialize()
    try:
        # 1. 设置依赖隔离
        vendorDir = os.path.join(pluginDirPath, 'vendor')
        if os.path.exists(vendorDir):
            sys.path.insert(0, vendorDir)
        sys.path.insert(0, pluginDirPath)

        # 2. 获取 COM 句柄
        try:
            comApp = win32com.client.GetActiveObject(progId)
        except Exception:
            logQueue.put(
                {
                    'level': 'INFO',
                    'message': f'未检测到运行中的 [{progId}]，尝试新建进程...',
                    'plugId': 'PluginRunner',
                    'filename': os.path.basename(__file__),
                    'funcName': sys._getframe().f_code.co_name,
                    'lineno': sys._getframe().f_lineno,
                }
            )
            comApp = win32com.client.Dispatch(progId)
            comApp.Visible = True

        # 3. 加载插件入口模组
        pydPath = os.path.join(pluginDirPath, f'{pluginFolderName}.pyd')
        pyPath = os.path.join(pluginDirPath, f'{pluginFolderName}.py')
        targetEntry = (
            pydPath
            if os.path.exists(pydPath)
            else (pyPath if os.path.exists(pyPath) else None)
        )

        if not targetEntry:
            raise FileNotFoundError(
                f'未在 [{pluginDirPath}] 目录下找到入口文件 {pluginFolderName}.py 或 .pyd'
            )

        spec = importlib.util.spec_from_file_location(
            f'ext_{pluginFolderName}', targetEntry
        )
        pluginModule = importlib.util.module_from_spec(spec)
        sys.modules[f'ext_{pluginFolderName}'] = pluginModule
        spec.loader.exec_module(pluginModule)

        # 4. 执行插件入口 run 函数
        if hasattr(pluginModule, 'run'):
            pluginModule.run(comApp, logQueue)
            logQueue.put(
                {
                    'type': 'SUCCESS',
                    'message': f'插件 [{pluginFolderName}] 执行完毕。',
                    'plugId': 'PluginRunner',
                    'filename': os.path.basename(__file__),
                    'funcName': sys._getframe().f_code.co_name,
                    'lineno': sys._getframe().f_lineno,
                }
            )
        else:
            raise AttributeError(
                f'插件 [{pluginFolderName}] 未定义 run(appComHandle, logQueue) 函数！'
            )

    except Exception:
        import traceback

        errMsg = traceback.format_exc()
        logQueue.put(
            {
                'type': 'ERROR',
                'message': f'插件执行异常:\n{errMsg}',
                'plugId': 'PluginRunner',
                'filename': os.path.basename(__file__),
                'funcName': sys._getframe().f_code.co_name,
                'lineno': sys._getframe().f_lineno,
            }
        )
    finally:
        pythoncom.CoUninitialize()