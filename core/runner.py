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
import multiprocessing
import inspect

import pythoncom
import win32com.client

from core import constants

def pluginRunnerTask(pluginDirPath: str, pluginFolderName: str, progId: str, logQueue: multiprocessing.Queue[dict[str, str | int]]):
    '''独立子进程：执行插件的具体逻辑'''
    pythoncom.CoInitialize()
    try:
        # 设置依赖隔离
        vendorDir = constants.Path.vendor
        if os.path.exists(vendorDir):
            # 优先加入 Python 模块搜索路径
            sys.path.insert(0, vendorDir)

            # --- 兼容 Windows 平台下 C 扩展及 DLL 动态库加载 ---
            # 注册 vendor 根目录
            if hasattr(os, 'add_dll_directory'):
                try:
                    os.add_dll_directory(vendorDir)
                except Exception:
                    pass

            # 自动检索并注册所有的 *.libs 目录（如 numpy.libs）
            for item in os.listdir(vendorDir):
                if item.endswith('.libs'):
                    libsDir = os.path.join(vendorDir, item)
                    if os.path.isdir(libsDir):
                        if hasattr(os, 'add_dll_directory'):
                            try:
                                os.add_dll_directory(libsDir)
                            except Exception:
                                pass
                        os.environ['PATH'] = (
                            libsDir + os.path.pathsep + os.environ.get('PATH', '')
                        )

            # 注册 bin 目录（若存在 DLL 或可执行依赖）
            binDir = os.path.join(vendorDir, 'bin')
            if os.path.exists(binDir) and os.path.isdir(binDir):
                if hasattr(os, 'add_dll_directory'):
                    try:
                        os.add_dll_directory(binDir)
                    except Exception:
                        pass
                os.environ['PATH'] = (
                    binDir + os.path.pathsep + os.environ.get('PATH', '')
                )

        sys.path.insert(0, pluginDirPath)

        # 获取 COM 句柄
        try:
            comApp = win32com.client.GetActiveObject(progId) # type: ignore
        except Exception:
            frame = inspect.currentframe()
            funcName = frame.f_code.co_name if frame else '<unknown>'
            lineno = frame.f_lineno if frame else -1
            logQueue.put(
                {
                    'level': 'INFO',
                    'message': f'未检测到运行中的 [{progId}]，尝试新建进程...',
                    'plugId': 'PluginRunner',
                    'filename': os.path.basename(__file__),
                    'funcName': funcName,
                    'lineno': lineno,
                }
            )
            comApp = win32com.client.Dispatch(progId)
            comApp.Visible = True

        # 加载插件入口模组
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

        spec = importlib.util.spec_from_file_location(f'ext_{pluginFolderName}', targetEntry)
        assert spec is not None and spec.loader is not None, f"无法从 {targetEntry} 加载模块规范"
        pluginModule = importlib.util.module_from_spec(spec)
        sys.modules[f'ext_{pluginFolderName}'] = pluginModule
        spec.loader.exec_module(pluginModule)

        # 执行插件入口 run 函数
        if hasattr(pluginModule, 'run'):
            pluginModule.run(comApp, logQueue)
            frame = inspect.currentframe()
            funcName = frame.f_code.co_name if frame else '<unknown>'
            lineno = frame.f_lineno if frame else -1
            logQueue.put(
                {
                    'type': 'SUCCESS',
                    'message': f'插件 [{pluginFolderName}] 执行完毕。',
                    'plugId': 'PluginRunner',
                    'filename': os.path.basename(__file__),
                    'funcName': funcName,
                    'lineno': lineno,
                }
            )
        else:
            raise AttributeError(
                f'插件 [{pluginFolderName}] 未定义 run(appComHandle, logQueue) 函数！'
            )

    except Exception:
        import traceback

        errMsg = traceback.format_exc()
        frame = inspect.currentframe()
        funcName = frame.f_code.co_name if frame else '<unknown>'
        lineno = frame.f_lineno if frame else -1
        logQueue.put(
            {
                'type': 'ERROR',
                'message': f'插件执行异常:\n{errMsg}',
                'plugId': 'PluginRunner',
                'filename': os.path.basename(__file__),
                'funcName': funcName,
                'lineno': lineno,
            }
        )
    finally:
        pythoncom.CoUninitialize()