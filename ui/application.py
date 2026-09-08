'''
file: application.py
description: 主界面
author: IYATT-yx
copyright:   Copyright (c) 2026 IYATT-yx.
             Licensed under the MIT License. See LICENSE file in the project root for full license information.
'''
import configparser
import importlib.util
import multiprocessing
import os
import sys
import threading
import time
import tkinter as tk
from tkinter import messagebox, ttk
import inspect
import subprocess

import pythoncom
import win32com.client

from buildtime import buildTime
from core import constants
from core.config import Config
from core.dependency import Dependency
from core.logger import AppLogger
from core.runner import pluginRunnerTask

class MainWindow:
    '''Tkinter 主界面管理类'''

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(
            f'PyTableEngine 思能快表引擎 by IYATT-yx {buildTime}'
        )
        self.root.geometry('880x580')
        iconPath = os.path.join(constants.Path.appDir, 'icon.ico')
        self.root.iconbitmap(iconPath)

        self.setupVendorPath()

        self.logger = AppLogger.setupLogger()

        # 配置与依赖管理器初始化
        self.config = Config()
        self.depManager = Dependency(self.config)

        self.comProgIdMap = {
            'Excel (Microsoft Office)': 'Excel.Application',
            'WPS 表格 (WPS Office)': 'ET.Application',
            'WPS 兼容接口': 'Ket.Application',
        }

        self.loadedPlugins = []
        self.activeProcess = None
        self.logQueue = multiprocessing.Queue()
        self.configPath = constants.Path.config

        # 初始化 UI 组件
        self.initUiComponents()

        # 启动日志监听
        self.listenLogQueue()

        # 延迟 100ms 触发依赖检查与插件扫描（确保界面先渲染出来）
        self.root.after(100, self.processDependenciesAndScan)

    def restartSystem(self):
        '''重启当前应用程序进程'''
        if messagebox.askyesno('确认重启', '重启系统将重新加载所有 Python 环境与插件，是否继续？'):
            self.appendLog('正在准备重启系统...', level='INFO')
            self.root.destroy()
            subprocess.Popen(constants.Command.executableCommandString)
            sys.exit(0)

    def refreshPluginsAndDependencies(self):
        '''手动刷新：重新检查安装依赖项，并扫描重新加载所有插件'''
        self.processDependenciesAndScan()

    def setupVendorPath(self):
        '''给主进程挂载 vendor 目录及 Windows DLL 路径'''
        vendorDir = constants.Path.vendor
        if os.path.exists(vendorDir):
            if vendorDir not in sys.path:
                sys.path.insert(0, vendorDir)

            # 挂载 DLL 路径（针对 NumPy/SciPy 等 C 扩展）
            if hasattr(os, 'add_dll_directory'):
                try:
                    os.add_dll_directory(vendorDir)
                except Exception:
                    pass

            for item in os.listdir(vendorDir):
                if item.endswith('.libs'):
                    libsDir = os.path.join(vendorDir, item)
                    if os.path.isdir(libsDir):
                        if hasattr(os, 'add_dll_directory'):
                            try:
                                os.add_dll_directory(libsDir)
                            except Exception:
                                pass
                        if libsDir not in os.environ.get('PATH', ''):
                            os.environ['PATH'] = libsDir + os.path.pathsep + os.environ.get('PATH', '')

            binDir = os.path.join(vendorDir, 'bin')
            if os.path.exists(binDir) and os.path.isdir(binDir):
                if hasattr(os, 'add_dll_directory'):
                    try:
                        os.add_dll_directory(binDir)
                    except Exception:
                        pass
                if binDir not in os.environ.get('PATH', ''):
                    os.environ['PATH'] = binDir + os.path.pathsep + os.environ.get('PATH', '')

    def setUiInteractive(self, enabled: bool):
        '''
        控制界面关键控件的可交互状态，防止依赖未安装完成时用户误操作导致报错
        '''
        targetState = 'normal' if enabled else 'disabled'
        comboState = 'readonly' if enabled else 'disabled'

        # 禁用/启用顶部连接与切换按钮
        if hasattr(self, 'comCombo'):
            self.comCombo.config(state=comboState)
        if hasattr(self, 'connectBtn'):
            self.connectBtn.config(state=targetState)

        # 禁用/启用搜索与操作按钮
        if hasattr(self, 'toggleBtn'):
            self.toggleBtn.config(state=targetState)
        if hasattr(self, 'refreshBtn'):
            self.refreshBtn.config(state=targetState)
        if hasattr(self, 'searchEntry'):
            self.searchEntry.config(state=targetState)

        # 禁用/启用插件列表交互（通过 Treeview 选择模式控制）
        if hasattr(self, 'pluginTree'):
            self.pluginTree.config(selectmode='browse' if enabled else 'none')

    def processDependenciesAndScan(self):
        self.setUiInteractive(False)
        self.appendLog('正在准备检查插件依赖项...', level='INFO')

        extensionsDir = os.path.join(constants.Path.appDir, 'extensions')

        def workerTask():
            def logCallback(msg: str, level: str='INFO'):
                # 使用 inspect 获取当前的栈帧对象
                frame = inspect.currentframe()
                try:
                    callerFrame = frame.f_back if frame else None
                    callerFrame = callerFrame.f_back if callerFrame else None
                    if callerFrame:
                        file_name = os.path.basename(callerFrame.f_code.co_filename)
                        func_name = callerFrame.f_code.co_name
                        line_no = callerFrame.f_lineno
                    else:
                        file_name = func_name = line_no = None
                finally:
                    # 显式释放 frame 引用，防止在异常或复杂调用栈中产生循环引用导致内存泄漏
                    del frame

                # 将抓取到的具体源码位置作为闭包默认参数绑定，推入主线程事件队列
                self.root.after(
                    0, 
                    lambda f=file_name, fn=func_name, l=line_no: self.appendLog(
                        msg, 
                        level=level, 
                        plugId='Dependency', 
                        filename=f, 
                        funcName=fn, 
                        lineno=l
                    )
                )

            isSuccess = False
            errorMsg = None

            try:
                logCallback('开始检查插件依赖项...', level='INFO')
                isSuccess = self.depManager.collectAndInstallPluginDependencies(
                    extensionsDir, logCallback
                )
            except Exception as err:
                errorMsg = str(err)

            self.root.after(0, lambda: self.onDependenciesFinished(isSuccess, errorMsg))

        depThread = threading.Thread(target=workerTask, daemon=True)
        depThread.start()

    def onDependenciesFinished(self, isSuccess: bool, errorMsg: str | None = None):
        '''依赖项处理完毕后的回调函数（运行在主线程）'''
        
        # 优先处理安装失败/异常的情况
        if not isSuccess or errorMsg:
            final_err = errorMsg or self.depManager.lastError or ""
            
            if final_err:
                # 转为小写进行不区分大小写匹配
                err_lower = final_err.lower()
                lock_keywords = [
                    'permissionerror', 
                    'winerror 5', 
                    '_handle_target_dir', 
                    'rmtree', 
                    'access is denied',
                    '拒绝访问'
                ]
                is_file_locked = any(kw in err_lower for kw in lock_keywords)

                if is_file_locked:
                    result = messagebox.askyesno(
                        '依赖文件锁冲突',
                        '检测到部分依赖文件（如 .pyd / .dll）正被当前应用程序或后台进程占用，导致 pip 无法覆盖更新。\n\n'
                        '是否立即重启系统以释放文件锁并自动完成安装？'
                    )
                    if result:
                        self.restartSystem()
                        return  # 用户选择重启，直接退出
                else:
                    messagebox.showerror('依赖错误', f'处理插件依赖时发生异常:\n{final_err}')
            else:
                messagebox.showwarning('依赖警告', '部分依赖项安装失败，相关插件可能无法正常运行！')

        importlib.invalidate_caches()
        self.setupVendorPath()
        self.scanPlugins()
        self.setUiInteractive(True)
        
        if isSuccess:
            self.appendLog('系统刷新完毕，插件环境与列表均已重载。', level='SUCCESS')
        else:
            self.appendLog('依赖项未完全就绪，系统已恢复交互，请检查错误日志。', level='WARNING')

    def initUiComponents(self):
        # 顶部 COM 接口选择栏
        topFrame = ttk.LabelFrame(self.root, text=' COM 接口连接 ', padding=10)
        topFrame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(topFrame, text='选择目标表格软件:').pack(
            side=tk.LEFT, padx=5
        )
        self.comCombo = ttk.Combobox(
            topFrame,
            values=list(self.comProgIdMap.keys()),
            state='readonly',
            width=28,
        )
        self.comCombo.current(0)
        self.comCombo.pack(side=tk.LEFT, padx=5)

        self.connectBtn = ttk.Button(
            topFrame, text='测试/连接 COM', command=self.connectCom
        )
        self.connectBtn.pack(side=tk.LEFT, padx=5)

        self.comStatusLabel = ttk.Label(
            topFrame, text='未连接', foreground='gray'
        )
        self.comStatusLabel.pack(side=tk.LEFT, padx=10)

        # 中间搜索过滤与插件表格
        midFrame = ttk.LabelFrame(self.root, text=' 插件列表 ', padding=10)
        midFrame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        searchFrame = ttk.Frame(midFrame)
        searchFrame.pack(fill=tk.X, pady=5)

        ttk.Label(searchFrame, text='搜索插件:').pack(side=tk.LEFT, padx=5)
        self.searchVar = tk.StringVar()
        self.searchVar.trace_add('write', self.filterPlugins)
        self.searchEntry = ttk.Entry(
            searchFrame, textvariable=self.searchVar, width=25
        )
        self.searchEntry.pack(side=tk.LEFT, padx=5)

        self.toggleBtn = ttk.Button(
            searchFrame, text='切换启用/禁用', command=self.togglePluginState
        )
        self.toggleBtn.pack(side=tk.RIGHT, padx=5)

        self.restartBtn = ttk.Button(
            searchFrame, text='重启系统', command=self.restartSystem
        )
        self.restartBtn.pack(side=tk.RIGHT, padx=5)

        self.refreshBtn = ttk.Button(
            searchFrame, text='重新扫描', command=self.refreshPluginsAndDependencies
        )
        self.refreshBtn.pack(side=tk.RIGHT, padx=5)

        columns = ('name', 'author', 'status', 'description', 'version', 'entryType')
        self.pluginTree = ttk.Treeview(
            midFrame, columns=columns, show='headings', selectmode='browse'
        )

        self.pluginTree.heading('name', text='插件名')
        self.pluginTree.heading('author', text='作者')
        self.pluginTree.heading('status', text='状态')
        self.pluginTree.heading('description', text='功能描述')
        self.pluginTree.heading('version', text='版本')
        self.pluginTree.heading('entryType', text='文件类型')

        self.pluginTree.column('name', width=130)
        self.pluginTree.column('author', width=100)
        self.pluginTree.column('status', width=80, anchor=tk.CENTER)
        self.pluginTree.column('description', width=320)
        self.pluginTree.column('version', width=60, anchor=tk.CENTER)
        self.pluginTree.column('entryType', width=80, anchor=tk.CENTER)

        self.pluginTree.tag_configure('ENABLED', foreground='black')
        self.pluginTree.tag_configure('DISABLED', foreground='gray')
        self.pluginTree.tag_configure('LOAD_FAILED', foreground='orange')
        self.pluginTree.tag_configure('VALID_FAILED', foreground='red')

        treeScroll = ttk.Scrollbar(
            midFrame, orient=tk.VERTICAL, command=self.pluginTree.yview
        )
        self.pluginTree.configure(yscrollcommand=treeScroll.set)

        self.pluginTree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        treeScroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.pluginTree.bind('<Double-1>', self.onPluginDoubleClick)

        self.contextMenu = tk.Menu(self.root, tearoff=0)
        self.contextMenu.add_command(
            label='切换 启用/禁用', command=self.togglePluginState
        )
        self.pluginTree.bind('<Button-3>', self.showContextMenu)

        # 3. 底部运行日志栏
        bottomFrame = ttk.LabelFrame(self.root, text=' 运行日志 ', padding=5)
        bottomFrame.pack(fill=tk.X, padx=10, pady=5)

        self.logText = tk.Text(bottomFrame, height=6, state=tk.DISABLED, wrap=tk.WORD)
        logScroll = ttk.Scrollbar(bottomFrame, orient=tk.VERTICAL, command=self.logText.yview)
        self.logText.configure(yscrollcommand=logScroll.set)

        self.logText.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        logScroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 配置日志颜色标签
        self.logText.tag_config('INFO', foreground='#2b2b2b')
        self.logText.tag_config('WARNING', foreground='#d97706')
        self.logText.tag_config('ERROR', foreground='#dc2626')
        self.logText.tag_config('SUCCESS', foreground='#16a34a')

    def showContextMenu(self, event):
        item = self.pluginTree.identify_row(event.y)
        if item:
            self.pluginTree.selection_set(item)
            self.contextMenu.post(event.x_root, event.y_root)

    def getDisabledPluginsFromConfig(self):
        config = configparser.ConfigParser()
        if os.path.exists(self.configPath):
            config.read(self.configPath, encoding='utf-8')
        if config.has_section('disabled_plugins'):
            return dict(config.items('disabled_plugins'))
        return {}

    def setPluginDisabledConfig(self, pluginKey, isDisabled):
        config = configparser.ConfigParser()
        if os.path.exists(self.configPath):
            config.read(self.configPath, encoding='utf-8')
        if not config.has_section('disabled_plugins'):
            config.add_section('disabled_plugins')

        if isDisabled:
            config.set('disabled_plugins', pluginKey, 'true')
        else:
            config.remove_option('disabled_plugins', pluginKey)

        with open(self.configPath, 'w', encoding='utf-8') as f:
            config.write(f)

    def connectCom(self):
        displayName = self.comCombo.get()
        progId = self.comProgIdMap[displayName]
        pythoncom.CoInitialize()
        try:
            win32com.client.GetActiveObject(progId)
            self.comStatusLabel.config(
                text=f'已成功连接到已打开的 {displayName}',
                foreground='green',
            )
            self.appendLog(f'成功连接至正在运行的 COM 实例: {progId}', level='SUCCESS')
        except Exception:
            self.comStatusLabel.config(
                text=f'当前未运行 {displayName} (双击插件时将自动创建)',
                foreground='orange',
            )
            self.appendLog(
                f'当前未寻找到活跃的 {progId} 实例，已做好延迟启动准备。',
                level='WARNING'
            )
        finally:
            pythoncom.CoUninitialize()

    def scanPlugins(self):
        self.loadedPlugins.clear()
        extensionsDir = os.path.join(constants.Path.appDir, 'extensions')
        if not os.path.exists(extensionsDir):
            os.makedirs(extensionsDir, exist_ok=True)

        disabledMap = self.getDisabledPluginsFromConfig()

        for folder in os.listdir(extensionsDir):
            pluginDir = os.path.join(extensionsDir, folder)
            if not os.path.isdir(pluginDir):
                continue

            pydFile = os.path.join(pluginDir, f'{folder}.pyd')
            pyFile = os.path.join(pluginDir, f'{folder}.py')

            entryType = None
            if os.path.exists(pydFile):
                entryType = 'PYD 闭源'
            elif os.path.exists(pyFile):
                entryType = 'PY 源码'

            if not entryType:
                continue

            pluginInfo = {
                'folderName': folder,
                'name': '未知插件',
                'author': '未知作者',
                'description': '未提供描述信息',
                'version': '1.0.0',
                'entryType': entryType,
                'dirPath': pluginDir,
                'status': '已加载',
                'tag': 'ENABLED',
            }

            if '_' not in folder:
                pluginInfo['status'] = '校验失败'
                pluginInfo['tag'] = 'VALID_FAILED'
                pluginInfo['description'] = '目录格式非 [作者名]_[插件名]'
                self.loadedPlugins.append(pluginInfo)
                continue

            folderAuthor, folderName = folder.split('_', 1)
            pluginInfo['author'] = folderAuthor
            pluginInfo['name'] = folderName

            if entryType == 'PY 源码':
                try:
                    spec = importlib.util.spec_from_file_location(
                        f'info_{folder}', pyFile
                    )
                    assert spec is not None and spec.loader is not None, f'无法为插件 [{folder}] 创建模块规范或加载器'
                    mod = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(mod)

                    if hasattr(mod, 'pluginInfo'):
                        rawInfo = mod.pluginInfo
                        metaAuthor = rawInfo.get('author', '').strip()
                        metaName = rawInfo.get('name', '').strip()

                        if metaAuthor != folderAuthor or metaName != folderName:
                            pluginInfo['status'] = '校验失败'
                            pluginInfo['tag'] = 'VALID_FAILED'
                            pluginInfo['description'] = (
                                f'元信息({metaAuthor}_{metaName})与目录({folder})不符'
                            )
                        else:
                            pluginInfo['description'] = rawInfo.get(
                                'description', '未提供描述信息'
                            )
                            pluginInfo['version'] = rawInfo.get(
                                'version', '1.0.0'
                            )
                except Exception as e:
                    pluginInfo['status'] = '加载失败'
                    pluginInfo['tag'] = 'LOAD_FAILED'
                    pluginInfo['description'] = f'静态解析失败: {e}'

            if pluginInfo['status'] == '已加载':
                pluginKey = f"{pluginInfo['author']}_{pluginInfo['name']}".lower()
                if pluginKey in disabledMap:
                    pluginInfo['status'] = '已禁用'
                    pluginInfo['tag'] = 'DISABLED'

            self.loadedPlugins.append(pluginInfo)

        self.renderPluginTree(self.loadedPlugins)
        
        msg = f'完成插件扫描，共加载 {len(self.loadedPlugins)} 个插件。'
        self.appendLog(msg, level='INFO')

    def renderPluginTree(self, pluginList):
        for item in self.pluginTree.get_children():
            self.pluginTree.delete(item)

        for info in pluginList:
            itemIid = f"{info['author']}_{info['name']}"
            self.pluginTree.insert(
                '',
                tk.END,
                iid=itemIid,
                values=(
                    info['name'],
                    info['author'],
                    info['status'],
                    info['description'],
                    info['version'],
                    info['entryType'],
                ),
                tags=(info['tag'],),
            )

    def filterPlugins(self, *args):
        query = self.searchVar.get().strip().lower()
        if not query:
            self.renderPluginTree(self.loadedPlugins)
            return

        filtered = [
            p
            for p in self.loadedPlugins
            if query in p['name'].lower()
            or query in p['author'].lower()
            or query in p['description'].lower()
        ]
        self.renderPluginTree(filtered)

    def getSelectedPlugin(self):
        selectedItems = self.pluginTree.selection()
        if not selectedItems:
            return None

        selectedIid = selectedItems[0]

        for p in self.loadedPlugins:
            pluginKey = f"{p['author']}_{p['name']}"
            if pluginKey == selectedIid:
                return p
        return None

    def togglePluginState(self):
        targetPlugin = self.getSelectedPlugin()
        if not targetPlugin:
            return

        if targetPlugin['status'] in ['校验失败', '加载失败']:
            messagebox.showwarning(
                '无法切换',
                f'插件 [{targetPlugin["author"]}_{targetPlugin["name"]}] 状态为 {targetPlugin["status"]}，拒绝切换！',
            )
            return

        pluginKey = f"{targetPlugin['author']}_{targetPlugin['name']}".lower()

        if targetPlugin['status'] == '已加载':
            targetPlugin['status'] = '已禁用'
            targetPlugin['tag'] = 'DISABLED'
            self.setPluginDisabledConfig(pluginKey, True)
            self.appendLog(
                f'插件 [{targetPlugin["author"]}_{targetPlugin["name"]}] 已禁用。',
                level='WARNING'
            )
        else:
            targetPlugin['status'] = '已加载'
            targetPlugin['tag'] = 'ENABLED'
            self.setPluginDisabledConfig(pluginKey, False)
            self.appendLog(
                f'插件 [{targetPlugin["author"]}_{targetPlugin["name"]}] 已恢复启用。',
                level='INFO'
            )

        self.renderPluginTree(self.loadedPlugins)

    def onPluginDoubleClick(self, event):
        targetPlugin = self.getSelectedPlugin()
        if not targetPlugin:
            return

        pluginFullName = f"{targetPlugin['author']}_{targetPlugin['name']}"

        if targetPlugin['status'] == '已禁用':
            messagebox.showwarning(
                '提示', f'插件 [{pluginFullName}] 当前已被禁用，请右键开启后再试！'
            )
            return
        elif targetPlugin['status'] in ['校验失败', '加载失败']:
            messagebox.showerror(
                '错误',
                f'插件 [{pluginFullName}] 校验未通过，不可执行！\n原因: {targetPlugin["description"]}',
            )
            return

        if self.activeProcess and self.activeProcess.is_alive():
            messagebox.showwarning(
                '提示', '当前已有插件正在后台执行，请等待其完成后再试！'
            )
            return

        self.executePlugin(targetPlugin)

    def executePlugin(self, pluginInfo):
        pluginDir = pluginInfo['dirPath']
        folderName = pluginInfo['folderName']
        progId = self.comProgIdMap[self.comCombo.get()]

        self.appendLog(
            f'准备启动插件 [{pluginInfo["author"]}_{pluginInfo["name"]}]...',
            level='INFO'
        )

        self.activeProcess = multiprocessing.Process(
            target=pluginRunnerTask,
            args=(pluginDir, folderName, progId, self.logQueue),
            daemon=True,
        )
        self.activeProcess.start()

    def listenLogQueue(self):
        '''监听子进程日志队列'''
        while not self.logQueue.empty():
            logData = self.logQueue.get_nowait()
            
            msgType = logData.get('level', 'INFO')
            msg = logData.get('message', '')
            plugId = logData.get('plugId', 'UnknownPlugin')
            
            filename = logData.get('filename')
            funcName = logData.get('funcName')
            lineno = logData.get('lineno')
            
            self.appendLog(
                msg, 
                level=msgType, 
                plugId=plugId, 
                filename=filename, 
                funcName=funcName, 
                lineno=lineno
            )

        self.root.after(200, self.listenLogQueue)

    def appendLog(self, message: str, level: str='INFO', plugId: str='System', filename:str|None=None, funcName:str|None=None, lineno:int|None=None, stacklevel:int=2):
        '''统一日志接口：兼容主进程与子进程日志'''
        timestamp = time.strftime('%Y-%m-%d %H:%M:%S')
        formattedUiMsg = f'[{timestamp}] [{level}] [{plugId}] {message}\n'
        
        # 输出至 UI 日志组件
        self.logText.config(state=tk.NORMAL)
        tag = level if level in ['INFO', 'WARNING', 'ERROR', 'SUCCESS'] else 'INFO'
        self.logText.insert(tk.END, formattedUiMsg, tag)
        self.logText.see(tk.END)
        self.logText.config(state=tk.DISABLED)

        # 构造 extra（包含 plugId，如果存在源码位置也一并传入供 makeRecord 拦截）
        extraInfo: dict[str, str|int] = {'plugId': plugId}
        callStacklevel = stacklevel
        
        if filename and funcName and lineno:
            extraInfo.update({
                'filename': filename,
                'funcName': funcName,
                'lineno': lineno
            })
            callStacklevel = 1 

        # 写入日志文件
        if level == 'ERROR':
            self.logger.error(message, extra=extraInfo, stacklevel=callStacklevel)
        elif level == 'WARNING':
            self.logger.warning(message, extra=extraInfo, stacklevel=callStacklevel)
        else:
            self.logger.info(message, extra=extraInfo, stacklevel=callStacklevel)