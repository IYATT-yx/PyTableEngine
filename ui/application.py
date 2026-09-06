'''
file: application.py
description: 主界面
author: IYATT-yx
copyright:  Copyright (c) 2026 IYATT-yx.
            Licensed under the MIT License. See LICENSE file in the project root for full license information.
'''
import configparser
import importlib.util
import multiprocessing
import os
import time
import tkinter as tk
from tkinter import messagebox, ttk

import pythoncom
import win32com.client

from buildtime import buildTime
from core.dependency import DependencyManager
from core.logger import AppLogger
from core.runner import pluginRunnerTask
from core import constants


class MainWindow:
    '''Tkinter 主界面管理类'''

    def __init__(self, root):
        self.root = root
        self.root.title(f'PyTabEngine 思能快表引擎 by IYATT-yx {buildTime}')
        self.root.geometry('880x580')
        iconPath = os.path.join(constants.Path.appDir, 'icon.ico')
        if os.path.exists(iconPath):
            self.root.iconbitmap(iconPath)

        self.logger = AppLogger.setupLogger()
        self.configPath = os.path.join(constants.Path.appDir, 'config.ini')
        self.depManager = DependencyManager(self.configPath)

        self.comProgIdMap = {
            'Excel (Microsoft Office)': 'Excel.Application',
            'WPS 表格 (WPS Office)': 'ET.Application',
            'WPS 兼容接口': 'Ket.Application',
        }

        self.loadedPlugins = []
        self.activeProcess = None
        self.logQueue = multiprocessing.Queue()

        self.initUiComponents()
        self.scanPlugins()
        self.listenLogQueue()

    def initUiComponents(self):
        # 1. 顶部 COM 接口选择栏
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

        # 2. 中间搜索过滤与插件表格
        midFrame = ttk.LabelFrame(self.root, text=' 插件列表 ', padding=10)
        midFrame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        searchFrame = ttk.Frame(midFrame)
        searchFrame.pack(fill=tk.X, pady=5)

        ttk.Label(searchFrame, text='搜索插件:').pack(side=tk.LEFT, padx=5)
        self.searchVar = tk.StringVar()
        self.searchVar.trace_add('write', self.filterPlugins)
        searchEntry = ttk.Entry(
            searchFrame, textvariable=self.searchVar, width=25
        )
        searchEntry.pack(side=tk.LEFT, padx=5)

        toggleBtn = ttk.Button(
            searchFrame, text='切换启用/禁用', command=self.togglePluginState
        )
        toggleBtn.pack(side=tk.RIGHT, padx=5)

        refreshBtn = ttk.Button(
            searchFrame, text='刷新列表', command=self.scanPlugins
        )
        refreshBtn.pack(side=tk.RIGHT, padx=5)

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

        try:
            self.depManager.checkAndInstallDependencies(pluginDir, self.appendLog)
        except Exception as err:
            messagebox.showerror('依赖错误', str(err))
            return

        self.activeProcess = multiprocessing.Process(
            target=pluginRunnerTask,
            args=(pluginDir, folderName, progId, self.logQueue, 'config.ini'),
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

    def appendLog(self, message, level='INFO', plugId='System', filename=None, funcName=None, lineno=None, stacklevel=2):
        '''统一日志接口：兼容主进程与子进程日志'''
        timestamp = time.strftime('%Y-%m-%d %H:%M:%S')
        formatted_ui_msg = f'[{timestamp}] [{level}] [{plugId}] {message}\n'
        
        # 1. 输出至 UI 日志组件
        self.logText.config(state=tk.NORMAL)
        tag = level if level in ['INFO', 'WARNING', 'ERROR', 'SUCCESS'] else 'INFO'
        self.logText.insert(tk.END, formatted_ui_msg, tag)
        self.logText.see(tk.END)
        self.logText.config(state=tk.DISABLED)

        # 2. 构造 extra（包含 plugId，如果存在源码位置也一并传入供 makeRecord 拦截）
        extra_info = {'plugId': plugId}
        call_stacklevel = stacklevel
        
        if filename and funcName and lineno:
            extra_info.update({
                'filename': filename,
                'funcName': funcName,
                'lineno': lineno
            })
            call_stacklevel = 1 

        # 3. 写入日志文件
        if level == 'ERROR':
            self.logger.error(message, extra=extra_info, stacklevel=call_stacklevel)
        elif level == 'WARNING':
            self.logger.warning(message, extra=extra_info, stacklevel=call_stacklevel)
        else:
            self.logger.info(message, extra=extra_info, stacklevel=call_stacklevel)