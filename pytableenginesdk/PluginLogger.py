"""
file: PluginLogger.py
description: 插件日志代理
author: IYATT-yx
copyright:  Copyright (c) 2026 IYATT-yx.
            Licensed under the MIT License. See LICENSE file in the project root for full license information.
"""
import multiprocessing
import logging
import inspect
import os
from tkinter import messagebox

class PluginLogger:
    '''进程间通信日志代理类（供子进程插件使用）'''
    
    INFO = logging.INFO
    WARNING = logging.WARNING
    ERROR = logging.ERROR
    isInit = False

    def __init__(self, logQueue: multiprocessing.Queue, plugInfo: dict):
        '''
        插件日志代理类

        Args:
            logQueue (multiprocessing.Queue): 日志队列
            plugInfo (dict): 插件信息
        '''
        self.queue = logQueue
        author = plugInfo.get('author', 'UnknownAuthor')
        name = plugInfo.get('name', 'UnknownPlugin')
        self.plugId = f"{author}_{name}"
        self.isInit = True

    def _Log(self, level: int=INFO, msg: str = ''):
        '''
        日志记录

        Args:
            level (int): 日志级别
            msg (str): 日志内容
        '''
        if not self.isInit:
            messagebox.showerror('错误', '日志代理未初始化')
            return

        frame = inspect.currentframe().f_back.f_back
        filename = os.path.basename(frame.f_code.co_filename)
        funcName = frame.f_code.co_name
        lineno = frame.f_lineno

        logData = {
            'level': None,
            'message': msg,
            'plugId': self.plugId,
            'filename': filename,
            'funcName': funcName,
            'lineno': lineno,
        }

        match level:
            case self.INFO:
                logData['level'] = 'INFO'
            case self.WARNING:
                logData['level'] = 'WARNING'
            case self.ERROR:
                logData['level'] = 'ERROR'

        self.queue.put(logData)

    def info(self, msg: str = ''):
        '''
        日志记录

        Args:
            msg (str): 日志内容
        '''
        self._Log(self.INFO, msg)

    def warning(self, msg: str = ''):
        '''
        警告日志记录

        Args:
            msg (str): 日志内容
        '''
        self._Log(self.WARNING, msg)

    def error(self, msg: str = ''):
        '''
        错误日志记录

        Args:
            msg (str): 日志内容
        '''
        self._Log(self.ERROR, msg)