'''
file: constants.py
description: 常量定义
author: IYATT-yx
copyright:  Copyright (c) 2026 IYATT-yx.
            Licensed under the MIT License. See LICENSE file in the project root for full license information.
'''
import os
import sys

class Path:
    isPackaged = '__compiled__' in globals()
    appDir = os.path.dirname(os.path.abspath(sys.argv[0]))
    config = os.path.join(appDir, 'config.ini')
    log = os.path.join(appDir, 'logs', 'PyTableEngine.log')
    vendor = os.path.join(appDir, 'vendor')
    extensions = os.path.join(appDir, 'extensions')
    cache = os.path.join(appDir, 'cache')
    repositoryIndexFile = os.path.join(cache, 'index.json')

class Command:
    argv0 = os.path.abspath(sys.argv[0])
    executableCommandString = argv0 if Path.isPackaged else sys.executable + ' ' + argv0

class Site:
    pypi =  'https://mirrors.tuna.tsinghua.edu.cn/pypi/web/simple'
    repositoryIndex = 'https://raw.githubusercontent.com/IYATT-yx/PyTableEnginePluginRepository/main/index.json'

class Web:
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/152.0.0.0 Safari/537.36 Edg/152.0.0.0',
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'zh-CN,zh;q=0.9,en-US;q=0.8,en-GB;q=0.7,en;q=0.6',
        'Connection': 'close',
    }