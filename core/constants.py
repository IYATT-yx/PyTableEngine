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

class Command:
    argv0 = os.path.abspath(sys.argv[0])
    executableCommandString = argv0 if Path.isPackaged else sys.executable + ' ' + argv0

class Plugin:
    repositoryIndex = 'https://raw.githubusercontent.com/IYATT-yx/PyTableEnginePluginRepository/main/index.json'