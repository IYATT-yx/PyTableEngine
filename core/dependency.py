'''
file: dependency.py
description: 依赖检查与安装管理
author: IYATT-yx
copyright:  Copyright (c) 2026 IYATT-yx.
            Licensed under the MIT License. See LICENSE file in the project root for full license information.
'''
import configparser
import os
import subprocess
import sys


class DependencyManager:
    '''插件依赖检查与安装管理'''

    def __init__(self, configPath):
        self.configPath = configPath
        self.config = configparser.ConfigParser()
        if os.path.exists(self.configPath):
            self.config.read(self.configPath, encoding='utf-8')

    def getPipOption(self, key):
        if self.config.has_section('pip') and key in self.config['pip']:
            val = self.config['pip'][key].strip()
            return val if val else None
        return None

    def checkAndInstallDependencies(self, pluginDir, logger_func):
        reqPath = os.path.join(pluginDir, 'requirements.txt')
        if not os.path.exists(reqPath):
            return

        vendorDir = os.path.join(pluginDir, 'vendor')
        os.makedirs(vendorDir, exist_ok=True)

        with open(reqPath, 'r', encoding='utf-8') as f:
            lines = [
                line.strip()
                for line in f
                if line.strip() and not line.startswith('#')
            ]

        missingPackages = []
        for line in lines:
            pkgName = line.split('==')[0].split('>=')[0].strip()
            if not self.isPackageAvailable(pkgName, vendorDir):
                missingPackages.append(line)

        if not missingPackages:
            return

        logger_func(f'检测到缺失的依赖项 {missingPackages}，准备隔离安装...', 'INFO')
        indexUrl = self.getPipOption('indexUrl')
        proxy = self.getPipOption('proxy')

        cmd = [
            sys.executable,
            '-m',
            'pip',
            'install',
            '-t',
            vendorDir,
            '-r',
            reqPath,
        ]
        if indexUrl:
            cmd.extend(['-i', indexUrl])
        if proxy:
            cmd.extend(['--proxy', proxy])

        try:
            res = subprocess.run(
                cmd, capture_output=True, text=True, check=True
            )
            logger_func(f'依赖安装成功: {res.stdout.strip()}', 'INFO')
        except Exception as err:
            logger_func(f'依赖隔离安装失败: {err}', 'ERROR')
            raise RuntimeError(f'依赖安装失败，请检查网络或配置: {err}')

    def isPackageAvailable(self, pkgName, vendorDir):
        sysPathBackup = list(sys.path)
        if vendorDir not in sys.path:
            sys.path.insert(0, vendorDir)
        try:
            importlib.import_module(pkgName)
            return True
        except ImportError:
            return False
        finally:
            sys.path = sysPathBackup