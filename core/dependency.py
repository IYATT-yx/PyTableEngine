'''
file: dependency.py
description: 插件依赖项集中管理
author: IYATT-yx
copyright: Copyright (c) 2026 IYATT-yx.
           Licensed under the MIT License. See LICENSE file in the project root for full license information.
'''
import os
import subprocess
import sys
import tempfile
from importlib.metadata import Distribution
from pathlib import Path
from packaging.requirements import Requirement
from typing import Callable

from core.config import Config
from core import constants

class Dependency:
    '''插件依赖项集中管理器'''

    def __init__(self, config: Config):
        self.config = config
        self.lastError: str | None = None

    def getInstalledVendorPackages(self, vendorDir: str) -> dict[str, str]:
        '''
        扫描 vendor 目录下的 .dist-info，获取已安装包的名称与版本映射表
        返回格式: {'numpy': '2.5.2', 'requests': '2.31.0'}

        Args:
            vendorDir (str): vendor 目录路径
        '''
        installed: dict[str, str] = {}
        vendorPath = Path(vendorDir)
        if not vendorPath.exists():
            return installed

        distributions = Distribution.discover(path=[str(vendorPath)])
        for dist in distributions:
            pkgName = dist.metadata['Name'].lower()
            pkgVersion = dist.version
            installed[pkgName] = pkgVersion

        return installed

    def collectAndInstallPluginDependencies(self, extensionsDir: str, logCallback: Callable[[str, str], None]) -> bool:
        '''
        扫描所有插件的 requirements.txt，进行预检；仅在缺失或版本不匹配时增量安装
        '''
        self.lastError = None

        def log(msg: str, level: str = 'INFO'):
            logCallback(msg, level)

        vendorDir = constants.Path.vendor
        if not os.path.exists(vendorDir):
            os.makedirs(vendorDir, exist_ok=True)

        allRequirements: set[str] = set()
        if os.path.exists(extensionsDir):
            for folder in os.listdir(extensionsDir):
                pluginDir = os.path.join(extensionsDir, folder)
                reqFile = os.path.join(pluginDir, 'requirements.txt')
                if os.path.isdir(pluginDir) and os.path.exists(reqFile):
                    with open(reqFile, 'r', encoding='utf-8') as f:
                        for line in f:
                            line = line.strip()
                            if line and not line.startswith('#'):
                                allRequirements.add(line)

        if not allRequirements:
            log('未检测到任何插件依赖声明，跳过依赖安装。', level='INFO')
            return True

        log(
            f'检测到总计 {len(allRequirements)} 条依赖声明，正在预检 vendor 目录...',
            level='INFO',
        )

        installedVendorPkgs = self.getInstalledVendorPackages(vendorDir)

        missingOrOutdatedReqs: list[str] = []
        for reqStr in allRequirements:
            try:
                req = Requirement(reqStr)
                reqName = req.name.lower()

                if reqName not in installedVendorPkgs:
                    log(f'依赖缺失: [{reqStr}]，加入安装列表', level='INFO')
                    missingOrOutdatedReqs.append(reqStr)
                else:
                    currentVer = installedVendorPkgs[reqName]
                    if req.specifier and not req.specifier.contains(currentVer):
                        log(
                            f'依赖版本不匹配: [{reqStr}] (当前 vendor 版本为 {currentVer})，准备更新',
                            level='WARNING',
                        )
                        missingOrOutdatedReqs.append(reqStr)
                    else:
                        log(
                            f'依赖已满足: [{reqStr}] (已有版本 {currentVer})，跳过安装。',
                            level='INFO',
                        )
            except Exception as e:
                log(f'解析依赖声明 [{reqStr}] 失败 ({e})，将强制交由 pip 处理', level='WARNING')
                missingOrOutdatedReqs.append(reqStr)

        if not missingOrOutdatedReqs:
            log('所有插件依赖项均已满足，无需重复安装！', level='SUCCESS')
            return True

        try:
            import pip
        except ImportError:
            errMsg = '环境致命错误: 当前 Python 解释器缺少 pip 模块，无法安装依赖包！'
            log(errMsg, level='ERROR')
            self.lastError = errMsg
            return False

        log(
            f'需安装/更新 {len(missingOrOutdatedReqs)} 项依赖，交由 pip 安装至 vendor 目录...',
            level='INFO',
        )

        tmpReqPath = None
        try:
            with tempfile.NamedTemporaryFile(
                mode='w+', delete=False, suffix='_req.txt', encoding='utf-8'
            ) as tmp:
                tmp.write('\n'.join(missingOrOutdatedReqs))
                tmpReqPath = tmp.name

            cmd = [
                sys.executable,
                '-m',
                'pip',
                'install',
                '-v',
                '--target',
                vendorDir,
                '--upgrade',
            ]

            # 读取 [pip] pypi 配置
            pypiIndexUrl = self.config.getCleanOption('pip', 'pypi')
            if pypiIndexUrl:
                cmd.extend(['-i', pypiIndexUrl])
                log(f'正在使用镜像源地址: {pypiIndexUrl}', level='INFO')
            else:
                log('未配置或留空 PyPI 镜像源，将使用 pip 默认源', level='INFO')

            # 读取 [pip] proxy 配置
            pipProxy = self.config.getCleanOption('pip', 'proxy')
            if pipProxy:
                cmd.extend(['--proxy', pipProxy])
                log(f'正在使用网络代理: {pipProxy}', level='INFO')
            else:
                log('未配置或留空网络代理，将直接进行连接', level='INFO')

            cmd.extend(['-r', tmpReqPath])

            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding='utf-8',
                errors='replace',
                bufsize=1,
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0,
            )

            hasError: bool = False
            collectedErrors: list[str] = []
            isInTraceback: bool = False

            if process.stdout:
                for line in iter(process.stdout.readline, ''):
                    cleanLine = line.strip()
                    if not cleanLine:
                        continue

                    if 'Traceback (most recent call last):' in cleanLine or cleanLine.startswith('ERROR:'):
                        isInTraceback = True
                        hasError = True

                    if (
                        hasError
                        or cleanLine.startswith('ERROR:')
                        or 'No module named' in cleanLine
                        or 'PermissionError' in cleanLine
                        or 'WinError' in cleanLine
                        or 'Access is denied' in cleanLine
                    ):
                        log(cleanLine, level='ERROR')
                        hasError = True
                        collectedErrors.append(cleanLine)
                    elif cleanLine.startswith('WARNING:'):
                        log(cleanLine, level='WARNING')
                    else:
                        log(cleanLine, level='INFO')

                process.stdout.close()

            process.wait()

            if process.returncode == 0 and not hasError:
                log('插件缺失依赖项已成功集中安装至 vendor 目录！', level='SUCCESS')
                return True
            else:
                errDetail = "\n".join(collectedErrors) if collectedErrors else f"exit code: {process.returncode}"
                errMsg = f'pip 执行失败:\n{errDetail}'
                log('pip 执行失败，请检查 Python 环境或依赖声明！', level='ERROR')
                self.lastError = errMsg
                return False
        except Exception as e:
            errMsg = f'执行 pip 安装时发生致命异常: {e}'
            log(errMsg, level='ERROR')
            self.lastError = errMsg
            return False
        finally:
            if tmpReqPath and os.path.exists(tmpReqPath):
                try:
                    os.remove(tmpReqPath)
                except OSError:
                    pass