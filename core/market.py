'''
file: market.py
description: 插件市场管理器
author: IYATT-yx
copyright: Copyright (c) 2026 IYATT-yx.
           Licensed under the MIT License. See LICENSE file in the project root for full license information.
'''
import json
import os
import shutil
import ssl
import time
import urllib.request
from contextlib import contextmanager
from typing import Callable, Generator, Optional

from dulwich import porcelain

from core import constants
from core.config import Config


class Market:
    '''插件市场管理器'''

    def __init__(self, config: Config):
        self.config = config
        self.lastError: Optional[str] = None

    def downloadRepositoryIndex(self, logCallback: Optional[Callable[[str, str], None]] = None) -> bool:
        '''
        仅负责从远程下载插件索引 JSON 并安全写入 cache 目录。
        采用“写临时文件+原子替换”机制，下载/校验失败时绝不覆盖已有旧缓存。
        '''
        self.lastError = None
        def log(msg: str, level: str = 'INFO'):
            if logCallback:
                logCallback(msg, level)

        repoIndexUrl = self.config.getCleanOption('market', 'repo_index')
        if not repoIndexUrl:
            repoIndexUrl = constants.Site.repositoryIndex
            log(f'未检测到自定义 repo_index 配置，使用默认索引源: {repoIndexUrl}', level='INFO')
        else:
            log(f'正在使用自定义索引源: {repoIndexUrl}', level='INFO')

        proxyUrl = self._getProxyUrl()
        if proxyUrl:
            log(f'网络代理设置: {proxyUrl}', level='INFO')
        else:
            log('网络代理设置: 直连 (未配置)', level='INFO')

        cacheFilePath = constants.Path.repositoryIndexFile
        tempFilePath = cacheFilePath + '.tmp'
        startTime = time.time()

        try:
            opener = self._buildOpener()
            headers = constants.Web.headers
            req = urllib.request.Request(repoIndexUrl, headers=headers)
            
            log('正在向插件仓库发起请求...', level='INFO')
            with opener.open(req, timeout=30) as resp:
                elapsed = time.time() - startTime
                status = getattr(resp, 'status', 200)
                rawData = resp.read()
                dataLen = len(rawData)
                log(f'索引数据下载成功! 状态码: {status}, 耗时: {elapsed:.2f}s, 大小: {dataLen} 字节', level='INFO')

                # 校验 JSON 格式正确性后再写入，防止下载了错误页面/损坏内容
                indexData = json.loads(rawData.decode('utf-8'))
                if not isinstance(indexData, (dict, list)):
                    raise ValueError("下载的 JSON 数据格式不是预期的列表或字典")

                # 写入临时文件并替换目标缓存文件
                with open(tempFilePath, 'wb') as f:
                    f.write(rawData)
                os.replace(tempFilePath, cacheFilePath)

                log('插件索引已安全更新至本地缓存', level='INFO')
                return True

        except json.JSONDecodeError as e:
            errMsg = f'解析插件仓库 JSON 数据失败 (格式非法，保留原缓存): {e}'
            log(errMsg, level='ERROR')
            self.lastError = errMsg
        except Exception as e:
            elapsed = time.time() - startTime
            errMsg = f'拉取插件仓库索引失败 (用时 {elapsed:.2f}s，保留原缓存): {e}'
            log(errMsg, level='ERROR')
            self.lastError = errMsg

        # 出现异常时清理遗留的临时文件
        if os.path.exists(tempFilePath):
            try:
                os.remove(tempFilePath)
            except OSError:
                pass

        return False

    def getCachedRepositoryIndex(self, logCallback: Optional[Callable[[str, str], None]] = None) -> Optional[dict]:
        '''
        仅从本地 cache 中读取并解析插件索引 JSON（不发起网络请求）。
        支持启动时快速加载或离线模式下读取。
        '''
        self.lastError = None
        def log(msg: str, level: str = 'INFO'):
            if logCallback:
                logCallback(msg, level)

        cacheFilePath = constants.Path.repositoryIndexFile
        if not os.path.exists(cacheFilePath):
            errMsg = '本地暂无插件索引缓存文件'
            log(errMsg, level='WARNING')
            self.lastError = errMsg
            return None

        try:
            with open(cacheFilePath, 'r', encoding='utf-8') as f:
                parsedData = json.load(f)

            plugins = parsedData.get('plugins', {})
            pluginCount = len(plugins)
            pluginIds = list(plugins.keys())
            log(f'成功读取插件中心索引文件缓存，共包含 {pluginCount} 个插件: {pluginIds}', level='SUCCESS')
            return parsedData

        except json.JSONDecodeError as e:
            errMsg = f'读取本地索引缓存失败 (JSON格式损坏): {e}'
            log(errMsg, level='ERROR')
            self.lastError = errMsg
            return None
        except Exception as e:
            errMsg = f'读取本地索引缓存文件失败: {e}'
            log(errMsg, level='ERROR')
            self.lastError = errMsg
            return None

    def _getProxyUrl(self) -> Optional[str]:
        '''获取清洗后的代理配置'''
        return self.config.getCleanOption('market', 'proxy')

    def _buildOpener(self) -> urllib.request.OpenerDirector:
        '''根据配置构建支持代理与跳过 SSL 校验的 URL Opener'''
        handlers = []
        proxyUrl = self._getProxyUrl()
        if proxyUrl:
            handlers.append(urllib.request.ProxyHandler({
                'http': proxyUrl,
                'https': proxyUrl,
            }))

        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        handlers.append(urllib.request.HTTPSHandler(context=ctx))

        return urllib.request.build_opener(*handlers)

    @contextmanager
    def _temporaryOpener(self) -> Generator[None, None, None]:
        '''
        局域上下文管理器：临时将默认 urllib opener 替换为带代理/SSL设置的 opener。
        同时向系统环境变量中注入代理，确保 Dulwich 及底层 C 模块/子进程也能精准走代理。
        执行完成后立即还原，避免污染全局环境。
        '''
        default_opener = getattr(urllib.request, '_opener', None)
        proxyUrl = self._getProxyUrl()
        
        # 记录原有的环境变量，用于退栈时恢复
        old_http_proxy = os.environ.get('HTTP_PROXY')
        old_https_proxy = os.environ.get('HTTPS_PROXY')
        old_all_proxy = os.environ.get('ALL_PROXY')

        try:
            customOpener = self._buildOpener()
            urllib.request.install_opener(customOpener)

            if proxyUrl:
                os.environ['HTTP_PROXY'] = proxyUrl
                os.environ['HTTPS_PROXY'] = proxyUrl
                os.environ['ALL_PROXY'] = proxyUrl

            yield
        finally:
            # 还原 urllib Opener
            if default_opener is not None:
                urllib.request.install_opener(default_opener)
            else:
                urllib.request.install_opener(urllib.request.build_opener())

            # 还原环境变量
            for env_key, old_val in [
                ('HTTP_PROXY', old_http_proxy),
                ('HTTPS_PROXY', old_https_proxy),
                ('ALL_PROXY', old_all_proxy),
            ]:
                if old_val is not None:
                    os.environ[env_key] = old_val
                else:
                    os.environ.pop(env_key, None)

    def installPlugin(
        self,
        pluginId: str,
        pluginInfo: dict,
        logCallback: Optional[Callable[[str, str], None]] = None
    ) -> bool:
        '''
        使用纯 Python (Dulwich) 执行 Git 协议克隆与完整性校验
        '''
        self.lastError = None
        def log(msg: str, level: str = 'INFO'):
            if logCallback:
                logCallback(msg, level)

        repoUrl = pluginInfo.get('repoUrl', '').rstrip('/')
        commitHash = pluginInfo.get('commitHash', '')
        entryFile = pluginInfo.get('entryFile', '')

        log(f'准备安装插件 [{pluginId}]', level='INFO')
        log(f'  ├─ 目标仓库: {repoUrl}', level='INFO')
        log(f'  ├─ 指定 Commit: {commitHash}', level='INFO')
        log(f'  └─ 入口文件: {entryFile}', level='INFO')

        if not repoUrl or not commitHash or not entryFile:
            errMsg = f'插件 [{pluginId}] 元数据不完整，缺少 repoUrl、commitHash 或 entryFile'
            log(errMsg, level='ERROR')
            self.lastError = errMsg
            return False

        gitUrl = repoUrl if repoUrl.endswith('.git') else f'{repoUrl}.git'
        targetPluginDir = os.path.join(constants.Path.extensions, pluginId)

        proxyUrl = self._getProxyUrl()
        if proxyUrl:
            log(f'使用代理克隆仓库: {proxyUrl}', level='INFO')

        startTime = time.time()
        try:
            # 1. 清理本地同名插件目录
            if os.path.exists(targetPluginDir):
                log(f'发现同名本地目录，清理中: {targetPluginDir}', level='INFO')
                shutil.rmtree(targetPluginDir, ignore_errors=True)

            # 2. 执行 Git 克隆
            log(f'开始从远程仓库 [{gitUrl}] 克隆数据对象...', level='INFO')
            with self._temporaryOpener():
                repo = porcelain.clone(
                    source=gitUrl,
                    target=targetPluginDir,
                    checkout=False
                )
            cloneTime = time.time() - startTime
            log(f'Git 仓库对象树下载完毕 (耗时 {cloneTime:.2f}s)', level='INFO')

            # 3. 校验并切换提交
            commitHashBytes = commitHash.encode('utf-8')
            if commitHashBytes not in repo:
                raise ValueError(f'仓库中未找到目标 CommitHash: [{commitHash}]')

            porcelain.reset(repo, mode='hard', treeish=commitHashBytes)
            log(f'工作区已硬重置至 Commit: {commitHash[:7]}', level='INFO')

            # 4. 入口文件是否存在
            entryPath = os.path.join(targetPluginDir, entryFile)
            if not os.path.exists(entryPath):
                raise ValueError(f'插件完整性校验失败: 目标路径未发现入口文件 [{entryFile}]')
            log(f'入口文件校验通过: {entryFile}', level='INFO')

            # 5. 清理 .git 目录
            gitDir = os.path.join(targetPluginDir, '.git')
            if os.path.exists(gitDir):
                shutil.rmtree(gitDir, ignore_errors=True)
                log('已自动清理内部 .git 隐藏版本库文件以节省空间的体积', level='INFO')

            metaPath = os.path.join(targetPluginDir, 'install_meta.json')
            try:
                with open(metaPath, 'w', encoding='utf-8') as f:
                    json.dump({
                        "pluginId": pluginId,
                        "commitHash": commitHash,
                        "installedAt": time.strftime('%Y-%m-%d %H:%M:%S')
                    }, f, ensure_ascii=False, indent=4)
                log('已成功写入插件安装元数据 install_meta.json', level='INFO')
            except Exception as e:
                log(f'写入安装元数据失败 (非致命警告): {e}', level='WARNING')

            totalTime = time.time() - startTime
            log(f'插件 [{pluginId}] 安装成功！总用时: {totalTime:.2f}s', level='SUCCESS')
            return True

        except Exception as e:
            totalTime = time.time() - startTime
            errMsg = f'安装插件 [{pluginId}] 失败 (耗时 {totalTime:.2f}s): {e}'
            log(errMsg, level='ERROR')
            self.lastError = errMsg
            if os.path.exists(targetPluginDir):
                shutil.rmtree(targetPluginDir, ignore_errors=True)
            return False