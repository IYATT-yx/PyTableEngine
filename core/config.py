'''
file: config.py
description: 配置管理器
author: IYATT-yx
copyright:  Copyright (c) 2026 IYATT-yx.
            Licensed under the MIT License. See LICENSE file in the project root for full license information.
'''
import configparser
import os

from core import constants

class Config:
    '''统一配置管理器'''

    def __init__(self, configPath=constants.Path.config):
        self.configPath = configPath
        self.config = configparser.ConfigParser()
        self.loadConfig()

    def loadConfig(self):
        '''加载配置文件，不存在则自动生成初始模板'''
        if os.path.exists(self.configPath):
            self.config.read(self.configPath, encoding='utf-8')
        else:
            self.generateDefaultConfig()

    def generateDefaultConfig(self):
        '''自动生成默认 config.ini'''
        self.config['pip'] = {
            'indexUrl': '',
            'proxy': ''
        }
        self.config['disabled_plugins'] = {}
        self.save()

    def save(self):
        '''持久化保存至 config.ini'''
        with open(self.configPath, 'w', encoding='utf-8') as f:
            self.config.write(f)

    def getCleanOption(self, section, option):
        '''安全获取配置项：自动去除首尾空格，若为空或全空格则返回 None'''
        if self.config.has_section(section) and option in self.config[section]:
            val = self.config[section][option].strip()
            return val if val else None
        return None

    def applyGlobalProxy(self):
        '''将 config.ini 中的代理设置全局注入当前进程环境变量'''
        proxy_url = self.getCleanOption('pip', 'proxy')
        if proxy_url:
            os.environ['HTTP_PROXY'] = proxy_url
            os.environ['HTTPS_PROXY'] = proxy_url
            os.environ['http_proxy'] = proxy_url
            os.environ['https_proxy'] = proxy_url
        else:
            # 若未配置或为空，清理可能继承自系统的代理环境变量，防止干扰
            for env_key in ['HTTP_PROXY', 'HTTPS_PROXY', 'http_proxy', 'https_proxy']:
                os.environ.pop(env_key, None)

    # ---- 禁用插件记录相关逻辑 ----
    def getDisabledPlugins(self):
        if self.config.has_section('disabled_plugins'):
            return dict(self.config.items('disabled_plugins'))
        return {}

    def setPluginDisabled(self, plugin_key, is_disabled):
        if not self.config.has_section('disabled_plugins'):
            self.config.add_section('disabled_plugins')

        if is_disabled:
            self.config.set('disabled_plugins', plugin_key, 'true')
        else:
            self.config.remove_option('disabled_plugins', plugin_key)
        self.save()