'''
file: config.py
description: 配置管理器
author: IYATT-yx
copyright: Copyright (c) 2026 IYATT-yx.
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
        parentDir = os.path.dirname(self.configPath)
        os.makedirs(parentDir, exist_ok=True)

        self.config['pip'] = {
            'pypi': constants.Site.pypi,
            'proxy': ''
        }
        self.config['market'] = {
            'repo_index': constants.Site.repositoryIndex,
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

    def getDisabledPlugins(self):
        '''获取被禁用的插件字典列表'''
        if self.config.has_section('disabled_plugins'):
            return dict(self.config.items('disabled_plugins'))
        return {}

    def setDisabledPlugins(self, pluginId:str, isDisabled:bool):
        '''设置插件禁用状态'''
        if not self.config.has_section('disabled_plugins'):
            self.config.add_section('disabled_plugins')

        if isDisabled:
            self.config.set('disabled_plugins', pluginId, 'true')
        else:
            self.config.remove_option('disabled_plugins', pluginId)

        self.save()