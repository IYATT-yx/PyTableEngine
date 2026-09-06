'''
file: logger.py
description: 日志管理封装
author: IYATT-yx
copyright:  Copyright (c) 2026 IYATT-yx.
            Licensed under the MIT License. See LICENSE file in the project root for full license information.
'''
import logging
import os
from logging.handlers import RotatingFileHandler

from core import constants

class CustomLogger(logging.Logger):
    def makeRecord(self, name, level, fn, lno, msg, args, exc_info, func=None, extra=None, sinfo=None):
        if extra:
            if 'filename' in extra:
                fn = extra.pop('filename')
            if 'funcName' in extra:
                func = extra.pop('funcName')
            if 'lineno' in extra:
                lno = extra.pop('lineno')
                
        return super().makeRecord(name, level, fn, lno, msg, args, exc_info, func, extra, sinfo)

class AppLogger:
    '''日志管理封装（包含调用源头追溯与日志轮转）'''

    @staticmethod
    def setupLogger(logFileName=constants.Path.log, maxBytes=5*1024*1024, backupCount=3):
        logFilePath = os.path.join(constants.Path.appDir, logFileName)
        os.makedirs(os.path.dirname(logFilePath), exist_ok=True)

        logging.setLoggerClass(CustomLogger)

        logger = logging.getLogger('PyTabEngine')
        logger.setLevel(logging.INFO)
        logger.handlers.clear()  

        log_format = logging.Formatter(
            '[%(asctime)s] [%(levelname)s] [%(plugId)s] [%(filename)s -> %(funcName)s():%(lineno)d] - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )

        file_handler = RotatingFileHandler(
            logFilePath, maxBytes=maxBytes, backupCount=backupCount, encoding='utf-8'
        )
        file_handler.setFormatter(log_format)
        logger.addHandler(file_handler)

        return logger