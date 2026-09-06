'''
file: PyTableEngine.py
description: 思能快表引擎
author: IYATT-yx
copyright:  Copyright (c) 2026 IYATT-yx.
            Licensed under the MIT License. See LICENSE file in the project root for full license information.
'''
import multiprocessing
import tkinter as tk
from ui.application import MainWindow

from core import config

if __name__ == '__main__':
    multiprocessing.freeze_support()

    configObj = config.Config()
    configObj.loadConfig()
    configObj.applyGlobalProxy()

    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()