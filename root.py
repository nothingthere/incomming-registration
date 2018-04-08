#!/usr/bin/python3
# root.py
# Author: Claudio <3261958605@qq.com>
# Created: 2018-04-07 16:51:32
# Code:
'''
图形界面入口
'''

import about_frame
import operation_frame
import os.path
import sys
import tkinter as tk

ico = 'sdgs.ico'

root = tk.Tk()

root.title('乐自高速长山站收文登记')
root.geometry('+10+10')
root.resizable(False, False)
if os.path.exists(ico) and 'win' in sys.platform:
    root.wm_iconbitmap(ico)

about_frame.Frame(root).grid(column=0, row=0, sticky='w', padx=(10, 10))
operation_frame.Frame(root).grid(
    column=1, row=0, rowspan=3, sticky='e', padx=(10, 10), pady=(0, 0))


#
# 运行
#


root.mainloop()
