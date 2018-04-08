#!/usr/bin/python3
# validate.py
# Author: Claudio <3261958605@qq.com>
# Created: 2018-04-07 20:57:33
# Code:
'''
验证
'''
import os
from tkinter import messagebox


def show_error(message, widget=None):
    '''
    操作错误提示

    @widget：如果出现操作错误时，需focus的元素
    '''
    messagebox.showerror(title='错误', message=message)
    if widget:
        widget.focus()


def excel_file(tk_str, focus):
    file_path = tk_str.get().strip()

    if '' == file_path:
        show_error('请选择需操作的excel文件', focus)
        return False

    if not os.path.exists(file_path):
        show_error('{}：不存在此文件'.format(file_path))
        return False

    if '.xlsx' != os.path.splitext(file_path)[1]:
        show_error('{}：不为xlsx文件'.format(file_path))
        return False

    return file_path


def folder(tk_str, focus):
    file_path = tk_str.get().strip()

    if '' == file_path:
        show_error('请选择需操作的excel文件', focus)
        return False

    if not os.path.exists(file_path):
        show_error('{}：不存在此文件'.format(file_path))
        return False

    if not os.path.isdir(file_path):
        show_error('{}：不为文件夹'.format(file_path))
        return False

    return file_path
