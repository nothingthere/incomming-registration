#!/usr/bin/python3
# restore.py
# Author: Claudio <3261958605@qq.com>
# Created: 2018-04-07 21:14:01
# Code:
'''

'''
import json
import os.path

FILE = 'configuration.json'


def load(tk_folder, tk_excel_file):
    if not os.path.exists(FILE):
        with open(FILE, 'w') as _fp:
            pass
        return

    config = dict()

    try:
        with open(FILE, 'r') as fp:
            config = json.load(fp)
    except:
        pass

    folder = config.get('folder')
    excel_file = config.get('excel_file')

    if config:
        if folder:
            tk_folder.set(folder)

        if excel_file:
            tk_excel_file.set(excel_file)


def dump(folder, excel_file):
    if not os.path.exists(FILE):
        with open(FILE, 'w') as _fp:
            pass

    config = dict()

    config['folder'] = folder
    config['excel_file'] = excel_file

    try:
        with open(FILE, 'w') as fp:
            json.dump(config, fp, ensure_ascii=False)
    except:
        pass
