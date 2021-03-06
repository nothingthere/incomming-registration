#!/usr/bin/python3
# frames.py
# Author: Claudio <3261958605@qq.com>
# Created: 2018-04-07 16:55:54
# Code:
'''

'''
import tkinter as tk
import webbrowser
from public import FONT


class Frame(tk.Frame):
    def __init__(self, master=None):
        super(Frame, self).__init__(master)

        self.infos = ['Version: 0.1',
                      'Contact: 3261958605@qq.com']
        # self.manual_url = r'https://github.com/nothingthere/pics_handler/blob/master/README.org'
        self.logo = tk.PhotoImage(file='sdgs.png')
        self.font = (FONT['family'], FONT['size'] - 2, 'italic')
        self.add_infos()
        # self.add_manual()
        self.add_logo()

    def add_logo(self):
        tk.Label(self, image=self.logo, compound='right').grid(
            row=0, column=0, sticky='w', pady=(30, 0)
        )

    def add_infos(self):
        for row, text in enumerate(self.infos):
            tk.Label(self,
                     text=text,
                     font=self.font
                     ).grid(row=row + 1, column=0, sticky='w')

    def add_manual(self):
        manual = tk.Label(self,
                          text='使用说明',
                          font=(self.font[0], self.font[1],
                                'italic underline'),
                          fg='blue',
                          cursor='hand2',
                          )
        manual.bind(
            '<Button-1>', lambda e: webbrowser.open_new(self.manual_url))
        manual.grid(row=len(self.infos) + 1, column=0, sticky='w')
