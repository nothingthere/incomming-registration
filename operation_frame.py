#!/usr/bin/python3
# operation_frame.py
# Author: Claudio <3261958605@qq.com>
# Created: 2018-04-07 17:00:18
# Code:
'''

'''
import operation
import restore
import tkinter as tk
import tkinter.filedialog as filedialog
import tkinter.messagebox as messeagebox
import validate
from public import FONT


class Frame(tk.Frame):
    def __init__(self, master):
        super(Frame, self).__init__(master)
        self.master = master
        self.font = FONT
        self.folder_icon = tk.PhotoImage(file='folder.png')
        self.excel_icon = tk.PhotoImage(file='excel.png')

        self.folder = tk.StringVar()
        self.excel_file = tk.StringVar()
        self.progress_str = tk.StringVar()

        self.folder_widget = None
        self.folder_button_widget = None
        self.excel_file_widget = None
        self.excel_file_button_widget = None
        self.submit_button_widget = None

        self.add_dir()
        self.add_excel_file()
        self.add_commit_button()
        self.add_progress()

        self.interactive_widgets = [
            self.folder_widget,
            self.folder_button_widget,
            self.excel_file_widget,
            self.excel_file_button_widget,
            self.submit_button_widget,
        ]

        # 加载上次配置
        restore.load(self.folder, self.excel_file)

    def ask_dir(self):
        dirname = filedialog.askdirectory(title='请选择待处理照片所在文件夹')
        self.folder.set(dirname)

    def add_dir(self):
        tk.Label(self,
                 text='待处理文件所在文件夹：',
                 font=self.font,
                 ).grid(row=0, column=0, sticky='w', pady=(10, 0))

        self.folder_widget = tk.Entry(self,
                                      textvariable=self.folder,
                                      width=33,
                                      font=self.font,
                                      )
        self.folder_widget.grid(row=1, column=0, sticky='w')

        self.folder_button_widget = tk.Button(self,
                                              image=self.folder_icon,
                                              command=self.ask_dir
                                              )
        self.folder_button_widget.grid(
            row=1, column=1, sticky='e', padx=(0, 0))

    def ask_excel_file(self):
        name = filedialog.askopenfilename(title='请选择需操作的excel文件',
                                          filetypes=(
                                              ('xlsx文件', '.xlsx'),
                                              ('all files', '*.*')),
                                          )
        self.excel_file.set(name)

    def add_excel_file(self):
        tk.Label(self,
                 text='需操作的Excel文件：',
                 font=self.font,
                 ).grid(row=4, column=0, sticky='w', pady=(10, 0))

        self.excel_file_widget = tk.Entry(self,
                                          textvariable=self.excel_file,
                                          width=33,
                                          font=self.font,
                                          )
        self.excel_file_widget.grid(row=5, column=0, sticky='w')

        self.excel_file_button_widget = tk.Button(self,
                                                  image=self.excel_icon,
                                                  command=self.ask_excel_file
                                                  )

        self.excel_file_button_widget.grid(
            row=5, column=1, sticky='e', padx=(5, 0))

    def add_progress(self):
        tk.Label(self,
                 textvariable=self.progress_str,
                 font=(FONT['family'], FONT['size'], 'italic'),
                 ).grid(row=6, column=0, columnspan=2, pady=(10, 10))

    def update_progress(self, message=''):
        self.progress_str.set(message)
        self.master.update_idletasks()

    def add_commit_button(self):
        self.submit_button_widget = tk.Button(self,
                                              text='开始',
                                              font=(
                                                  FONT['family'], FONT['size'] + 2, 'bold'),
                                              command=self.run,
                                              )
        self.submit_button_widget.grid(
            row=7, column=0, columnspan=2, pady=(0, 20))

    def disable_interactive(self):
        for widget in self.interactive_widgets:
            widget.configure(state='disabled')

    def enable_interactive(self):
        for widget in self.interactive_widgets:
            widget.configure(state='normal')

    def validate(self):
        folder = validate.folder(self.folder, self.folder_widget)
        if False is folder:
            return False

        excel_file = validate.excel_file(
            self.excel_file, self.excel_file_widget)
        if False is excel_file:
            return False

        return {'folder': folder, 'excel_file': excel_file}

    def run(self):
        infos = self.validate()
        if not infos:
            return
        print(infos)

        folder = infos.get('folder')
        excel_file = infos.get('excel_file')

        #
        # 为退出程序添加事件，存储配置
        #
        def on_closing():
            restore.dump(folder, excel_file)
            self.master.destroy()

        self.master.protocol('WM_DELETE_WINDOW', on_closing)

        #
        # 执行操作
        #
        self.update_progress('执行中')
        self.disable_interactive()
        operation.dump_files(folder, excel_file)
        self.enable_interactive()
        self.update_progress()
        messeagebox.showinfo(title='成功', message='操作成功')
