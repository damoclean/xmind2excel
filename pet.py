#!/usr/bin/python3
import tkinter
import re
import tkmacosx
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
from common.xmind2execl import Xmind2Excel


class MainUI(object):

    def __init__(self, title="sulink", geometrysize="350x250", geometry="+500+350"):
        #self.top = tkinter.Tk()  # 生成主窗口
        self.top = tkinter.Tk()
        self.top.eval('package forget Tcl')
        self.top.eval('package provide Tcl 8.6')
        self.top.eval('package require Tcl')  # 生成主窗口

        self.top.title(title)  # 设置窗口的标题
        self.top.geometry(geometrysize)  # 设置窗口的大小
        self.top.geometry(geometry)  # 设置窗口出现的位置
        self.top.resizable(0, 0)  # 将窗口大小设置为不可变
        self.path = tkinter.StringVar()  # 生成一个StringVar 对象，来保存下面输入框中的内容
        self.person = tkinter.StringVar()
        self.version = tkinter.StringVar()
        #打开是在所有应用前面
        self.top.lift()
        self.top.attributes('-topmost', True)
        self.top.after_idle(self.top.attributes, '-topmost', False)

        # 调用自己写的create_widgets()方法
        self.create_widgets()

    def get_value(self):
        """获取文本框中数据，并调用XmindToXsl类"""

        xmindPath = self.path.get()
        operator = self.person.get()
        ver = self.version.get()
        #print(f"地址：{xmindPath}，测试人员：{operator}，测试版本：{ver}")
        regvalue = '.*\.xmind$'
        xmind_reg = re.match(regvalue, xmindPath)
        if xmind_reg:
            # xmind转换成xls
            self.Xmind2Excel = Xmind2Excel()
            self.Xmind2Excel.xmind2excel(xmindPath, operator)
        else:
            messagebox.showinfo(title='Tips', message='Please select the correct XMIND file')

    def select_path(self):
        """选择要转换成excel的xmind地址"""

        path_ = askopenfilename()
        self.path.set(path_)

    def create_widgets(self):
        """创建窗口中的各种元素"""

        # 文件的路径
        first_label = tkinter.Label(self.top, text='  Path：')  # 生成一个标签
        first_label.grid(row=0, column=0)  # 使用grid布局，标签显示在第一行，第一列

        first_entry = tkinter.Entry(self.top, textvariable=self.path)  # 生成一个文本框，内容保存在上面变量中
        first_entry.grid(row=0, column=1)  # 使用grid布局，文本框显示在第一行，第二列
        way_button = tkmacosx.Button(self.top, text="Select", command=self.select_path)
        way_button.grid(row=0, column=2)  # 使用grid布局，按钮显示在第一行，第三列

        # 测试人员
        second_label = tkinter.Label(self.top, text="Owner：")
        second_label.grid(row=1, column=0)
        second_entry = tkinter.Entry(self.top, textvariable=self.person)
        second_entry.grid(row=1, column=1)

        # 版本
        #third_label = tkinter.Label(self.top, text="测试版本：")
        #third_label.grid(row=2, column=0)
        #third_entry = tkinter.Entry(self.top, textvariable=self.version)
        #third_entry.grid(row=2, column=1)

        # 提交按钮
        #f_btn = tkinter.Frame(self.top, bg='red')  # 设置一个frame框架，并设置背景颜色为红色
        # 支持 m1
        f_btn = tkmacosx.SFrame(self.top)  # 设置一个frame框架，并设置背景颜色为红色
        f_btn.place(x=0, y=208, width=460, height=45)  # 设置框架的大小，及在top窗口显示位置
        # submit_button = tkinter.Button(f_btn, text="Submit", command=self.get_value, width=36, height=3,bg='green')  # 设置按钮的文字，调用方法，大小，颜色，显示框架
        submit_button = tkmacosx.Button(f_btn, text="Submit", command=self.get_value, width=350)  # 设置按钮的文字，调用方法，大小，颜色，显示框架
        submit_button.grid(row=0, column=1)  # 使用grid布局，按钮显示在第一行，第一列

        # 进入消息循环（必需组件）
        self.top.mainloop()


if __name__ == "__main__":
    mu = MainUI(title="xmind2excel")