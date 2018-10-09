#!/usr/bin/env python
# -*- coding:utf-8 -*-
# author zhibin time:2018/10/4
import xlrd
import tkinter


class DanCiBen1:
    WordList = []
    ChineseList = []
    WordNature = []
    jsq = 0

    def load_word(self, j):
        wb = xlrd.open_workbook("3000coreword.xls")  # 打开文件
        sheet = wb.sheet_by_index(0)  # 指定读取第一个表格的信息
        t = (j - 1) * 15 + 1
        for i in range(15):
            r = i + t  # r为取数列
            self.WordList.append(sheet.cell(r, 2).value)  # 获取第r行，第3列的英文单词
            self.ChineseList.append(sheet.cell(r, 3).value)  # 获取第r行，第4列的中文意思

    def loadwordlist2(self, j):
        wb2 = xlrd.open_workbook("text50.xlsx")  # 打开文件
        sheet2 = wb2.sheet_by_index(0)  # 指定读取第一个表格的信息
        r = 1  # 从第二行开始读取
        while sheet2.cell(r, 0).value != j:  # 循环到(r,2)数值 与 设定数相等为止
            r += 1
        while sheet2.cell(r, 0).value == j:  # 循环到(r,2)数值 与 设定数值不等为止
            self.WordList.append(sheet2.cell(r, 1).value)  # 获取第r行，第3列的英文单词
            self.ChineseList.append(sheet2.cell(r, 3).value)
            self.WordNature.append(sheet2.cell(r, 2).value)
            r += 1

    def clear(self):
        self.WordList = []
        self.ChineseList = []
        var.set(" ")
        var2.set(" ")
        var3.set(" ")
        var4.set(" ")

    def auto_appear(self):
        N1 = [1, 1, 2, 2, 1, 2, 1, 2, 3, 3, 1, 2, 3, 1, 2, 3, 2, 3, 4, 4, 2, 3, 4, 1, 2, 3, 4, 2, 3, 4, 5, 5, 1, 3, 4,
              5, 2, 3, 4, 5, 1, 3, 4, 5, 6, 6, 2, 4, 6, 1, 3, 5, 6, 2, 4, 6, 7, 7, 3, 5, 7, 2, 4, 6, 7, 3, 5, 7, 1, 2,
              3, 8, 8, 4, 6, 8, 3, 5, 7, 8, 4, 6, 8, 9, 9, 5, 7, 9, 4, 6, 8, 9, 3, 5, 7, 9, 10, 10, 6, 8, 10, 5, 7, 9,
              10, 4, 6, 8, 10, 11, 11, 5, 7, 9, 11, 6, 8, 10, 11, 7, 9, 11, 12, 12, 6, 8, 10, 12, 7, 9, 11, 12, 6, 8,
              10, 12, 13, 13, 9, 11, 13, 1, 2, 3, 8, 10, 12, 13, 9, 11, 13, 14, 14, 8, 10, 12, 14, 4, 5, 6, 9, 11, 13,
              14, 12, 8, 10, 12, 15, 15, 11, 13, 15, 4, 10, 12, 14, 15, 11, 13, 15, 1, 3, 5, 7, 9, 14, 11, 13, 15, 2, 4,
              14, 6, 8, 15, 10, 12, 14, 1, 3, 5, 7, 2, 4, 6, 8, 9, 11, 13, 15, 10, 12, 14, 1, 1]
        word = self.WordList[N1[self.jsq] - 1]  # 读取第i个单词
        chinese = self.ChineseList[N1[self.jsq] - 1]  # 读取第i个单词的中文意思
        try:
            if N1[self.jsq] == N1[self.jsq + 1]:
                var.set(word)
                var2.set(chinese)  # 如果该单词第一次出现，在lb2上显示中文单词意思
                BDC.after(4000, self.auto_appear, )  # 第一次出现的单词停留4s
            else:
                var.set(word)
                var2.set(" ")
                BDC.after(2000, self.auto_appear, )  # 再一次出现的单词停留2.0S

        except IndexError:  # 当程序执行完成的时候，会进入益处异常，跳到该程序段
            var.set(" ")
            var2.set(" ")
            ks1 = '{0}\n{1}\n{2}\n{3}\n{4}\n{5}\n{6}\n{7}\n{8}\n{9}\n{10}\n{11}\n{12}\n{13}\n{14}'.format(L1[0], L1[1],
                                                                                                          L1[2], L1[3],
                                                                                                          L1[4], L1[5],
                                                                                                          L1[6], L1[7],
                                                                                                          L1[8], L1[9],
                                                                                                          L1[10],
                                                                                                          L1[11],
                                                                                                          L1[12],
                                                                                                          L1[13],
                                                                                                          L1[14], )
            ks2 = '{0}\n{1}\n{2}\n{3}\n{4}\n{5}\n{6}\n{7}\n{8}\n{9}\n{10}\n{11}\n{12}\n{13}\n{14}'.format(L2[0], L2[1],
                                                                                                          L2[2], L2[3],
                                                                                                          L2[4], L2[5],
                                                                                                          L2[6], L2[7],
                                                                                                          L2[8], L2[9],
                                                                                                          L2[10],
                                                                                                          L2[11],
                                                                                                          L2[12],
                                                                                                          L2[13],
                                                                                                          L2[14], )
            var3.set(ks1)
            var4.set(ks2)
            self.jsq = 0
            BDC.after(40000, self.clear)

    def fuxi(self):
        try:
            word = self.WordList[self.jsq]
            var.set(word)
            var2.set("")  # 清空中文意思
            self.jsq += 1
        except IndexError:
            self.jsq = 0
            self.fuxi()

    def toukan(self):
        chinese = self.ChineseList[self.jsq - 1]  # 读取第i个单词的中文意思
        nature = self.WordNature[self.jsq - 1]
        kkk = nature + " " + chinese
        var2.set(kkk)
        self.jsq -= 1
        BDC.after(800, self.fuxi)


dcb = DanCiBen1()  # 预先准备工作 创建一个实例


# 3个调用的中转程序
def run1():  # 调用单词组1
    a = int(inp1.get())
    dcb.load_word(a)


def run2():  # 调用单词滚动
    dcb.auto_appear()


def run3(i=0):  # 调用下一个，默认0是为了防止出错，无实际意义
    dcb.fuxi()


def run6(i=0):  # 调用上一个
    dcb.jsq -= 2
    dcb.fuxi()


def run4(i=0):  # 调用偷看功能
    dcb.toukan()


def run5():  # 调用50text的单词组2
    a = int(inp1.get())
    dcb.loadwordlist2(a)


def run7():  # 调用清空功能
    dcb.clear()


# 构建窗口用的按钮、显示器
BDC = tkinter.Tk()
BDC.title("极速闪回法背单词")
BDC.geometry("2400x1200+100+100")
BDC.bind("<space>", run3)
BDC.bind("<Up>", run4)
BDC.bind("<Left>", run6)
# BDC.bind("<Key-f>", run7)

# 结构框和事件绑定
frame = tkinter.Frame(height=1200, width=2400, bg="Wheat")  # 界面背景色：麦黄色
frame.pack(expand="YES", )

# 输出变量，用于将输出的数据传给输出框
var = tkinter.StringVar()
var2 = tkinter.StringVar()
var3 = tkinter.StringVar()
var4 = tkinter.StringVar()

# 输入框，输入要学习的单元
inp1 = tkinter.Entry(BDC)
inp1.place(relx=0.79, y=1050, height=50, width=100)  # 指定输入框的位置

# 按键，负责调用功能。
b1 = tkinter.Button(BDC, text="载入单词组\n共1-219组", command=run1, width=15, height=2, bd=4)
b1.place(relx=0.85, y=1050, relheight=0.05)  # 指定按钮框的位置
b6 = tkinter.Button(BDC, text="载入单词组\n共1-50篇文章", command=run5, width=15, height=2, bd=4)
b6.place(relx=0.85, y=1110, relheight=0.05)  # 指定按钮框的位置
b2 = tkinter.Button(BDC, text="自动滚动模式", command=run2, width=20, height=3, bd=4, )
b2.place(relx=0.92, y=1050, relheight=0.1)  # 指定按钮框的位置
b3 = tkinter.Button(BDC, text="点动复习模式\n下一个（→）\n上一个（←）", command=run3, width=20, height=2, bd=4, )
b3.place(relx=0.02, y=1050, relheight=0.1)  # 指定按钮框的位置
b4 = tkinter.Button(BDC, text="清空单词列表", command=run7, width=15, height=2, bd=4, )
b4.place(relx=0.79, y=1110, relheight=0.05)
b5 = tkinter.Button(BDC, text="偷看一下\n（↑）", command=run4, width=15, height=2, bd=4)
b5.place(relx=0.02, y=950, relheight=0.05)

# 输出框，分别负责输出 1英文 2中文 3全英文列表 4全中文列表
lb1 = tkinter.Label(BDC, textvariable=var, fg="black", font=("Calibri", 200), bg="Wheat")
lb1.place(x=300, y=200)  # 该位置负责输出英文
lb2 = tkinter.Label(BDC, textvariable=var2, fg="black", font=("黑体", 50), bg="Wheat")
lb2.place(x=600, y=600)  # 该位置负责输出中文
lb3 = tkinter.Label(BDC, textvariable=var3, fg="black", font=("Calibri", 45), bg="Wheat")
lb3.place(x=300, y=0)  # 该位置负责结束的时候输出英文列表
lb4 = tkinter.Label(BDC, textvariable=var4, fg="black", font=("Calibri", 45), bg="Wheat")
lb4.place(x=700, y=0)  # 该位置负责结束的时候输出英文列表

BDC.mainloop()
