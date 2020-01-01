#  coding=utf-8
import os # 设置虚拟路径,以供将py文件所在路径的类型文件全部导入
from os import listdir
import numpy as np
import pandas as pd
import win32com.client
from docx import Document  # docx用来操作word.docx,document用来新建空白文档
from docx.enum.text import WD_LINE_SPACING  # 设置段落的行间距
from docx.oxml.ns import qn  # 设置中文字体类型
from docx.shared import Cm, Pt  # 设置字体大小Pt磅值,设置行高Cm厘米
import time
import re
import pythoncom



class RemoteWord:
    def __init__(self, filename=None):
        self.xlApp = win32com.client.DispatchEx('Word.Application')
        self.xlApp.Visible = 0
        self.xlApp.DisplayAlerts = 0  # 后台运行，不显示，不警告
        if filename:
            self.filename = filename
            if os.path.exists(self.filename):
                self.doc = self.xlApp.Documents.Open(filename)
            else:
                self.doc = self.xlApp.Documents.Add()  # 创建新的文档
                self.doc.SaveAs(filename)
        else:
            self.doc = self.xlApp.Documents.Add()
            self.filename = ''

    def add_doc_end(self, string):
        '''在文档末尾添加内容'''
        rangee = self.doc.Range()
        rangee.InsertAfter('\n'+string)

    def add_doc_start(self, string):
        '''在文档开头添加内容'''
        rangee = self.doc.Range(0, 0)
        rangee.InsertBefore(string+'\n')

    def insert_doc(self, insertPos, string):
        '''在文档insertPos位置添加内容'''
        rangee = self.doc.Range(0, insertPos)
        if (insertPos == 0):
            rangee.InsertAfter(string)
        else:
            rangee.InsertAfter('\n'+string)

    def replace_doc(self, string, new_string):
        '''替换文字'''
        self.xlApp.Selection.Find.ClearFormatting()
        self.xlApp.Selection.Find.Replacement.ClearFormatting()
        self.xlApp.Selection.Find.Execute(string, False, False, False, False, False, True, 1, True, new_string, 2)

    def save(self):
        '''保存文档'''
        self.doc.Save()

    def save_as(self, filename):
        '''文档另存为'''
        self.doc.SaveAs(filename)

    def w_to_pdf(self, out_path:str):
        '''另存为pdf'''
        self.doc.SaveAs2(out_path, 17)
        self.xlApp.Documents.Close()
        self.xlApp.Quit()

    def close(self):
        '''保存文件、关闭文件'''
        self.save()
        self.xlApp.Documents.Close()
        self.xlApp.Quit()

    def PageSetup_Orientation(self, Find_string) -> int:
        '''查找文本内容所在章节或者页面并返回章节'''
        self.xlApp.Selection.Find.ClearFormatting()
        #  查找字符串查找完代表选中
        self.xlApp.Selection.Find.Execute(Find_string)
        # 可以返回查找文件所在章节
        page_number = self.xlApp.Selection.Information(2)
        return page_number
        # 把选中的内容所在章节设定为纵向
        # doc_s5 = self.xlApp.Selection.Sections[page_number]
        # PageSetup.PaperSize 
        # doc_s5.PageSetup.Orientation = 0
        # self.xlApp.Selection.Sections[4].PageSetup.Orientation = 0


def word_to_p(word_path, out_path):
    '''把一个word设定页面为A4 并把7000)所在章节设定为竖向'''
    pythoncom.CoInitialize()
    docxlist = [fn for fn in listdir(word_path) if fn.endswith('.docx')]
    word_path = os.path.join(word_path, docxlist[0])
    document = Document(word_path)  # python-docx
    dsec_len = len(document.sections)  # word共有多少章节
    # document.close()
    dpos_n_l = set()
    listfind = ['7000)', '一般状态观察个体数据', 'FACSCalibur']
    for o in listfind:
        doc = RemoteWord(word_path)  # 初始化一个doc对象 win32aip
        dpos_n_l.add(doc.PageSetup_Orientation(o))
        doc.close()
    print(dpos_n_l)

    '''设定纸张依据节返回值'''

    # 设定没张纸的大小所有第单数节为横向
    for i in range(2, dsec_len, 2):
        section = document.sections[i]
        section.page_height, section.page_width = Cm(21.0), Cm(29.7)

    # section.orientation = WD_ORIENT.LANDSCAPE  # 设定章节横纵向
    for i in dpos_n_l:  # 循环一个找到的节 集合设定页面方向
        try:
            section = document.sections[i-1]
            new_width, new_height = section.page_height, section.page_width
            section.page_width, section.page_height  = new_width, new_height
        except:
            print('转换方向错误')
    # 设定没张纸的大小所有第双数数节为纵向
    for i in range(1, dsec_len, 2):
        section = document.sections[i]
        section.page_height, section.page_width = Cm(29.7), Cm(21.0)

    # 第1也就是0节修改为竖向--Find不到的内容会返回1所以最后在设定第一张
    section = document.sections[0]
    section.page_height, section.page_width = Cm(29.7), Cm(21.0)
    # 保存word到pdf
    document.save(word_path)
    doc2 = RemoteWord(word_path)
    out_path = os.path.join(out_path, '{}.pdf'.format(docxlist[0][:-5]))
    doc2.w_to_pdf(out_path)


if __name__ == "__main__":
    time_start = time.clock()
    word_path = r'C:\Users\admin\Desktop\Django\uploadFiles-master\media\input'
    out_path = r'C:\Users\admin\Desktop\Django\uploadFiles-master\media\input\A2018033-T015-01-1.pdf'
    word_to_p(word_path, out_path)
    print('恭喜你搞定了总耗时', time.clock() - time_start)
