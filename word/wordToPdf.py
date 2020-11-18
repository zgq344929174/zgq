# -*- coding: utf-8 -*-

import sys, os
# 调用com组件包
import comtypes.client

'''
word批量转换成Pdf
'''

def init_word():
    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = 1
    return word

# 第二步:找到该路径下的所有doc(x)文件,并将其路径添加到cwd
def convert_files_in_folder(word, folder):
    # 将当前所有文件及文件夹添加进列表
    files = os.listdir(folder)
    print('files:',files)
    # 将所有以.doc(x)结尾的文件加入cwd path
    pptfiles = [f for f in files if f.endswith((".doc", ".docx"))]
    for pptfile in pptfiles:
        # 加入判断,如果当前转换成的pdf已存在,就跳过不添加
        if pptfile+'.pdf' in files :
            break
        # 加入cwd环境
        fullpath = os.path.join(cwd, pptfile)
        # print('fullpath====', fullpath)
        outputFileName = ''
        if fullpath.endswith(".doc"):
            outputFileName = fullpath.replace(".doc","")
        else:
            outputFileName = fullpath.replace(".docx","")
        outputFileName = outputFileName.replace("2020.8.24-法律法规i汇编-外网(1)", "2020.8.24-法律法规i汇编-外网-pdf")
        # print(outputFileName)
        ppt_to_pdf(word, fullpath, outputFileName)

#第三步:将cwd路径下转换成pdf格式
def ppt_to_pdf(word, inputFileName, outputFileName, formatType = 17):
    # 切片取后缀是否为pdf
    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName + ".pdf"
        print('outputFileName======',outputFileName)
    # 调用接口进行转换
    print(inputFileName)
    deck = word.Documents.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType) # formatType = 17 for word to pdf
    deck.Close()

if __name__ == "__main__":
    # 创建Word应用
    word = init_word()
    # 得到当前路径
    # cwd = os.getcwd()
    ##  word所在的目录
    cwd = "D:\\Users\\zgq\\工作\\03-需求分析\\303-产品\\003-执法项目\\2020.8.24-法律法规i汇编-外网(1)\\通用型"
    # 打印当前路径
    # print(cwd)
    # 调用Word进行转换cwd path下的doc(x)格式
    convert_files_in_folder(word, cwd)
    # 转换结束后关闭
    word.Quit()