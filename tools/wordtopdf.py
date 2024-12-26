# -*- coding:utf-8 -*-
import os
import os
import comtypes.client
import shutil
import pythoncom
import os
import constants  # 确保你已经正确导入了 constants 模块
import time  # 如果需要的其他模块
from PyQt5.QtWidgets import QMessageBox
from win32com.client import Dispatch, DispatchEx  # 导入pywin32模块的client包下的函数
from win32com.client import constants  #  导入pywin32模块的client包下的保存COM常量的类
from win32com.client import gencache    #  导入pywin32模块的client包下的gencache函数
from PyPDF2 import  PdfFileReader  # 获取页码用
import pythoncom  # 导入封装了OLE自动化API的模块，该模块为pywin32的子模块
from PyPDF2 import PdfReader  # 替代 PdfFileReader
totalPages = 0  # 记录总页数的全局变量
returnlist = []  # 保存文件列表的全局变量

# Word转换为PDF(多个文件)

def getPdfPageNum(pdf_name):
    try:
        # 使用 PdfReader 打开 PDF 文件
        with open(pdf_name, "rb") as file:
            reader = PdfReader(file)
            # 返回 PDF 页数
            return len(reader.pages)
    except Exception as e:
        raise Exception(f"无法获取 PDF 页数: {e}")


def wordtopdf(filelist, targetpath):
    # 创建 Word 应用程序实例
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False  # 设置为不可见，避免显示 Word 窗口

    # 存储转换后的 PDF 文件路径
    pdf_files = []

    try:
        for doc_file in filelist:
            if not doc_file.lower().endswith('.doc') and not doc_file.lower().endswith('.docx'):
                continue  # 忽略非 Word 文件

            # 获取文件名
            filename = os.path.basename(doc_file)
            pdf_filename = os.path.splitext(filename)[0] + '.pdf'
            pdf_filepath = os.path.join(targetpath, pdf_filename)

            # 打开 Word 文件
            doc = word.Documents.Open(doc_file)
            # 转换为 PDF 格式
            doc.SaveAs(pdf_filepath, FileFormat=17)  # 17 是 PDF 格式的常量
            doc.Close()  # 关闭文档

            # 将 PDF 路径添加到列表中
            pdf_files.append(pdf_filepath)

        # 返回转换后的文件路径列表
        return pdf_files

    except Exception as e:
        print(f"转换失败: {e}")
        QMessageBox.warning(None, "错误", f"转换过程中出现问题: {str(e)}", QMessageBox.Ok)
        return -1

    finally:
        # 退出 Word 应用程序
        word.Quit()

# Word转换为PDF并提取页码
def wordtopdf1(filelist):
    totalPages = 0  # 不需要 global 关键字
    valueList = []

    try:
        # 初始化COM库
        print("初始化 COM 库...")
        pythoncom.CoInitializeEx(0)
        print("COM 库初始化成功")

        # 确保加载 Word 的 COM 模块
        print("确保加载 Word COM 模块...")
        gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 6)  # 确保与缓存文件版本一致
        word = DispatchEx('Word.Application')  # 使用 DispatchEx 创建新的进程
        print("Word COM 模块加载成功")

        # 启动 Word 应用
        print("启动 Word 应用...")
        word.Visible = False  # 如果不希望显示 Word 窗口，可以将其设置为 False
        print("Word 应用已成功启动")

        for fullfilename in filelist:
            print(f"处理文件: {fullfilename}")
            temp = fullfilename.split('\\')
            path = '\\'.join(temp[:-1])  # 获取上级目录（去掉文件名部分）
            filename = temp[-1]  # 获取文件名
            print("使得1   " + path)
            print("使得2  " + filename)

            # 获取文件的绝对路径
            doc = os.path.abspath(fullfilename)
            print("使得3  " + doc)

            # 确定目标目录（即上级目录）
            target_directory = os.path.join(path, "backup")  # 在上级目录创建一个 'backup' 文件夹

            # 如果目标目录不存在，则创建该目录
            if not os.path.exists(target_directory):
                os.makedirs(target_directory)

            # 复制文件到目标目录
            target_path = os.path.join(target_directory, filename)
            shutil.copy(doc, target_path)

            print(f"文件已复制到: {target_path}")

            # 分离文件名和扩展名
            filename_without_ext, ext = os.path.splitext(doc)
            output = filename_without_ext + '.pdf'
            pdf_name = os.path.join(path, "pdf", output)  # 输出文件路径
            print(f"输出 PDF 文件路径: {pdf_name}")
            # 确保文件存在
            if not os.path.isfile(doc):
                print(f"文件不存在: {doc}")
                continue

            # 打开 Word 文档
            try:
                print(f"打开文档: {doc}")
                doc_obj = word.Documents.Open(f'"{doc}"', ReadOnly=1)
                print("文档已成功打开")

                # 转换为 PDF
                print(f"正在将文档转换为 PDF: {pdf_name}")
                doc_obj.ExportAsFixedFormat(pdf_name, constants.wdExportFormatPDF,
                                            Item=constants.wdExportDocumentWithMarkup,
                                            CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
                print(f"文档已转换为 PDF: {pdf_name}")
            except Exception as e:
                print(f"打开文档或转换为 PDF 时出错: {e}")
                continue  # 跳过当前文档，继续处理其他文件

            # 检查是否成功生成 PDF
            if os.path.isfile(pdf_name):
                try:
                    print(f"获取 PDF 页数: {pdf_name}")
                    pages = getPdfPageNum(pdf_name)  # 获取页数
                    print(f"成功获取页数: {pages}")
                    valueList.append([fullfilename, str(pages)])
                    totalPages += pages  # 累加总页数
                except Exception as e:
                    print(f"获取页数失败: {e}")

                # 删除临时 PDF 文件
                os.remove(pdf_name)
                print(f"已删除临时 PDF 文件: {pdf_name}")
            else:
                print(f"{fullfilename} 转换失败！")
                continue

        # 退出 Word 应用
        word.Quit(constants.wdDoNotSaveChanges)  # 退出 Word 应用
        print("已退出 Word 应用")
        return totalPages, valueList  # 返回总页数和文件列表

    except Exception as e:
        print(f"出错了: {e}")
        return -1, []

####################### 统计页码 ############################################

def getPdfPageNum(path):
    try:
        with open(path, "rb") as file:
            reader = PdfReader(file)  # 使用 PdfReader 代替 PdfFileReader
            pagecount = len(reader.pages)  # 获取 PDF 页数
        return pagecount
    except Exception as e:
        print(f"获取页数失败: {e}")
        return 0  # 返回 0 表示获取页数失败

####################### 提取目录 ############################################

# def getPdfOutlines(pdfpath,listpath,isList):
#     print("提取目录")
#     with open(pdfpath, "rb") as file:
#         doc = PdfFileReader(file)
#         outlines = doc.getOutlines()  # 获取大纲
#         global returnlist  # 全局变量，保存大纲的列表
#         returnlist = []   # 创建一个空列表
#         mylist = getOutline(outlines,isList)  # 递归获取大纲
#         w = DispatchEx("Word.Application")  # 创建Word文档应用程序对象
#         w.Visible = 1
#         w.DisplayAlerts = 0
#         doc1 = w.Documents.Add()# 添加一个Word文档对象
#         range1 = doc1.Range(0,0)
#         for item in mylist:       # 通过循环将获取的目录列表插入到Word文档对象中
#              range1.InsertAfter(item)
#         outpath = os.path.join(listpath,'list.docx') # 连接Word文档路径
#
#         doc1.SaveAs(outpath)  # 保存文件
#         doc1.Close()  # 关闭Word文档对象
#         w.Quit()  # 退出Word文档应用程序对象
#     return outpath

from PyPDF2 import PdfReader
import os
from win32com.client import DispatchEx


def getPdfOutlines(pdfpath, listpath, isList):
    print("提取目录")
    with open(pdfpath, "rb") as file:
        reader = PdfReader(file)  # 使用 PdfReader 替代 PdfFileReader
        outlines = reader.outline  # 使用新的 'outline' 属性获取大纲

        global returnlist  # 全局变量，保存大纲的列表
        returnlist = []  # 创建一个空列表

        def extract_outlines(outlines):
            result = []
            for outline in outlines:
                if isinstance(outline, list):  # 如果是嵌套大纲（子目录）
                    result.extend(extract_outlines(outline))  # 递归提取
                else:
                    result.append(outline.title)  # 提取标题
            return result

        mylist = extract_outlines(outlines)  # 提取大纲内容

        w = DispatchEx("Word.Application")  # 创建Word文档应用程序对象
        w.Visible = 1
        w.DisplayAlerts = 0
        doc1 = w.Documents.Add()  # 添加一个Word文档对象
        range1 = doc1.Range(0, 0)

        for item in mylist:  # 通过循环将获取的目录列表插入到Word文档对象中
            range1.InsertAfter(item)

        outpath = os.path.join(listpath, 'list.docx')  # 连接Word文档路径
        doc1.SaveAs(outpath)  # 保存文件
        doc1.Close()  # 关闭Word文档对象
        w.Quit()  # 退出Word文档应用程序对象

    return outpath


def getOutline(obj,isList):  # 递归获取大纲
    global returnlist
    for o in obj:
        if type(o).__name__ == 'Destination':
            # mypage = getRealPage(doc, pagecount, o.get('/Page').idnum)
            if isList:  # 包括页码
                returnlist.append( o.get('/Title') + "\t\t" + str(o.get('/Page') + 1) + "\n")
            else:       # 不包括页码
                returnlist.append(o.get('/Title') + "\n")
        elif type(o).__name__ == 'list':
            getOutline(o,isList)  # 递归调用获取大纲
    return returnlist
