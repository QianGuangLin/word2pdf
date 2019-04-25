from win32com.client import Dispatch # pip install pywin32
from win32com.client import DispatchEx
from os import walk
import os
import gc


wdFormatPDF = 17  # win32提供了多种word转换为其他文件的接口，其中FileFormat=17是转换为pdf

def doc2pdf(input_file, input_file_name, output_dir):

	# 声明全局变量
    global doc, word

    try:
        # word = Dispatch('Word.Application')
        word = DispatchEx('Word.Application')
        doc = word.Documents.Open(input_file)
    except Exception as e:
        print("word无法打开, 发生如下错误:\n{}".format(e))

    try:
        pdf_file_name = input_file_name.replace(".docx", ".pdf").replace(".doc", ".pdf")
        pdf_file = os.path.join(output_dir, pdf_file_name)
        doc.SaveAs(pdf_file, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()
        print("成功转换\"{}\"".format(input_file_name))
        print()

        # or 杀死进程
        # del(word)
        # or 内存回收
        gc.collect()

    except Exception as e:
        print("文件保存失败, 发生如下错误:\n{}".format(e))


if __name__ == "__main__":
    doc_files = []

    # 绝对路径
    # directory = "E:\\Python\\work\\word2pdf\\word" # word文件夹
    # output_dir = "E:\\Python\\work\\word2pdf\\pdf" # pdf文件夹

    # 相对路径
    path1 = os.path.abspath('.')  # 表示当前所处的文件夹的绝对路径
    path2 = os.path.abspath('..')  # 表示当前所处的文件夹上一级文件夹的绝对路径
    directory = path1+"/word"  # word文件夹
    output_dir = path1+"/pdf"  # pdf文件夹


    for root, _, filenames in walk(directory):  # 第二个返回值是dirs， 用不上使用_占位
        for file in filenames:
            if file.endswith(".doc") or file.endswith(".docx"):
                print("转换{}中......".format(file))
                doc2pdf(os.path.join(root, file), file, output_dir)


