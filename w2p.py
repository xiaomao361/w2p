# encoding: utf8
import os
import re
from win32com import client
from pathlib import Path


def doc2pdf(doc_name, pdf_name):
    try:
        word = client.DispatchEx("Word.Application")
        if os.path.exists(pdf_name):
            os.remove(pdf_name)
        word_doc = word.Documents.Open(doc_name, ReadOnly=1)
        word_doc.SaveAs(pdf_name, FileFormat=17)
        word_doc.Close()
        word.Quit()
        return pdf_name
    except Exception as e:
        print('convert error:' + str(e))
        return 1


if __name__ == '__main__':
    base_dir = '.'
    for doc_file in Path(base_dir).glob('*.doc*'):
        # 取右侧第一个点，左边的内容作为文件的base name
        base_name = str(doc_file)[:str(doc_file).rindex(".")]
        # 不对word缓存文件做转码（如果有文件打开中）
        if "~$" not in base_name:
            print('convert the word file: ' + str(doc_file))
            url = os.getcwd()
            doc_name = url + "\\" + str(doc_file)
            pdf_name = url + "\\" + str(base_name) + '.pdf'
            doc2pdf(doc_name, pdf_name)
