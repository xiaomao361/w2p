# encoding: utf8
from pathlib import Path
import os
from win32com import client
import re

def doc2pdf(doc_name, pdf_name):  
    try:
        word = client.DispatchEx("Word.Application")
        if os.path.exists(pdf_name):
            os.remove(pdf_name)
        worddoc = word.Documents.Open(doc_name,ReadOnly = 1)
        worddoc.SaveAs(pdf_name, FileFormat = 17)
        worddoc.Close()
        return pdf_name
    except Exception as e:
        print('convert error:' + str(e))
        return 1

if __name__ == '__main__':
    base_dir = '.'
    for doc_file in Path(base_dir).glob('*.doc*'):
        print('convert the word file: ' + str(doc_file))
        base_name = re.findall(r'(.+?)\.', str(doc_file))
        url = os.getcwd()
        doc_name = url + "\\" + str(doc_file)
        pdf_name = url + "\\" + str(base_name[0]) + '.pdf'
        doc2pdf(doc_name, pdf_name)
        