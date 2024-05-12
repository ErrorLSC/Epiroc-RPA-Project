# -*- coding: utf-8 -*-
"""
Created on Thu Jan 19 18:06:54 2023

@author: JPEQZ
"""

import PyPDF2
#from AzureOCR import AzureOCR
from io import BytesIO


#new_fileName = '' # 分割後のファイル名

def splitPDF(src_path):
    org_pdf = PyPDF2.PdfReader(src_path)
    pagelist=[]
    for i in range(len(org_pdf.pages)):
        new_pdf = PyPDF2.PdfWriter()
        new_pdf.add_page(org_pdf.pages[i])
        #new_pdf.write("./multipage/page"+str(i)+".pdf")
        with BytesIO() as bytes_stream:
            new_pdf.write(bytes_stream)
            pagelist.append(bytes_stream.getvalue())
        new_pdf.close()

def split_merge_PDF(src_path,pages_list,output_location):
    org_pdf = PyPDF2.PdfReader(src_path)
    new_pdf = PyPDF2.PdfWriter()
    for page in pages_list:
        new_pdf.add_page(org_pdf.pages[page])
    new_pdf.write(output_location)
    return
        
    #print(pagelist) 
    #return pagelist

if __name__ == "__main__": 
    org_fileName = r"C:\Users\jpeqz\Desktop\HighPressureTanks\pdfcoffee.com_asme-viii-div1-ed2019-pdf-free.pdf"  # 分割したいファイルのファイル名
    target_fileName = r"C:\Users\jpeqz\Desktop\HighPressureTanks\神島組PRD60\ASME_VIII.pdf"
    split_merge_PDF(org_fileName,[246],target_fileName)
    #doc=splitPDF(org_fileName)
    #endpoint = "https://takanoocr.cognitiveservices.azure.com/"
    #credential = "69586bdfca724c7c942c226b52b4a0f7"
    #documentpath = "20221012133840.pdf"
    #AzureOCR(doc[109],endpoint,credential)