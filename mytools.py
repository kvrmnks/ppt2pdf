import comtypes.client
import os

import PyPDF2
import re


def convert(input_file: str, output: str, format_code=32) -> None:
    print('converting ' + input_file)
    comtypes.CoInitialize()  # 一定要初始化
    ppt = comtypes.client.CreateObject('Powerpoint.Application')
    ppt.Visible = 1
    deck = ppt.Presentations.Open(input_file)
    deck.SaveAs(output, format_code)
    deck.Close()
    ppt.Quit()


def convert_ppt_pdf(fileName: str, path: str) -> str:
    if fileName.endswith('ppt') or fileName.endswith('pptx'):
        pdf_name = fileName
        if fileName.endswith('ppt'):
            pdf_name = fileName[:-3] + 'pdf'
        if fileName.endswith('pptx'):
            pdf_name = fileName[:-4] + 'pdf'
        convert('"' + path + '"', '"' + os.getcwd() + '/' + pdf_name + '"')
        return os.getcwd() + '/' + pdf_name


def merge_pdf(fileList: list, labelList: list, out: str) -> None:
    main = PyPDF2.PdfFileMerger()
    for i in range(len(fileList)):
        main.append(fileList[i], bookmark=labelList[i])
        print('merged ' + fileList[i])
    main.write(out)
