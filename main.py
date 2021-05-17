import comtypes.client
import os
import PyPDF2
import re


def convert(input_file: str, output: str, format_code=32) -> None:
    print(input_file)
    ppt = comtypes.client.CreateObject('Powerpoint.Application')
    ppt.Visible = 1
    deck = ppt.Presentations.Open(input_file)
    deck.SaveAs(output, format_code)
    deck.Close()
    ppt.Quit()


def convert_ppt_pdf() -> list:
    regex = re.compile(r'\d+')
    file_list = os.listdir()
    file_list = [x for x in file_list if len(regex.findall(x)) > 0]
    file_list.sort(key=lambda x: int(regex.findall(x)[0]))
    print(file_list)
    pdf_list = []
    for x in file_list:
        if x.endswith('ppt') or x.endswith('pptx'):
            pdf_name = x
            if x.endswith('ppt'):
                pdf_name = x[:-3] + 'pdf'
            if x.endswith('pptx'):
                pdf_name = x[:-4] + 'pdf'
            pdf_list.append(pdf_name)
            convert('"' + os.getcwd() + '/' + x + '"', '"' + os.getcwd() + '/' + pdf_name + '"')
    return pdf_list


def merge():
    pdf_file_list = [x for x in os.listdir() if x.endswith('.pdf')]
    regex = re.compile(r'\d+')
    pdf_file_list = [x for x in pdf_file_list if len(regex.findall(x)) > 0]
    pdf_file_list.sort(key=lambda x: int(regex.findall(x)[0]))
    print(pdf_file_list)
    main = PyPDF2.PdfFileMerger()
    for x in pdf_file_list:
        main.append(os.getcwd() + '/' + x, bookmark=x[:-4])
    main.write('./main.pdf')


if __name__ == '__main__':
    # print(os.listdir())
    convert_ppt_pdf()
    merge()
