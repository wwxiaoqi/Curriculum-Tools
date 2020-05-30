import os, datetime
from comtypes.client import CreateObject

# PPT 转 PDF
def ppt_pdf(path):
    # 开始时间
    startTime = datetime.datetime.now()

    pdf_path = path.replace('ppt', 'pdf')
    p = CreateObject("PowerPoint.Application")
    ppt = p.Presentations.Open(path)
    ppt.ExportAsFixedFormat(pdf_path, 2, PrintRange = None)
    ppt.Close()
    p.Quit()

    # 结束时间
    endTime = datetime.datetime.now()
    print('PPT 转换为 PDF, 用时：',(endTime - startTime).seconds)


# Word 转 PDF
def word_pdf(path):
    # 开始时间
    startTime = datetime.datetime.now()

    pdf_path = path.replace('doc', 'pdf')
    w = CreateObject("Word.Application")
    doc = w.Documents.Open(path)
    doc.ExportAsFixedFormat(pdf_path, 17)
    doc.Close()
    w.Quit()

    # 结束时间
    endTime = datetime.datetime.now()
    print('Word 转换为 PDF, 用时：',(endTime - startTime).seconds)

# Excel 转 PDF
def excel_pdf(path):
    # 开始时间
    startTime = datetime.datetime.now()

    pdf_path = path.replace('xls', 'pdf')

    # 如果 PDF 存在就删除
    if os.path.exists(pdf_path):
        os.remove(pdf_path)

    xlApp = CreateObject("Excel.Application")
    books = xlApp.Workbooks.Open(path)
    books.ExportAsFixedFormat(0, pdf_path)
    xlApp.Quit()

    # 结束时间
    endTime = datetime.datetime.now()
    print('Excel 转换为 PDF, 用时：',(endTime - startTime).seconds)