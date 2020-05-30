import os, sys, datetime
from utils.PdfConverter import excel_pdf
from utils.PdfConvertToPicture import pdf_png
from utils.OpenTemplateSave import save_with_template

# 行数
mLineNumber = 5

# 运行路径
mPath = sys.path[0] + '\\'

# 读取的 XLS
mOpenFullPath = mPath + 'Test_Template.xls'

# 写入的 XLS
mSaveFullPath = mPath + 'Test_Curriculum.xls'

# 导出的 PDF
mExportPDFPath = mPath + 'Test_Curriculum.pdf'

# 按照模板另存为 XLS
save_with_template(mLineNumber, mOpenFullPath, mSaveFullPath)

# Excel 转 PDF
excel_pdf(mSaveFullPath)

# PDF 转 PNG
pdf_png(mExportPDFPath, mPath)