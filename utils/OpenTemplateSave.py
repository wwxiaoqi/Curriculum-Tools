import xlrd, xlwt
import os, sys, datetime
from xlutils.copy import copy

def save_with_template(mLineNumber, mOpenFullPath, mSaveFullPath):
    #  XLS 模板
    mTemplateFullPath = sys.path[0] + '\\' + 'template\\curriculum_template.xls'

    save(mLineNumber, mOpenFullPath, mTemplateFullPath, mSaveFullPath)


def save(mLineNumber, mOpenFullPath, mTemplateFullPath, mSaveFullPath):
    # 开始时间
    startTime = datetime.datetime.now()

    # 如果 XLS 存在就删除
    if os.path.exists(mSaveFullPath):
        os.remove(mSaveFullPath)

    # 打开 XLS && 读取 XLS
    xlsx = xlrd.open_workbook(mOpenFullPath)
    table = xlsx.sheet_by_index(0)
    old_excel = xlrd.open_workbook(mTemplateFullPath, formatting_info = True)
    new_excel = copy(old_excel)
    ws = new_excel.get_sheet(0)

    # 边框
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN

    # 字体
    font = xlwt.Font()
    # 字体类型
    font.name = 'Arial'
    # 字体颜色
    font.colour_index = 0
    # 字体大小，11为字号，20为衡量单位
    font.height = 20 * 11
    # 字体加粗
    font.bold = False
    # 下划线
    font.underline = False
    # 斜体字
    font.italic = False

    # 写入 XLS 的主题，默认设置为：自动换行，水平居中，垂直居中，边框，字体....
    style = xlwt.XFStyle()
    alignment = xlwt.Alignment()
    alignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER
    style.alignment = alignment
    style.borders = borders
    style.font = font

    count = 0
    data1 = 5 # [5, 9, 13, 17, 21]
    data2 = 0 # [0, 2, 4]
    for i in range(1, 31, 2):
        count += 1
        ws.write(data1, data2, table.cell(mLineNumber - 1, i).value, style)

        if count >= 3:
            data1 += 4
            count = 0
    
        if data2 == 4:
            data2 = 0
        else:
            data2 += 2

    ws.write(0, 0, table.cell(0, 0).value, style)
    new_excel.save(mSaveFullPath)

    # 结束时间
    endTime = datetime.datetime.now()
    print('保存成功，用时：', (endTime - startTime).seconds)