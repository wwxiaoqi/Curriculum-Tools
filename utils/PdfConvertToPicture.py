import fitz
import os, sys, datetime

def pdf_png(pdfPath, imagePath):
    # 开始时间
    startTime = datetime.datetime.now()

    pdfDoc = fitz.open(pdfPath)
    for pg in range(pdfDoc.pageCount):
        page = pdfDoc[pg]
        rotate = int(0)

        # 每个尺寸的缩放系数为 1.3，这将为我们生成分辨率提高 2.6 的图像
        # 此处若是不做设置，默认图片大小为：792X612, dpi = 96
        # 1.33333333 --> 1056x816
        # 2.00000000 --> 1584x1224
        zoom_x = 5
        zoom_y = 5
        mat = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
        pix = page.getPixmap(matrix = mat, alpha = False)
        
        # 判断文件夹是否存在
        if not os.path.exists(imagePath):
            os.makedirs(imagePath)
        
        # 图片写入
        pix.writePNG(imagePath + '/' + 'Curriculum_%s.png' % pg)
    
    # 结束时间
    endTime = datetime.datetime.now()

    print('PDF 转换为 PNG, 用时：',(endTime - startTime).seconds)
 
 