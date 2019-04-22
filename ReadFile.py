#!usr/bin/env python3
#-*- coding:utf-8 _*-
__author__ = 'Deicide-PiLi'

import xlrd
import xlsxwriter
import requests
import io

InfoPath = r'G:/Python_Code/Sypder-Adidas/Adidas-Document/adidas.xls'
ImagePath = r'G:/Python_Code/Sypder-Adidas/Adidas-Document/Image.xlsx'

'''
读取源文件中的图片链接，下载并写入另一个xlsx中
src_path : 源文件
des_path : 目的文件
'''
def ReWrite(scr_path,des_path):
    # 创建一个新Excel文件并添加一个工作表。
    des_book = xlsxwriter.Workbook(des_path)
    des_sheet = des_book.add_worksheet()
    des_sheet.set_column('B:B', 30)

    src_book = xlrd.open_workbook(scr_path, formatting_info=True)
    sheet1 = src_book.sheet_by_index(0)  # sheet索引从0开始
    shoesnum=0
    cols = sheet1.col_values(2)  # 获取第三列内容
    for name in cols:
        if name != '':
            shoesnum=shoesnum+1
    shoesnum=shoesnum-1     #表格中鞋子的数量
    for i in range(shoesnum):
        ImgURL = str(sheet1.cell(i*6+2, 0).value)
        #url=r'https://img.adidas.com.cn/resources/2019/4/15/15553080368405000_230X230.jpg'
        print('正在打印第 '+str(i+1)+' 张图片')
        ShoesImageData = requests.get(ImgURL)  # 获取图片
        des_sheet.insert_image(i*6+1 , 1 , ImgURL,
                           {'x_scale': 0.5, 'y_scale': 0.5, 'image_data': io.BytesIO(ShoesImageData.content)})  #下载并写入图片
    des_book.close()

def read_excel(file_path):
    read_info = []
    # 打开文件
    workbook = xlrd.open_workbook(file_path,formatting_info=True)
    # 获取所有sheet
    #print(workbook.sheet_names())  # [u'sheet1', u'sheet2']
    # 根据sheet索引或者名称获取sheet内容
    sheet1 = workbook.sheet_by_index(0)  # sheet索引从0开始
    # sheet的名称，行数，列数
    #print(sheet1.name, sheet1.nrows, sheet1.ncols)
    # 获取整行和整列的值（数组）
    #rows = sheet1.row_values(3)  # 获取第四行内容
    cols = sheet1.col_values(2)  # 获取第三列内容
    #print(rows)
    #print(cols)
    for name in cols:
        if name != '':
            read_info.append(name)
    print(read_info)
    # 获取单元格内容
    # print(sheet1.cell(1, 0).value)
    # print(sheet1.cell_value(1, 0).encode('utf-8'))
    # print(sheet1.row(1)[0].value.encode('utf-8'))

    return workbook,read_info

if __name__ == '__main__':
    ReWrite(InfoPath,ImagePath)