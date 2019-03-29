#!usr/bin/env python3
# -*- coding:utf-8 _*-

__author__ = 'Deicide-PiLi'
import io
import ssl
import re
import requests
from bs4 import BeautifulSoup
import xlsxwriter
import time

url = 'https://www.adidas.com.cn/men_shoes'

response = requests.get(url)  # 请求数据
soup = BeautifulSoup(response.text, 'html.parser')

print(soup.prettify())  # 打印出数据的文本内容
# I_want = soup.find_all('div', class_="product-grid")#.find('div').find_all('a',class_="col-12-3 col-sm-12-6 list-item")
# I_want = soup.find_all('div',class_="row product-grid-con float-clearfix product-list-grid")
# I_want.text
# for Iwant in I_want:
# Shoes_Info = I_want[0].find_all('div',class_="col-12-3 col-sm-12-6 list-item")
Shoes_Info = soup.find_all('div', class_="col-12-3 col-sm-12-6 list-item")
'''
#print(Shoes_Info.__len__())        #查看爬取到的鞋子数量
I_want = Shoes_Info[0]
print(I_want)
print(I_want.img['data-img-src'])
print(I_want.find_all('span')[1].text)
print(I_want.find_all('span')[2].text)
print(I_want.p.text)
#print(type(img_data))
file_path = 'G:/Python_Code/Spyder_First/Img_Get/'
file_name = '1.jpg'
'''
# 创建一个新Excel文件并添加一个工作表。
workbook = xlsxwriter.Workbook('G:/Python_Code/Sypder-Adidas/Adidas-Document/adidas.xlsx')
worksheet = workbook.add_worksheet()
i = 2
# 加宽第一列使文本更清晰
worksheet.set_column('A:A', 30)
worksheet.set_column('B:B', 30)
worksheet.set_column('C:C', 30)
worksheet.write('A1', "鞋子类型")
worksheet.write('B1', "鞋子数量： " + str(Shoes_Info.__len__()))
worksheet.write('C1', '鞋子名称：')
worksheet.write('D1', time.strftime("%Y-%m-%d", time.localtime()))
# 将鞋子的信息依次写入Excel表格中
for I_want in Shoes_Info:
    TypePos = 'A' + str(i)
    ImagePos = 'B' + str(i)
    NamePos = 'C' + str(i)
    PricePos = 'D' + str(i)
    ShoesType = I_want.find_all('span')[1].text  # 鞋子类型
    try:
        ShoesName = I_want.find_all('span')[2].text
    except:
        ShoesName = 'xxx'  # 鞋子名称
    ShoesPrice = I_want.p.text  # 鞋子价格
    worksheet.write(TypePos, ShoesType)
    worksheet.write(NamePos, ShoesName)
    worksheet.write(PricePos, ShoesPrice)
    ShoesImage_URL = I_want.img['data-img-src']  # 图片链接
    ShoesImageData = requests.get(I_want.img['data-img-src'])
    worksheet.insert_image(ImagePos, I_want.img['data-img-src'],
                           {'x_scale': 0.8, 'y_scale': 0.8, 'image_data': io.BytesIO(ShoesImageData.content)})
    i = i + 15
workbook.close()

# with open(file_path+file_name,"wb") as img:
#    img.write(img_data.content)
#    img.close()


