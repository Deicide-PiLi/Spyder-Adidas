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

url = 'https://www.adidas.com.cn/search?ni=62&pf=25-40%2C25-60%2C25-60%2C25-40%2C25-60%2C25-60&cf=2-8%2C2-8&pr=-&fo=p25%2Cp25%2Cc2%2Cp25%2Cp25%2Cc2&pn=2&pageSize=120&c=%E9%9E%8B%E7%B1%BB-%E9%9E%8B%E7%B1%BB&p=undefined-%E7%94%B7%E5%AD%90%26undefined-%E4%B8%AD%E6%80%A7%26undefined-%E7%94%B7%E5%AD%90%26undefined-%E4%B8%AD%E6%80%A7&isSaleTop=false'
txt_path = r'G:/Python_Code/Sypder-Adidas/Adidas-Document/url.txt'
response = requests.get(url)  # 请求数据
soup = BeautifulSoup(response.text, 'html.parser')
# with open(txt_path, 'a', encoding='utf-8') as f:
# #     f.write(soup.prettify())
# # f.close
print(soup.prettify())  # 打印出数据的文本内容
Shoes_Info = soup.find_all('div', class_="col-12-3 col-sm-12-6 list-item")

# 创建一个新Excel文件并添加一个工作表。
workbook = xlsxwriter.Workbook('G:/Python_Code/Sypder-Adidas/Adidas-Document/adidas.xls')
worksheet = workbook.add_worksheet()
i = 2
# 加宽第一列使文本更清晰
worksheet.set_column('A:A', 30)
worksheet.set_column('B:B', 30)
worksheet.set_column('C:C', 30)
worksheet.write('A1', "鞋子类型")
worksheet.write('B1', "鞋子数量： " + str(Shoes_Info.__len__()))
worksheet.write('C1', '鞋子名称：')
worksheet.write('D1', time.strftime("%m-%d", time.localtime()))
# 将鞋子的信息依次写入Excel表格中
for I_want in Shoes_Info:
    TypePos = 'A' + str(i)
    ImagePos = 'B' + str(i)
    NamePos = 'C' + str(i)
    PricePos = 'D' + str(i)
    ImageURL = 'A' + str(i+1)
    ShoesType = I_want.find('div',class_='goods-info').span.text            # 鞋子类型
    ShoesName = I_want.find('div',class_='goods-title').find('span').text   # 鞋子名称
    ShoesPrice = I_want.find('p',class_ = 'goods-price price-single').text  # 鞋子价格
    ShoesImage_URL = I_want.img['data-img-src']  # 图片链接
    worksheet.write(TypePos, ShoesType)
    worksheet.write(NamePos, ShoesName)
    worksheet.write(PricePos, ShoesPrice)
    worksheet.write(ImageURL, ShoesImage_URL)
    ShoesImageData = requests.get(I_want.img['data-img-src'])   #获取图片
    worksheet.insert_image(ImagePos, I_want.img['data-img-src'],
                           {'x_scale': 0.5, 'y_scale': 0.5, 'image_data': io.BytesIO(ShoesImageData.content)})  #下载并写入图片
    i = i + 6
workbook.close()

# with open(file_path+file_name,"wb") as img:
#    img.write(img_data.content)
#    img.close()


