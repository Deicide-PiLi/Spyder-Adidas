#!usr/bin/env python3
#-*- coding:utf-8 _*-

__author__ = 'Deicide-PiLi'
import requests
from bs4 import BeautifulSoup
import xlrd
from xlutils.copy import copy
import io

src_url = r'https://www.adidas.com.cn/men_shoes'
FilePath = r'G:/Python_Code/Sypder-Adidas/Adidas-Document/adidas.xls'
ImagePath = r'G:/Python_Code/Sypder-Adidas/Adidas-Document/Image.xlsx'
ShoesName=[]
ShoesType=[]
ShoesPrice=[]
ShoesImageURL=[]

def GetShoesInfo(url):
    response = requests.get(url)  # 请求数据
    soup = BeautifulSoup(response.text, 'html.parser')
    ShoesTxt = soup.find_all('div', class_="col-12-3 col-sm-12-6 list-item")
    # 将鞋子的信息依次写入Excel表格中
    for I_want in ShoesTxt:
        Name = I_want.find('div', class_='goods-title').find('span').text  # 鞋子名称
        Type = I_want.find('div', class_='goods-info').span.text  # 鞋子类型
        Price = I_want.find('p', class_='goods-price price-single').text  # 鞋子价格
        ImageURL = I_want.img['data-img-src']  # 图片链接
        ShoesName.append(Name)
        ShoesType.append(Type)
        ShoesPrice.append(Price)
        ShoesImageURL.append(ImageURL)
        #ShoesImageData = requests.get(I_want.img['data-img-src'])  # 获取图片

def XlsWrite(file_path):
    OldShoesName=[]
    #读取现有信息
    workbook = xlrd.open_workbook(file_path,formatting_info=True)
    sheet1 = workbook.sheet_by_index(0)  # sheet索引从0开始
    # sheet的名称，行数，列数
    cols = sheet1.col_values(2)  # 获取第三列内容
    for name in cols:
        if name != '':
            OldShoesName.append(name)
    OldNum = len(OldShoesName)
    xlsc = copy(workbook)
    worksheet = xlsc.get_sheet(0)
    i = (OldNum-1)*6+1
    #爬取网站信息
    GetShoesInfo(src_url)
    NewShoes=0
    for j in range(len(ShoesName)):
        if ShoesName[j] not in OldShoesName:
            NewShoes=NewShoes+1
            worksheet.write(i,0,ShoesType[j])
            worksheet.write(i,2,ShoesName[j])
            worksheet.write(i+1, 0, ShoesImageURL[j])
            worksheet.write(i,3,ShoesPrice[j])
            i = i+6
    if NewShoes==0:
        print('There are not new shoes !')
    else:
        print('主人，我找到了'+str(NewShoes)+'双鞋子')
    # for j in range(NewShoes+OldNum):
    #     ImgURL = sheet1.cell((j-1)*6+2, 0).value
    #     print(ImgURL)
        # ShoesImageData = requests.get(ImgURL)  # 获取图片
        # ImagePos='B'+str((j-1)*6+1)
        # sheet1.insert_image(ImagePos, ImgURL,{'x_scale': 0.5, 'y_scale': 0.5,'image_data': io.BytesIO(ShoesImageData.content)})  # 下载并写入图片
    xlsc.save(file_path)


if __name__ == '__main__':
    XlsWrite(FilePath)