#""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
#
# Copyright (c) 2024 愤怒的it男, All Rights Reserved.
# FileName : code.py
# Date     : 2024.01.22
# Author   : 愤怒的it男
# Version  : 1.0.0
# Node     : 欢迎关注微信公众号【愤怒的it男】
#
#""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

import csv
import json
import requests
from tabulate import tabulate

def getData(baseUrl, headers):
    response = requests.get(url=baseUrl, headers=headers)
    data = json.loads(response.text)['data']
    result = []
    for i,d in enumerate(data):
        index = str(i+1)
        title = d['Title']
        label = d['Label']
        url = d['Url']
        result.append([index,title,label,url])
    return result

def printData(result):
    headers = ["排名", "标题", "标签", "链接"]
    table = tabulate(result, headers=headers, maxcolwidths=[None, None, None, 80], tablefmt='grid')
    print(table)

def saveData(result):
    with open('头条热榜.csv', 'w', encoding='utf-8', newline='') as file:
        writer = csv.writer(file)
        writer.writerows(result)

def main():
    baseUrl = 'https://www.toutiao.com/hot-event/hot-board/?origin=toutiao_pc'
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'}
    result = getData(baseUrl, headers)
    printData(result)
    saveData(result)

if __name__== "__main__" :
    main()
