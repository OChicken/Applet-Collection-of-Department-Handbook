#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Date    : 2018-09-27 21:27:16
# @Author  : Ma Seoyin (Ma.Seoyin@gmail.com)
# @Link    : https://github.com/OChicken
# @Version : V1

import os
import xlrd
import numpy as np

# 从RawData中pick out出萌新们的个人信息和第一第二志愿. RawData导出自www.wjx.cn问卷星 ------------------------------------------------------------------
Dir  = os.getcwd() + '/'
print('欢迎使用 物理学术竞赛报名信息整理 小程序 :)\n<<<<<<<<<<<<<<<<<< 啦啦啦我是分割线 >>>>>>>>>>>>>>>>>>')
print('author:\n马守然 (2014级应用物理学)\n学术科创部\n物理与光电学院团委学生会'
      '\nEmail: 1941688873@qq.com / Ma.Seoyin@gmail.com\n<<<<<<<<<<<<<<<<<< 啦啦啦我是分割线 >>>>>>>>>>>>>>>>>>')
FileName = input('请输入 从问卷星后台\'按选项序号下载\'导出的xls文件的文件名(切勿包含扩展名!):\n')
# FileName = '16987266_2_华南理工大学第六届物理学术竞赛报名表_73_73'; print(FileName)
Info = Dir + '报名信息/'
if os.path.exists(Info.rstrip('/')) == False:
    os.mkdir(Info)
RawData = xlrd.open_workbook(Dir + FileName + '.xls').sheet_by_index(0)
Num  = list(map(str, map(int, RawData.col_values(0)[1:])))
for i in range(9):
    Num[i] = '0' + Num[i]
Name = list(map(str, RawData.col_values(6)[1:]))
Len = len(Num)
for i in range(Len):
    FolderName = Info + Num[i] + '-' + Name[i]
    if os.path.exists(FolderName) == False:
        os.mkdir(FolderName)
input('\'报名信息\'文件夹已生成, 请按回车关闭本宝宝 :)')
