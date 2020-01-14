#!/usr/bin/env python3
# -*- coding: utf-8 -*-

'''
@Author: Pengyang Zhao
@Date: 2020-01-14 09:37:55
@LastEditTime : 2020-01-14 14:31:06
@Description: 助教工具-期末成绩计算
@FilePath: \成绩汇总\cal_score.py
'''

import pandas as pd
import math

if __name__ == '__main__':
    dict_list = []
    # ==================
    #  导入数据
    # ==================
    # 导入平时作业，其命名格式为'1.xsl','2.xsl'....
    for i in range(1,9):
        try:
            data_fram = pd.read_excel(str(i)+'.xls', sheet_name='sheet1')
        except:
            continue
        stu_num = data_fram['学号'].values
        name = data_fram['姓名'].values
        score = data_fram['成绩（录入项）'].values
        score_dict = dict(zip(stu_num, score))
        dict_list.append(score_dict)

    # 导入期末大作业成绩
    data_fram = pd.read_excel('final.xls', sheet_name='sheet1')
    stu_num = data_fram['学号'].values
    name = data_fram['姓名'].values
    score = data_fram['成绩（录入项）'].values
    final_dict = dict(zip(stu_num, score))
    name_id_dict = dict(zip(stu_num, name)) #制作姓名和学号匹配字典
    #* 找出退课的同学，有需要的话后面可以展示出来
    withdraw = final_dict.keys()-dict_list[0].keys()# 默认第一次和最后一次的差为退课的同学

    # =====================
    #       计算成绩
    # =====================
    # 相关参数
    normal_p = 0.5  #平时成绩占的百分比
    final_p = 0.5   #期末成绩占的百分比
    normal_times = 7.0  #平时成绩的次数
    # 对于没交作业同学nan处理方式
    score_nan = 60
    # 平时成绩求和
    n_dict = {}
    count_dict = {}
    for key in dict_list[0]:
        count_dict[key]=0
        if dict_list[1].get(key):
            if dict_list[2].get(key):
                if dict_list[3].get(key):
                    if dict_list[4].get(key):
                        if dict_list[5].get(key):
                            if dict_list[6].get(key):
                                # 找出没给成绩的同学，是没交的同学，并统计没交次数
                                if math.isnan(dict_list[0][key]):
                                    count_dict[key]=count_dict[key]+1
                                    dict_list[0][key] = score_nan
                                if math.isnan(dict_list[1][key]):
                                    count_dict[key]=count_dict[key]+1
                                    dict_list[1][key] = score_nan
                                if math.isnan(dict_list[2][key]):
                                    count_dict[key]=count_dict[key]+1
                                    dict_list[2][key] = score_nan
                                if math.isnan(dict_list[3][key]):
                                    count_dict[key]=count_dict[key]+1
                                    dict_list[3][key] = score_nan
                                if math.isnan(dict_list[4][key]):
                                    count_dict[key]=count_dict[key]+1
                                    dict_list[4][key] = score_nan
                                if math.isnan(dict_list[5][key]):
                                    count_dict[key]=count_dict[key]+1
                                    dict_list[5][key] = score_nan
                                if math.isnan(dict_list[6][key]):
                                    count_dict[key]=count_dict[key]+1
                                    dict_list[6][key] = score_nan
                                # 算平时成绩的加和
                                n_dict[key] = dict_list[0][key]+dict_list[1][key]+dict_list[2][key]+dict_list[3][key]\
                                            +dict_list[4][key]+dict_list[5][key]+dict_list[6][key]
    # 写入Excel操作
    my_fram = []
    for key in name_id_dict:
        ff = final_dict[key].copy()
        nn = n_dict[key].copy()
        score = ff * final_p + nn / normal_times * normal_p
        my_fram.append([key, name_id_dict[key], final_dict[key], n_dict[key], count_dict[key], score])
    my_fram1 = pd.DataFrame(data =my_fram,
                            columns=['学号','姓名','期末成绩','平时总成绩','未交次数','机算期末成绩'])
    writer = pd.ExcelWriter("程序汇总.xls")
    my_fram1.to_excel(writer, 'sheet1')
    writer.save()
    print('统计完毕')