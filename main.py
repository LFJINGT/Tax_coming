from concurrent.futures import ThreadPoolExecutor, wait
import yaml
import os
import openpyxl as xl
import win32com.client as win32
import currency_symbols._constants as currency_symbols_constants
import xlrd
import xlsxwriter
import threading
import shutil
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import datetime
import time
import logging


# Path: main.py

# 读取配置文件
def read_config():
    with open('config.yaml', 'r', encoding='utf-8') as f:
        config = yaml.load(f, Loader=yaml.FullLoader)
    return config

# 初始化

# 读取配置文件
config = read_config()
logistic_config = config['logistic']
supply_config = config['supply']
trade_config = config['trade']
manufacture_config = config['manufacture']
time_list = [time_info for time_info,
                time_value in config['time'].items() if time_value == True]
company_id_list = [company_info['code'] for company_info in config['manufacture']]
company_name_list = [company_info['name'] for company_info in config['manufacture']]
company_dir_list = list(map(''.join, zip(company_id_list, company_name_list)))
logger = logging.getLogger(__name__)

# 根据时间戳生成log文件在/log文件夹中
def generate_log():
    if not os.path.exists('log'):
        os.mkdir('log')
    log_time = time.strftime("%Y%m%d%H%M%S", time.localtime())
    log_file = os.path.join('log', log_time + '.log')
    logging.basicConfig(filename=log_file, level=logging.INFO)



# 收集创建文件夹
'''
处理后目录结构
    data
        第x期
            企业代码+企业名称
                企业代码第X期财务表.xlsx
                企业代码第X期纳税申报表.xlsx
'''

# 构建目录结构
def build_dir():
# 文件夹初始化
    # 创建以期数命名的一级分类文件夹
    # 创建以企业名称命名的二级分类文件夹
    if not os.path.exists('data'):
        os.mkdir('data')
    for time in time_list:
        time_dir = os.path.join('data', time)
        if not os.path.exists(time_dir):
            os.mkdir(time_dir)
        for company_dir in company_dir_list:
            if not os.path.exists(os.path.join(time_dir, company_dir)):
                os.mkdir(os.path.join(time_dir, company_dir))

# 从origin/第一期文件夹中复制含有公司ID的文件到data/第一期文件夹中对应的公司文件夹中
def copy_file():
    for time in time_list:
        origin_dir = os.path.join('origin', time)
        data_dir = os.path.join('data', time)
        for company_dir in company_dir_list:
            company_id = company_dir[:4]
            for file in os.listdir(origin_dir):
                if company_id in file:
                    shutil.copy(os.path.join(origin_dir, file), os.path.join(data_dir, company_dir))

# 修改含财务关键词的文件名为企业代码：第X期财务表.xlsx 含纳税关键词的文件名为企业代码：第X期纳税申报表.xlsx 如果已经存在则删除
def rename_file():
    for time in time_list:
        data_dir = os.path.join('data', time)
        for company_dir in company_dir_list:
            for file in os.listdir(os.path.join(data_dir, company_dir)):
                if '财务' in file:
                    try:
                        os.rename(os.path.join(data_dir, company_dir, file),
                                    os.path.join(data_dir, company_dir, company_dir[:4] + time + '财务表.xlsx'))
                                    # 输出log 企业代码：第X期财务表 重命名成功
                        logging.info(company_dir[:4] + time + '财务表 重命名成功')

                    except:
                        os.remove(os.path.join(data_dir, company_dir, company_dir[:4] + time + '财务表.xlsx'))
                        os.rename(os.path.join(data_dir, company_dir, file),
                                    os.path.join(data_dir, company_dir, company_dir[:4] +  time + '财务表.xlsx'))
                        logging.info(company_dir[:4] + time + '财务表 重命名成功')
                elif '纳税' in file:
                    try:
                        os.rename(os.path.join(data_dir, company_dir, file),
                                    os.path.join(data_dir, company_dir, company_dir[:4] +  time + '纳税申报表.xlsx'))
                                    # 输出log 企业代码：第X期纳税申报表 重命名成功
                        logging.info(company_dir[:4] + time + '纳税申报表 重命名成功')
                    except:
                        os.remove(os.path.join(data_dir, company_dir, company_dir[:4] + time + '纳税申报表.xlsx'))
                        os.rename(os.path.join(data_dir, company_dir, file),
                                    os.path.join(data_dir, company_dir, company_dir[:4] + time + '纳税申报表.xlsx'))
                        logging.info(company_dir[:4] + time + '纳税申报表 重命名成功')

# 删除公司文件夹中所有文件
def delete_all_file():
    for time in time_list:
        data_dir = os.path.join('data', time)
        for company_dir in company_dir_list:
            for file in os.listdir(os.path.join(data_dir, company_dir)):
                os.remove(os.path.join(data_dir, company_dir, file))

# 删除公司文件夹中不含有企业代码的文件
def delete_file():
    for time in time_list:
        data_dir = os.path.join('data', time)
        for company_dir in company_dir_list:
            for file in os.listdir(os.path.join(data_dir, company_dir)):
                if company_dir[:4] not in file:
                    os.remove(os.path.join(data_dir, company_dir, file))
                    print('删除文件：' + os.path.join(data_dir, company_dir, file))


generate_log()
build_dir()
copy_file()
rename_file()
# delete_all_file()
delete_file()
logging.info('处理完成')
