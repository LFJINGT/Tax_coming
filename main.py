from concurrent.futures import ThreadPoolExecutor, wait
import yaml
import os
import openpyxl
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

# 引入本地模块
from xls2xlsx import XLS2XLSX


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
company_code_list = [company_info['code']
                   for company_info in config['manufacture'] + config['trade'] + config['supply'] + config['logistic']]
company_name_list = [company_info['name']
                     for company_info in config['manufacture'] + config['trade'] + config['supply'] + config['logistic']]
company_dir_list = list(map(''.join, zip(company_code_list, company_name_list)))
company_code_to_name = dict(zip(company_code_list, company_name_list))

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
    logging.debug('创建目录结构成功!')

# 从origin/第一期文件夹中复制含有公司ID的文件到data/第一期文件夹中对应的公司文件夹中


def copy_file():
    for time in time_list:
        origin_dir = os.path.join('origin', time)
        data_dir = os.path.join('data', time)
        for company_dir in company_dir_list:
            company_id = company_dir[:4]
            for file in os.listdir(origin_dir):
                if company_id in file:
                    shutil.copy(os.path.join(origin_dir, file),
                                os.path.join(data_dir, company_dir))
    logging.debug('复制文件成功!')

# 修改含财务关键词的文件名为企业代码：第X期财务表.xlsx 含纳税关键词的文件名为企业代码：第X期纳税申报表.xlsx 如果已经存在则删除


def rename_file():
    for time in time_list:
        data_dir = os.path.join('data', time)
        for company_dir in company_dir_list:
            for file in os.listdir(os.path.join(data_dir, company_dir)):
                if '财务' in file and file != company_dir[:4] + time + '财务表.xlsx':
                    try:
                        os.rename(os.path.join(data_dir, company_dir, file),
                                  os.path.join(data_dir, company_dir, company_dir[:4] + time + '财务表.xlsx'))
                        # 输出log 企业代码：第X期财务表 重命名成功
                        logging.info(
                            company_dir[:4] + time + '财务表 重命名成功!' + " 原文件名为:" + file)
                        print(company_dir[:4] + time +
                              '财务表 重命名成功!' + " 原文件名为:" + file)

                    except:
                        os.remove(os.path.join(data_dir, company_dir,
                                  company_dir[:4] + time + '财务表.xlsx'))
                        os.rename(os.path.join(data_dir, company_dir, file),
                                  os.path.join(data_dir, company_dir, company_dir[:4] + time + '财务表.xlsx'))
                        # 输出log 企业代码：第X期财务表 重命名成功
                        logging.info(
                            company_dir[:4] + time + '财务表 重命名成功!' + " 原文件名为:" + file)
                        print(company_dir[:4] + time +
                              '财务表 重命名成功!' + " 原文件名为:" + file)
                elif '纳税' in file and file != company_dir + time + '纳税申报表.xlsx':
                    try:
                        os.rename(os.path.join(data_dir, company_dir, file),
                                  os.path.join(data_dir, company_dir, company_dir[:4] + time + '纳税申报表.xlsx'))
                        # 输出log 企业代码：第X期纳税申报表 重命名成功
                        logging.info(
                            company_dir[:4] + time + '纳税申报表 重命名成功!' + " 原文件名为:" + file)
                        print(company_dir[:4] + time +
                              '纳税申报表 重命名成功!' + " 原文件名为:" + file)
                    except:
                        os.remove(os.path.join(data_dir, company_dir,
                                  company_dir[:4] + time + '纳税申报表.xlsx'))
                        os.rename(os.path.join(data_dir, company_dir, file),
                                  os.path.join(data_dir, company_dir, company_dir[:4] + time + '纳税申报表.xlsx'))
                        # 输出log 企业代码：第X期纳税申报表 重命名成功
                        logging.info(
                            company_dir[:4] + time + '纳税申报表 重命名成功!' + " 原文件名为:" + file)
                        print(company_dir[:4] + time +
                              '纳税申报表 重命名成功!' + " 原文件名为:" + file)

# 如果存在xls文件则转换为xlsx文件


def xls_to_xlsx():
    for time in time_list:
        data_dir = os.path.join('data', time)
        for company_dir in company_dir_list:
            for file in os.listdir(os.path.join(data_dir, company_dir)):
                if file.endswith('.xls'):
                    try:
                        excel = XLS2XLSX(os.path.join(
                            data_dir, company_dir, file))
                        excel.to_xlsx(os.path.join(
                            data_dir, company_dir, file[:-4] + '.xlsx'))
                        os.remove(os.path.join(data_dir, company_dir, file))
                        # 输出log 企业代码：第X期财务表 转换成功
                        logging.info(
                            company_dir[:4] + time + '财务表 转换成功!' + " 原文件名为:" + file)
                        print(company_dir[:4] + time +
                              '财务表 转换成功!' + " 原文件名为:" + file)
                    except:
                        # 输出log 企业代码：第X期财务表 转换失败
                        logging.info(
                            company_dir[:4] + time + '财务表 转换失败!' + " 原文件名为:" + file)
                        print(company_dir[:4] + time +
                              '财务表 转换失败!' + " 原文件名为:" + file)


# 删除公司文件夹中所有文件


def delete_all_file():
    for time in time_list:
        data_dir = os.path.join('data', time)
        for company_dir in company_dir_list:
            for file in os.listdir(os.path.join(data_dir, company_dir)):
                os.remove(os.path.join(data_dir, company_dir, file))
    logging.debug('删除所有文件成功!')

# 删除公司文件夹中不含有企业代码的文件


def delete_file():
    for time in time_list:
        data_dir = os.path.join('data', time)
        for company_dir in company_dir_list:
            for file in os.listdir(os.path.join(data_dir, company_dir)):
                if company_dir[:4] not in file:
                    os.remove(os.path.join(data_dir, company_dir, file))
                    print('删除文件：' + os.path.join(data_dir, company_dir, file))

# 复制汇总表模板为新的汇总表


def copy_template():
    shutil.copyfile('税务局汇总表_模板.xlsx', '税务局汇总表.xlsx')

# 财务表


class finance_table:
    def __init__(self, company_code, time):
        self.company_code = company_code
        self.company_name = company_code_to_name[company_code]
        self.time = time
        self.data_dir = os.path.join('data', time)
        self.file_name = os.path.join(
            self.data_dir, company_code + self.company_name, company_code + time + '财务表.xlsx')
        self.sheet_name = '简易利润表'
        self.df = pd.read_excel(
            self.file_name, sheet_name=self.sheet_name, header=0, index_col=0)
        # 清洗数据
        self.df = self.df.iloc[5:19, 0:1]
        # 重命名列名
        self.df.columns = ['本期金额（不含税）']
        # 重命名索引
        self.df.index = ['营业收入', '营业成本', '管理费用', '销售费用', '财务费用', '研发费用', '投资收益', '资产处置收益',
                         '营业利润', '营业外收入', '营业外支出', '利润总额', '所得税费用', '净利润']
        # 清洗空值为0
        self.df = self.df.fillna(0)

    # 返回总表
    def get_total_table(self):
        return self.df

    # 返回公司代码
    def get_company_code(self):
        return self.company_code

    # 返回公司名称
    def get_company_name(self):
        return self.company_name

    # 返回期数
    def get_time(self):
        return self.time

    # 返回营业收入
    def get_revenue(self):
        return self.df.loc['营业收入', '本期金额（不含税）']

    # 返回营业成本
    def get_cost(self):
        return self.df.loc['营业成本', '本期金额（不含税）']

    # 返回管理费用
    def get_management_expense(self):
        return self.df.loc['管理费用', '本期金额（不含税）']

    # 返回销售费用
    def get_sales_expense(self):
        return self.df.loc['销售费用', '本期金额（不含税）']

    # 返回财务费用
    def get_financial_expense(self):
        return self.df.loc['财务费用', '本期金额（不含税）']

    # 返回研发费用
    def get_research_expense(self):
        return self.df.loc['研发费用', '本期金额（不含税）']

    # 返回投资收益
    def get_investment_income(self):
        return self.df.loc['投资收益', '本期金额（不含税）']

    # 返回资产处置收益
    def get_asset_disposal_income(self):
        return self.df.loc['资产处置收益', '本期金额（不含税）']

    # 返回营业利润
    def get_operating_profit(self):
        return self.df.loc['营业利润', '本期金额（不含税）']

    # 返回营业外收入
    def get_non_operating_income(self):
        return self.df.loc['营业外收入', '本期金额（不含税）']

    # 返回营业外支出
    def get_non_operating_expense(self):
        return self.df.loc['营业外支出', '本期金额（不含税）']

    # 返回利润总额
    def get_total_profit(self):
        return self.df.loc['利润总额', '本期金额（不含税）']

    # 返回所得税费用
    def get_income_tax_expenses(self):
        return self.df.loc['所得税费用', '本期金额（不含税）']

    # 返回净利润
    def get_net_profit(self):
        return self.df.loc['净利润', '本期金额（不含税）']


# 纳税申报表
class tax_table:
    def __init__(self, company_code, time):
        self.company_code = company_code
        self.company_name = company_code_to_name[company_code]
        self.time = time
        self.data_dir = os.path.join('data', time)
        self.file_name = os.path.join(
            self.data_dir, company_code + self.company_name, company_code + time + '纳税申报表.xlsx')
        self.sheet_name = '汇总表'
        self.df = pd.read_excel(
            self.file_name, sheet_name=self.sheet_name)
        # 清洗数据
        self.df = self.df.iloc[3:9, 1:6]
        # 重命名列名
        self.df.columns = ['企业增值税', '企业所得税', '五险一金', '合计']
        # 重命名索引
        self.df.index = ['第一期', '第二期', '第三期', '第四期', '第五期', '合计']
        # 检查数据是否正确 企业增值税+企业所得税+五险一金=合计 输出公司代码+期数+错误信息
        if self.df['企业增值税'].sum() + self.df['企业所得税'].sum() + self.df['五险一金'].sum() != self.df['合计'].sum():
            print(self.company_code + self.time + '纳税申报表数据错误')
        # 清洗空值为0
        self.df = self.df.fillna(0)

    # 返回总表
    def get_total_table(self):
        return self.df

    # 返回企业名称
    def get_company_name(self):
        return self.company_name

    # 返回企业代码
    def get_company_code(self):
        return self.company_code

    # 返回期数
    def get_time(self):
        return self.time

    # 根据期数返回企业增值税
    def get_vat(self):
        return self.df.loc[self.time, '企业增值税']

    # 返回企业所得税
    def get_income_tax(self):
        return self.df.loc[self.time, '企业所得税']

    # 返回五险一金
    def get_insurance(self):
        return self.df.loc[self.time, '五险一金']

    # 返回合计
    def get_total(self):
        return self.df.loc[self.time, '合计']


# 企业信息表
# 汇总表
class summary_table:
    def __init__(self, time):
        self.time = time
        self.df = pd.read_excel(
            '税务局汇总表.xlsx', sheet_name=time + '企业缴税汇总表', header=0)
        self.df_info = pd.read_excel(
            '税务局汇总表.xlsx', sheet_name=time + '企业漏缴税款情况', header=0)
        # 企业缴税汇总表清洗数据
        # 删除序号列为非数字的行
        self.df = self.df[self.df['序号'].apply(lambda x: isinstance(x, int))]

    # 返回总表
    def get_total_table(self):
        return self.df

    # 返回企业信息表
    def get_info_table(self):
        return self.df_info

    # 将纳税申报表并到汇总表中
    def merge_tax_table(self, tax_table):
        company_code = tax_table.get_company_code()
        # 如果输入表格的期数不是汇总表的期数，返回错误
        if tax_table.get_time() != self.time:
            logging.error('输入表格的期数不是汇总表的期数')
            print('输入表格的期数不是汇总表的期数')
            return
        # 如果输入表格的企业代码不在汇总表中，返回错误
        if company_code not in self.df['机构类型'].values:
            logging.error('输入表格的企业代码不在汇总表中')
            print('输入表格的企业代码不在汇总表中')
            return
        # 如果输入表格的企业代码在汇总表中，将纳税申报表并到汇总表中
        self.df.loc[self.df['机构类型'] == company_code,
                    self.time + '增值税'] = tax_table.get_vat()
        self.df.loc[self.df['机构类型'] == company_code,
                    self.time + '企业所得税'] = tax_table.get_income_tax()
        self.df.loc[self.df['机构类型'] == company_code,
                    self.time + '工资和五险一金'] = tax_table.get_insurance()

    # 将财务报表并到汇总表中
    def merge_financial_table(self, financial_table):
        company_code = financial_table.get_company_code()
        # 如果输入表格的期数不是汇总表的期数，返回错误
        if financial_table.get_time() != self.time:
            logging.error('输入表格的期数不是汇总表的期数')
            print('输入表格的期数不是汇总表的期数')
            return
        # 如果输入表格的企业代码不在汇总表中，返回错误
        if company_code not in self.df['机构类型'].values:
            logging.error('输入表格的企业代码不在汇总表中')
            print('输入表格的企业代码不在汇总表中')
            return
        # 如果输入表格的企业代码在汇总表中，将财务报表并到汇总表中
        # 第一期营业收入	第一期营业成本	第一期销售费用	第一期管理费用	除工资以外的管理费用	第一期财务费用	第一期研发费用	第一期投资收益	第一期营业利润	第一期营业外收入	第一期营业外支出	第一期利润总额
        self.df.loc[self.df['机构类型'] == company_code,
                    self.time + '营业收入'] = financial_table.get_revenue()
        self.df.loc[self.df['机构类型'] == company_code,
                    self.time + '营业成本'] = financial_table.get_cost()
        self.df.loc[self.df['机构类型'] == company_code, self.time +
                    '销售费用'] = financial_table.get_sales_expense()
        self.df.loc[self.df['机构类型'] == company_code, self.time +
                    '管理费用'] = financial_table.get_management_expense()
        self.df.loc[self.df['机构类型'] == company_code, self.time +
                    '财务费用'] = financial_table.get_financial_expense()
        self.df.loc[self.df['机构类型'] == company_code, self.time +
                    '研发费用'] = financial_table.get_research_expense()
        self.df.loc[self.df['机构类型'] == company_code, self.time +
                    '投资收益'] = financial_table.get_investment_income()
        self.df.loc[self.df['机构类型'] == company_code, self.time +
                    '营业外收入'] = financial_table.get_non_operating_income()
        self.df.loc[self.df['机构类型'] == company_code, self.time +
                    '营业外支出'] = financial_table.get_non_operating_expense()

    # 将汇总表保存到excel中

    def save(self):
        # 构造索引名与excel中的列名的对应关系
        columns_name = {
            '序号': 0,
            '审核': 1,
            '机构类型': 2,
            '公司名称': 3,
            '第一期营业收入': 4,
            '第一期营业成本': 5,
            '第一期销售费用': 6,
            '第一期管理费用': 7,
            '除工资以外的管理费用': 8,
            '第一期财务费用': 9,
            '第一期研发费用': 10,
            '第一期投资收益': 11,
            '第一期营业利润': 12,
            '第一期营业外收入': 13,
            '第一期营业外支出': 14,
            '第一期利润总额': 15,
            '第一期企业所得税': 16,
            '第一期工资和五险一金': 17,
            '第一期增值税': 18,
            '合计': 19,
            '第一期增值税税负': 20,
            '真人工资和五险一金': 21,
            '虚拟五险一金': 22
        }
        df = self.df[[self.time + '营业收入', self.time + '营业成本', self.time + '销售费用', self.time + '管理费用', self.time + '财务费用', self.time + '研发费用',
                      self.time + '投资收益', self.time + '营业外收入', self.time + '营业外支出', self.time + '增值税', self.time + '工资和五险一金', self.time + '企业所得税']]
        df = df.fillna(0)
        # 获取df中列记录的列表
        index = [index + 2 for index in df.index.tolist()]
        columns = [columns_name[column] + 1 for column in df.columns.tolist()]
        wb = openpyxl.load_workbook('税务局汇总表.xlsx')
        ws = wb[self.time + '企业缴税汇总表']
        # 使用openpyxl写入数据
        for i in range(len(index)):
            for j in range(len(columns)):
                ws.cell(row=index[i], column=columns[j], value=df.iloc[i, j])
                print(index[i], columns[j], df.iloc[i, j])
        wb.save('税务局汇总表.xlsx')


# 初始化
generate_log()
build_dir()
copy_file()
rename_file()
xls_to_xlsx()
# delete_all_file()
delete_file()
# logging.debug('处理完成')

copy_template()

summary_table = summary_table('第一期')
tax_table = tax_table('制造01','第一期')
finally_table = finance_table('制造01','第一期')
summary_table.merge_tax_table(tax_table)
summary_table.merge_financial_table(finance_table)
# summary_table.save()


# # 根据期数和公司代码进行合并流程
# for time in time_list:
#     summary_table = summary_table(time)
#     for company_code in company_code_list[0:1]:
#         try:
#             tax_table = tax_table(company_code, time)
#             finance_table = finance_table(company_code, time)
#         except:
#             pass
#         summary_table.merge_tax_table(tax_table)
#         # summary_table.merge_finance_table(finance_table)
#     # summary_table.save()

