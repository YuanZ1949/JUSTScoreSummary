
from requests import session
from bs4 import BeautifulSoup 
from pandas import read_html, DataFrame, set_option, isnull
from decimal import Decimal
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
from traceback import print_exc



def html_table(html_content, num = 0):
    '''
    传入原始网页数据,提取网页内的表格数据,传出DataFrame形式的表格
    '''
    table_df = read_html(html_content, encoding='utf-8', header=0)[num]
    return table_df

def table_xls(table_df, file_name):
    '''
    传入DataFrame形式的表格并保存成指定名字的.xls文件
    '''
    try:
        table_df.to_excel(f'{file_name}.xlsx', index=False)
    except:
        print('保存失败')

def html_xls(html_content, file_name, num=0):
    '''
    传入原始网页数据, 保存网页内的表格为指定名字的.xls文件
    '''
    try:
        table_df = html_table(html_content, num)
        table_xls(table_df, file_name)
    except:
        print('转换失败')

def table_finder(table_df, column, condition_list):
    '''
    传入DF数据帧表,想要检索的列名,要含有的内容列表,传回`含有`指定字符的DF数据帧
    '''
    table_num = table_df.shape[0]
    table_df_column = table_df.loc[:,column]
    table_finder_list = []
    for row_num in range(table_num):
        for each_condition in condition_list:
            if table_df_column.iloc[row_num].find(each_condition) != -1:
                table_finder_list.append(table_df.iloc[row_num])
            
    table_finder_table = DataFrame(table_finder_list)
    return table_finder_table

def table_filter(table_df, column, condition_list):
    '''
    全字匹配传入函数的表格指定列的每个格,返回符合条件的行组成的数据帧
    '''
    return table_df.loc[table_df.loc[:,column].isin(condition_list)]

def yes_pass(table_df, column):
    '''
    传入表格和要筛选的列名, 返回指定列成绩为合格的行
    '''
    condition_list = [str(x/100) for x in range(6000,10001)] + ['及格','通过','优','良','中']
    condition_list = condition_list + [str(x) for x in range(60,101)]
    return table_filter(table_df, column, condition_list)

def no_pass(table_df, column):
    '''
    传入表格和要筛选的列名, 返回指定列成绩为不合格的行
    '''
    condition_list = [str(x/100) for x in range(0,6000)] + ['不及格','未通过']
    condition_list = condition_list + [str(x) for x in range(0,60)]
    return table_filter(table_df, column, condition_list)
