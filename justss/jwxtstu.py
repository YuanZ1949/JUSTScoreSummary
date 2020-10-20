# coding: utf-8

from requests import session
from bs4 import BeautifulSoup 
from pandas import read_html, DataFrame, set_option, isnull
from decimal import Decimal
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
from traceback import print_exc
from .table_working import *

set_option('mode.chained_assignment', None)

class Jwxtstu():
    """
    这是个教务系统学生类
    """
    def __init__(self, stu_id='', stu_password=''):
        self.stu_name = ''
        self.stu_id = stu_id 
        self.stu_password = stu_password
        self.score_table = ''
        self.all_score_table = ''
        self.public_elective_list = []
        self.course_replacement_table = ''
        self.course_replacement_dict = {}
        self.now_no_pass_table = ''
        self.training_program_table = ''
        self.point_dict = {}
        self.jwxt_session = session()
        self.userAgent_PC = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3956.0 "
            "Safari/537.36 Edg/80.0.328.4")
        if stu_id != '' and stu_password != '':
            self.login_jwxt()
            self.stu_info()
            self.score()
            self.training_program()
            self.course_replacement()
            self.replacement_dict()
            self.now_no_pass()
            self.public_elective()
            self.point_summary()
            self.now_no_pass()

    def login_jwxt(self, account='', password=''):
        if account == '' or password =='' : 
            account = self.stu_id
            password = self.stu_password

        self.jwxt_session.get("http://jwgl.just.edu.cn:8080/jsxsd/")
        
        postUrl = "http://jwgl.just.edu.cn:8080/jsxsd/xk/LoginToXk"
        postData = {
            'USERNAME' : account ,
            'PASSWORD' : password }
        header = {
            'Host' : "jwgl.just.edu.cn:8080" ,
            'Referer' : "http://jwgl.just.edu.cn:8080/jsxsd/" ,
            'User-Agent' : self.userAgent_PC ,
            'Cache-Control': 'max-age=0',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests' : '1' }
        main_page = self.jwxt_session.post(postUrl, data = postData, headers = header)
        main_page_content = BeautifulSoup(main_page.content.decode('utf-8'), 'lxml')
        stu_info = main_page_content.find(class_='block1text').text.split('\n')
        self.stu_name = stu_info[1][stu_info[1].find('：')+1:stu_info[1].find(' \r')]
        return self.stu_name

    def stu_info(self):
        if self.stu_name == '': 
            try:
                self.login_jwxt()
            except:
                print('请使用jwxt_login()登录后尝试')
                return
        print('姓名:', self.stu_name)
        print('学号:', self.stu_id)
        return [self.stu_name, self.stu_id]

    def score(self, kksj='', kcxz='', kcmc='', xsfs='all'):
        '''
        成绩获取方法(共四个参数)
        
        [开课时间]: 格式举例: '2019-2020-1'
        [课程性质]: '公共课', '基础课', '公共基础课', '学科基础课', '专业课', '副修专业基础课程', '副修专业学位附加课程', '工程基础课'
        [课程名称]: 格式举例: '大学物理A1'
        [显示方式]: '显示全部成绩', '显示最好成绩' 这里第一个只需要传入'all' 第二个传入'max'
        '''
        post_url = 'http://jwgl.just.edu.cn:8080/jsxsd/kscj/cjcx_list'
        postData = {
            'kksj': kksj, # 开课时间
            'kcxz': kcxz, # 课程性质
            'kcmc': kcmc, # 课程名称
            'xsfs': xsfs  # 显示方式
        }
        header = {
            'Host' : "jwgl.just.edu.cn:8080" ,
            'Referer' : "http://jwgl.just.edu.cn:8080/jsxsd/kscj/cjcx_query" ,
            'User-Agent' : self.userAgent_PC ,
            'Cache-Control': 'max-age=0',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests' : '1' ,
        }
        result = self.jwxt_session.post(post_url, data = postData, headers = header)
        if kksj=='' and kcxz=='' and kcmc=='' and xsfs=='all':
            self.all_score_table = html_table(result.content)
            return self.all_score_table
        else:
            self.score_table = html_table(result.content)
            return self.score_table

    def theory_schedule(self, file_name = ''):
        '''
        直接返回Pandas数据帧形式的学期课表,如果传入文件名,会保存成对应文件名的文件.
        '''
        get_url = 'http://jwgl.just.edu.cn:8080/jsxsd/xskb/xskb_list.do'
        if file_name == '':
            return html_table(self.jwxt_session.get(get_url).content)
        else:
            theory_schedule_table = html_table(self.jwxt_session.get(get_url).content)
            table_xls(theory_schedule_table, file_name)
            return theory_schedule_table

    def public_elective(self, file_name=''):
        '''
        返回公选课数据(公选课总表,A类表,B类表)
            1.找出 所有成绩表A(all_score_table)内的带有'公选'关键词的记录
            2.按照'A','B','科技','人文'等关键词进行分列
        '''
        # 公选课总表
        self.public_elective_table = yes_pass(table_finder(self.all_score_table, '课程名称', ['校公选','校公共选']),'成绩')
        # A类公选课
        pe_A_condition = ['(校公选课人文类)','(校公共选修课人文、艺术类)','(校公选A类/人文)',
                            '(校公选A类/经管)','(校公选A类/人文、艺术)','（校公选A类/人文）',
                            '（校公选人文艺术类）','（校公选社会科学类）','（校公选课人文类）',
                            '（校公共选修课人文、艺术类）','（校公选A类/经管）','（校公选A类/人文、艺术）',
                            '(校公选人文艺术类)','(校公选社会科学类)','(校公选A类/人文艺术)']
        pe_A_table = table_finder(self.public_elective_table, '课程名称', pe_A_condition)
        # B类公选课
        pe_B_condition = ['(校公选B类/科技)','(校公选B类/科学技术)','（校公选自然科学类）',
                            '（校公选B类/科学技术）','(校公选自然科学类)','(校公选工程技术类)',
                            '（校公选工程技术类）','（校公选B类/科技）']
        pe_B_table = table_finder(self.public_elective_table, '课程名称', pe_B_condition)

        if file_name != '' : table_xls(self.public_elective_table, file_name)
        self.public_elective_list = [self.public_elective_table, pe_A_table, pe_B_table]
        return self.public_elective_list

    def course_replacement(self, file_name = ''):
        '''
        课程替代获取
        '''
        get_url = 'http://jwgl.just.edu.cn:8080/jsxsd/xkgl/tsqkxk_list'
        self.course_replacement_table = read_html(self.jwxt_session.get(get_url).content, encoding='utf-8', header=0)[1]
        if file_name == '':
            return self.course_replacement_table
        else:
            table_xls(self.course_replacement_table,file_name)

    def replacement_dict(self):
        '''
        返回所有已经被批准的课程替代的课程的课程号组成的字典
        '''
        cr_table = self.course_replacement_table
        cr_table = cr_table.loc[cr_table.loc[:,'审核状态'].isin(['通过'])]
        # cr_table
        cr_table_num = cr_table.shape[0]
        cr_table_dict = {}
        for row_num in range(cr_table_num):
            cr_table_dict[cr_table.iloc[row_num,7]] = cr_table.iloc[row_num,2]
        self.course_replacement_dict = cr_table_dict
        return cr_table_dict

    def now_no_pass(self):
        '''
        列出仍然未过的课程(挂科,补考,重修等所有数据)
            1.获取历史挂科成绩表A(all_no_pass)/及格成绩表B(all_yes_pass)/课程替代字典C(course_replacement_dict)
            2.依A表查B表,B表内有A表内的课程说明已经通过,跳过
            3.B表内没有,查课程替代字典C(course_replacement_dict),有就重复2步骤,没有就是还没过的
        '''
        # 逻辑步骤1
        # score() 获取所有成绩
        all_no_pass = no_pass(self.all_score_table, '成绩')   # 表A
        all_yes_pass = yes_pass(self.all_score_table, '成绩') # 表B
        # course_replacement_dict 字典C
        all_yes_pass_series = all_yes_pass.loc[:,'课程号']
        all_no_pass_num = all_no_pass.shape[0]
        now_no_pass_list = []
        
        for row_num in range(all_no_pass_num):
            replace_id = ''
            row_class_id = all_no_pass.iloc[row_num,2]
            if True in all_yes_pass_series.isin([row_class_id]).values: 
                continue # 逻辑步骤2
                
            # 逻辑步骤3
            replace_id = self.course_replacement_dict.get(row_class_id)
            if replace_id:
                if True in all_yes_pass_series.isin([replace_id]).values: # 重复逻辑步骤2
                    continue 

            all_no_pass.iat[row_num,0] = row_num+1
            if all_no_pass.iloc[row_num].loc['课程名称'].find('公共选') != -1:
                continue
            if all_no_pass.iloc[row_num].loc['课程名称'].find('公选') != -1:
                continue
            now_no_pass_list.append(all_no_pass.iloc[row_num])

        self.now_no_pass_table = DataFrame(now_no_pass_list)
        return self.now_no_pass_table

    def training_program(self, file_name = ''):
        '''
        培养方案获取
        '''
        get_url = 'http://jwgl.just.edu.cn:8080/jsxsd/pyfa/pyfazd_query'
        self.training_program_table = html_table(self.jwxt_session.get(get_url).content)
        if file_name == '':
            return self.training_program_table
        else:
            table_xls(self.training_program_table, file_name)

    def add_academic_credits(self, table_df, ignore=True):
        '''
        在传入的表的最右侧加一列`学分绩点`,并返回修改过的表,表格内的平均绩点,总学分
        ignore参数为在计算平均绩点和总学分的时候是否忽略公选课,体育课和补考通过的课程
        绩点计算: http://jwc.just.edu.cn/2018/0328/c5744a51661/page.htm
        '''
        table_num = table_df.shape[0]
        grade_point_sum = Decimal(0)   # 学分绩点
        credit_sum = Decimal(0)        # 总学分
        grade_point_list = []
        repeat_list = []               # 用于去重每一表格内的重复课程(得到最高分课程)
        for row_num in range(table_num):
            score  = table_df.iloc[row_num].loc['成绩']
            row_class = table_df.iloc[row_num].loc['课程号']
            credit = table_df.iloc[row_num].loc['学分']
            if isnull(credit): 
                get_table = table_filter(self.all_score_table, '课程号', [table_df.iloc[row_num].loc['课程号']])
                credit = get_table.iloc[0].loc['学分']
            credit = Decimal(credit)

            # 计算绩点
            point  = Decimal(0) # 绩点
            score_list = ['优','良','中','及格','不及格','通过','不通过']
            if score in score_list:
                score_dict = {
                    '优':4.5,'良':3.5,'中':2.5,'及格':1.5,'不及格':0.0,
                    '通过':2.5,'不通过':0.0 }
                point = Decimal(score_dict[score])
            else:
                float(score)
                point = (Decimal(score)/Decimal(10)) - Decimal(5)
                if point < Decimal(1): point = Decimal(0)

            grade_point_list.append(credit*point)
            if point == Decimal('0'):  # 如果绩点不及格尝试寻找合格成绩进行计算
                all_yes_pass = yes_pass(self.all_score_table, '成绩')
                filter_list = [row_class, self.course_replacement_dict.get(row_class)]
                if table_filter(all_yes_pass, '课程号', filter_list).empty == False:
                    continue
            if (table_df.iloc[row_num].loc['课程名称'].find('体育') != -1) and ignore: 
                continue
            if (table_df.iloc[row_num].loc['课程名称'].find('公选') != -1) and ignore: 
                continue
            if (table_df.iloc[row_num].loc['课程名称'].find('公共选') != -1) and ignore: 
                continue
            if row_class in repeat_list:
                continue
            repeat_list.append(row_class)
            grade_point_sum = grade_point_sum + credit*point
            credit_sum = credit_sum + credit
            
        table_df.loc[:,'学分绩点'] = grade_point_list
        mean_grade_point = grade_point_sum/credit_sum
        return [table_df, eval(str(mean_grade_point)), eval(str(credit_sum))]

    def point_summary(self):
        '''
        返回绩点字典
        
        字典数据结构说明
        第一层 point_dict{年份:[...],...} 
                                ^ 
        第二层 [年平均绩点,[第一学期成绩],[第二学期成绩]]
                                ^               ^
        第三层 [DF成绩表,学期平均绩点,学期总学分]  # 注意绩点和学分未计入体育和公选课
        '''
        term_list = self.all_score_table.loc[:,'开课学期'].drop_duplicates().values.tolist() # 获取学期列表
        year_list = ()
        term_dict = {}
        term_num = 0
        for each_term in term_list:
            term_num = term_num+1
            year_list = year_list + (each_term[:-2],)
            
            # 由培养方案获取各学期绩点计算表格
            std_class_table = table_filter(self.training_program_table, '开设学期', [term_num])
            std_class_list = list(std_class_table.loc[:,'课程号'])
            self.course_replacement_dict = self.course_replacement_dict
            for origin_class,replacement_class in self.course_replacement_dict.items():
                if origin_class in std_class_list: std_class_list.append(replacement_class)
            each_term_table = table_filter(self.all_score_table, '课程号', std_class_list)
            
            # 单学期绩点计算
            each_term_list = self.add_academic_credits(each_term_table)
            term_dict[each_term] = [each_term_list[0],each_term_list[1],each_term_list[2]]

        self.point_dict = {}
        year_list = list(year_list) 
        all_year_grade_point = Decimal('0')
        all_year_credit = Decimal('0')
        for each_year in year_list:
            grade_point_1 = Decimal(term_dict[each_year+'-1'][1])*Decimal(term_dict[each_year+'-1'][2])
            grade_point_2 = Decimal(term_dict[each_year+'-2'][1])*Decimal(term_dict[each_year+'-2'][2])
            all_credit = Decimal(term_dict[each_year+'-1'][2])+Decimal(term_dict[each_year+'-2'][2])
            year_mean_point = (grade_point_1 + grade_point_2) / all_credit
            year_mean_point = eval(str(year_mean_point))

            self.point_dict[each_year] = [year_mean_point,term_dict[each_year+'-1'],term_dict[each_year+'-2']]

            all_year_grade_point = all_year_grade_point+grade_point_1+grade_point_2
            all_year_credit = all_year_credit+all_credit
            
        all_year_mean_point = all_year_grade_point / all_year_credit
        all_year_mean_point = eval(str(all_year_mean_point))
        self.point_dict['all_mean_point'] = all_year_mean_point
        return self.point_dict

    def generate_summary(self, file_name = ''):
        '''
        传入学生对象, 总结表的存储位置与名字, 生成总结表
        '''
        def table_openpyxl(table_df, sheet):
            for r in dataframe_to_rows(table_df, index=False, header=True):
                sheet.append(r)
            return sheet

        def next_row(add_num=1):
            '''
            传入数字,最终传出的数字解读为"最后一行的下'add_num'行"
            '''
            # sheet.dimensions.split(':')[] 获取当前填充了数据的最大列和行
            row_id = sheet.dimensions.split(':')[1]
            return str(eval(row_id[1:])+add_num)
        
        def add_style_row():
            pass
        
        wb = Workbook()
        sheet = wb.active

        if self.stu_name == '' and self.stu_id == '': 
            print('请使用jwxt_login()登录后尝试')
            return
        
        sheet.title = self.stu_name

        # stu_name, stu_id
        font1 = Font(name="宋体",size=14,bold=True,italic=False,color="000000")
        font2 = Font(name="宋体",size=12,bold=True,italic=False,color="000000")  

        sheet['A'+next_row(1)] = '学生成绩统计概览'
        sheet.row_dimensions[eval(next_row(0))].height = 30
        sheet['A'+next_row(0)].font = font1
        sheet['A'+next_row(0)].font = Font(name="宋体",size=24,bold=True,italic=False,color="000000")
        
        sheet['A'+next_row(2)] = '姓名:'
        sheet['A'+next_row(0)].font = font1
        sheet['B'+next_row(0)] = self.stu_name
        
        sheet['A'+next_row(1)] = '学号:'
        sheet['A'+next_row(0)].font = font1
        sheet['B'+next_row(0)] = str(self.stu_id)

        # point_dict
        sheet['A'+next_row(2)] = '各学年平均绩点, 学期平均绩点'
        sheet['A'+next_row(0)].font = font1
        all_mean_point = self.point_dict.get('all_mean_point')
        sheet['A'+next_row(2)] = f'总平均绩点:{roundoff(all_mean_point)}'
        sheet['A'+next_row(0)].font = font2
        for year,content in self.point_dict.items():
            if type(content) == float : continue # 跳过`all_mean_point`这一项数据
            sheet['A'+next_row(2)] = f'学年:{year}'
            sheet['A'+next_row(0)].font = font2
            sheet['C'+next_row(1)] = f'学年平均绩点:{roundoff(content[0])}'
            sheet['B'+next_row(2)] = f'学期:{year}-1'
            sheet['D'+next_row(0)] = f'学期平均绩点:{roundoff(content[1][1])}'  
            sheet['B'+next_row(2)] = f'学期:{year}-2'
            sheet['D'+next_row(0)] = f'学期平均绩点:{roundoff(content[2][1])}'
            sheet['A'+next_row(1)] = ''
        
        # now_no_pass_table
        sheet['A'+next_row(2)] = '尚未通过的选修过的课程'
        sheet['A'+next_row(0)].font = font1
        sheet['A'+next_row(1)] = ''
        if self.now_no_pass_table.shape[0] != 0: 
            sheet.sheet_properties.tabColor = 'EE3131' # 标红工作表标签
            sheet = table_openpyxl(self.now_no_pass_table, sheet)
        else:
            sheet.sheet_properties.tabColor = '007979' # 标绿工作表标签
            sheet['A'+next_row(1)] = '所修皆已通过~'
        
        # self.public_elective_list
        sheet['A'+next_row(2)] = f'A类公选课'
        sheet['C'+next_row(0)] = f'已修学分:{self.add_academic_credits(self.public_elective_list[1],False)[2]}'
        sheet['A'+next_row(0)].font = font1
        sheet['C'+next_row(0)].font = font1
        sheet['A'+next_row(1)] = ''
        sheet = table_openpyxl(self.public_elective_list[1], sheet)
        
        sheet['A'+next_row(2)] = f'B类公选课'
        sheet['C'+next_row(0)] = f'已修学分:{self.add_academic_credits(self.public_elective_list[2],False)[2]}'
        sheet['A'+next_row(0)].font = font1
        sheet['C'+next_row(0)].font = font1
        sheet['A'+next_row(1)] = ''
        sheet = table_openpyxl(self.public_elective_list[2], sheet)
        
        # score_table
        yes_pass_list = self.add_academic_credits(yes_pass(self.all_score_table,'成绩'), False)    
        sheet['A'+next_row(2)] = '已通过的所有课程'
        sheet['D'+next_row(0)] = f'已获得的总学分: {yes_pass_list[2]}'
        sheet['A'+next_row(0)].font = font1
        sheet['D'+next_row(0)].font = font1
        sheet['A'+next_row(1)] = ''
        sheet = table_openpyxl(yes_pass_list[0], sheet)
        #yes_pass_credit_sum

        # point_dict
        sheet['A'+next_row(2)] = '各学年平均绩点, 总平均绩点'
        sheet['A'+next_row(0)].font = font1
        all_mean_point = self.point_dict.get('all_mean_point')
        sheet['A'+next_row(2)] = f'总平均绩点:{roundoff(all_mean_point)}'
        sheet['A'+next_row(0)].font = font2
        for year,content in self.point_dict.items():
            if type(content) == float : continue # 跳过每学年的`all_mean_point`这一项数据
            sheet['A'+next_row(2)] = f'学年:{year}'
            sheet['C'+next_row(0)] = f'学年平均绩点:{roundoff(content[0])}'
            sheet['A'+next_row(0)].font = font2
            sheet['C'+next_row(0)].font = font2
            
            sheet['A'+next_row(1)] = f'学期:{year}-1'
            sheet['C'+next_row(0)] = f'学期平均绩点:{roundoff(content[1][1])}'  
            sheet['A'+next_row(0)].font = font2
            sheet['C'+next_row(0)].font = font2
            sheet['A'+next_row(1)] = ''
            sheet = table_openpyxl(content[1][0], sheet)
            
            sheet['A'+next_row(2)] = f'学期:{year}-2'
            sheet['C'+next_row(0)] = f'学期平均绩点:{roundoff(content[2][1])}'
            sheet['A'+next_row(0)].font = font2
            sheet['C'+next_row(0)].font = font2       
            sheet['A'+next_row(1)] = ''
            sheet = table_openpyxl(content[2][0], sheet)
            
            sheet['A'+next_row(1)] = ''

        # 调整列宽
        sheet.column_dimensions["B"].width = 12
        sheet.column_dimensions["C"].width = 10
        sheet.column_dimensions["D"].width = 29
        sheet.column_dimensions["J"].width = 16
        sheet.column_dimensions["K"].width = 10
        sheet.column_dimensions["L"].width = 12
        sheet.column_dimensions["M"].width = 8
        if file_name == '':
            wb.save(filename = f"{self.stu_id}-{self.stu_name}-Summary.xlsx")
        else:
            wb.save(file_name+'.xlsx')

def roundoff(number):
    '''
    传入数字, 传出四舍五入的保留两位小数的数字
    '''
    if type(number) == float:
        number = Decimal(str(number))
    return number.quantize(Decimal('0.00'),'ROUND_HALF_UP')

def main(stu_id, stu_password):
    '''
    循环执行函数
    '''
    try:
        stu = Jwxtstu(stu_id, stu_password)
        print()
        print(f'姓名:{stu.stu_name}\n学号:{stu.stu_id}')

        print('正在汇总所有数据ing\n')
        stu.generate_summary()

        print('搞定^_^ ~ ')
        print('获取的文件已经保存在此运行文件所在的文件夹中')

    except:
        print_exc()
        print()
        print('出现如上错误,请确保账号可用以及网络正常')
        print('然后重试或者联系作者')

    finally:
        print()
        print('程序会自动退出 有问题找苑长反馈哈 504567327@qq.com')
           
if __name__ == '__main__':
    
    import os

    # md = os.sys.argv[1]  
    print()
    print('载入中ing~')
    print()
    
    print('输入完数据后记得回车喔~')
    print()
    account = int(input('请输入学号: '))
    password = input('请输入密码: ')
        
    main(account,password)
    os.system('pause')