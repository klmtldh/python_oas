from main_lib import *
from oas_docx import *
from oas_pptx import *
import oas_n_day
import datetime  # 用于日期操作
import time
import sys


def sfb_person_save_path():
    m_name = '安监人员'
    f_name = '昆明供电局安监部人员情况'
    io_name = '输出'
    file_map = read_file_map(m_name, f_name, io_name)
    return file_map[0]


def sfb_person_update():
    # 更新“昆明供电局安监部人员信息”到数据库
    m_name = '安监人员'
    f_name = '昆明供电局安监部人员信息'
    io_name = '输入'
    file_map = read_file_map(m_name, f_name, io_name)
    update_excel(file_map[0], file_map[1], file_map[3])

    # 更新“昆明供电局安监部岗位设置”到数据库
    m_name = '安监人员'
    f_name = '昆明供电局安监部岗位设置'
    io_name = '输入'
    file_map = read_file_map(m_name, f_name, io_name)
    update_excel(file_map[0], file_map[1], file_map[3])


def sfb_person_backup():
    # 将“昆明供电局安监部人员信息”数据库备份到本地
    m_name = '安监人员'
    f_name = '昆明供电局安监部人员信息（数据库备份）'
    io_name = '输出'
    file_map = read_file_map(m_name, f_name, io_name)
    backup_mysql(file_map[3], file_map[0], file_map[1])

    # 将“昆明供电局安监部岗位设置”数据库备份到本地
    m_name = '安监人员'
    f_name = '昆明供电局安监部岗位设置（数据库备份）'
    io_name = '输出'
    file_map = read_file_map(m_name, f_name, io_name)
    backup_mysql(file_map[3], file_map[0], file_map[1])


def redundant_vacancies(sql_person, sql_value):
    # 连接数据库，读取人员信息和人员配置信息
    dbm = DataFrameBothMysql()
    df_person = dbm.select_mysql(sql_person)
    df_value = dbm.select_mysql(sql_value)

    # 统一人员配置数据
    df_value = df_value[['科室或班组名称', '岗位名称', '职数']]
    df_value.rename(columns={'科室或班组名称': '科室', '岗位名称': '岗位'}, inplace=True)
    df_value1 = df_value.copy()
    df_value1.loc[:, '职数'] = df_value1.loc[:, '职数'].astype('int')

    # 对人员表进行透视
    df_person_table = pd.pivot_table(df_person,
                                     index=['科室', '岗位'],
                                     columns=['岗位状态'],
                                     values=['姓名'],
                                     aggfunc=len,
                                     fill_value=0,
                                     margins=True,
                                     margins_name='总数'
                                     )
    df_person_table.reset_index(inplace=True)
    df_person_table.columns = ['科室', '岗位', '借入', '借出', '在岗', '总数']

    df_table = pd.merge(df_value1, df_person_table, how='left', on=['科室', '岗位'])
    df_table['缺员'] = df_table['职数'] - df_table['总数']

    text = "昆明供电局安全监管部（应急指挥中心）定员{0}人，\
实际{1}人，缺员{2}人,其中：在岗{3}人，借入{4}人，借出{5}人。".format(
        df_table['职数'].sum(),
        df_table['总数'].sum(),
        df_table['缺员'].sum(),
        df_table['在岗'].sum(),
        df_table['借入'].sum(),
        df_table['借出'].sum()
    )
    analysis_columns = ['政治面貌', '性别', '民族', '文化程度', '职称']
    s_list = ['男', '女',
              '党员', '群众',
              '研究生', '本科', '大专',
              '高级工程师', '工程师', '助理工程师',
              '经理', '副经理', '主管',
              '安全监察专责(A)', '安全监察专责(B)', '风险体系管理专责(A)', '风险体系管理专责(B)', '应急管理专责',
              '班长', '安全监察员']
    text_list = []
    pa = PandasAnalysis(df_person)
    for item in analysis_columns:
        df_person_x = df_person[item].value_counts()
        df_person_x = pa.sort_list(df_person_x, s_list)
        df_person_x.name = item
        text_person_x = pa.pandas_text(df_person_x, unit='名')
        text_list.append(text_person_x)

    # 年龄
    df_person_age = pd.cut(df_person['年龄'].astype('int'), bins=[0, 30, 40, 50, 100],
                           labels=["30岁以下", "30-39岁", "40-49岁", "50岁以上"])
    df_person_age_count = df_person_age.value_counts()
    # df_person_age_count.index = ["30岁以下", "30-39岁", "40-49岁", "50岁以上",
    #                              '最小', '最大', '平均', '中位数', '方差']
    df_person_age_dscribe = pd.Series(data=[round(df_person['年龄'].astype(int).min(), 1),
                        round(df_person['年龄'].astype(int).max(), 1),
                        round(df_person['年龄'].astype(int).mean(), 1),
                        round(df_person['年龄'].astype(int).median(), 1),
                        round(df_person['年龄'].astype(int).std(), 1)
                        ], index=['最小', '最大', '平均', '中位数', '方差'])
    text_person_age_dscribe = pa.pandas_text(df_person_age_dscribe, unit='岁')
    df_person_age_count = pa.sort_list(df_person_age_count, ["30岁以下", "30-39岁", "40-49岁", "50岁以上"])
    df_person_age_count.name = '年龄'
    text_person_age_count = pa.pandas_text(df_person_age_count, unit='名', punctuation=[',', ','])+text_person_age_dscribe



    # 工龄
    df_person_wage = pd.cut(df_person['工龄'].astype('int'), bins=[0, 5, 10, 20, 30, 50],
                           labels=["5年以下", "5-9年", "10-19年","20-29年", "30年以上"])
    df_person_wage_count = df_person_wage.value_counts()
    df_person_wage_dscribe = pd.Series(data=[round(df_person['工龄'].astype(int).min(), 1),
                                            round(df_person['工龄'].astype(int).max(), 1),
                                            round(df_person['工龄'].astype(int).mean(), 1),
                                            round(df_person['工龄'].astype(int).median(), 1),
                                            round(df_person['工龄'].astype(int).std(), 1)
                                            ], index=['最小', '最大', '平均', '中位数', '方差'])
    text_person_wage_dscribe = pa.pandas_text(df_person_wage_dscribe, unit='年')
    df_person_wage_count = pa.sort_list(df_person_wage_count, ["5年以下", "5-9年", "10-19年","20-29年", "30年以上"])
    df_person_wage_count.name = '工龄'
    text_person_wage_count = pa.pandas_text(df_person_wage_count,
                                            unit='名',
                                            punctuation=[',', ','])+text_person_wage_dscribe

    # 字符
    text = text+''.join(text_list) + text_person_age_count + text_person_wage_count

    df_person = df_person[['姓名', '科室', '岗位', '兼任职务', '联系电话', '政治面貌', '工号', '性别', '岗位状态']]
    df_person.fillna('无', inplace=True)
    list_sort = [['部门领导', '安全监察科', '应急与保供电管理科', '安全督查大队'],
                 ['经理', '副经理', '主管', '安全监察专责(A)', '安全监察专责(B)',
                  '应急管理专责', '风险体系管理专责(A)', '风险体系管理专责(B)',
                  '班长', '安全监察员']]
    df_person = pa.sort_list(df_person, list_sort, ['科室', '岗位'])
    df_person.reset_index(drop=True, inplace=True)

    return df_table, text, df_person


def write_docx(sql, sql_value):
    doc = OasDocx()
    doc.hd('昆明供电局安全监管部（应急指挥中心）人员情况', font_size=22)
    correct_tex = "昆明供电局安全监管部（应急指挥中心）人员情况共有{0}条数据，以下数据不完善：{1}"\
        .format(analysis_null(sql)[0], analysis_null(sql)[3])
    doc.par('一、数据总体情况', bold=True)
    doc.par(correct_tex)

    doc.par('二、人员总体情况', bold=True)

    df, text_person, df_person = redundant_vacancies(sql, sql_value)
    doc.par(text_person)
    doc.par('附件：安监部人员在岗情况表')
    doc.tb(len(df) + 1, len(list(df)))
    col_list = list(df)
    for i in range(len(df) + 1):
        for j in range(len(list(df))):
            if i == 0:
                doc.tb_cell(col_list[j], i, j, f_name='黑体')
            else:
                doc.tb_cell(str(df.iloc[i - 1, j]), i, j)

    doc.par('附件：安监部人员信息表')

    doc.tb(len(df_person) + 1, len(list(df_person)))
    col_list = list(df_person)
    for i in range(len(df_person) + 1):

        for j in range(len(list(df_person))):

            if i == 0:
                doc.tb_cell(col_list[j], i, j, f_name='黑体')

            else:
                doc.tb_cell(str(df_person.iloc[i - 1, j]), i, j)

    doc.save_docx(sfb_person_save_path())


# 安监部人员分析模块
def sfb_person_main():
    # 用Excel文件更新MySQL数据库
    sfb_person_update()
    # 将MySQL中的表备份到本地Excel
    sfb_person_backup()
    # 统一查询库
    sql = "select * from 昆明供电局安监部人员信息 WHERE 岗位状态!='调离'"
    sql_value = "select * from 昆明供电局安监部岗位设置"
    # Word文档自动生成
    write_docx(sql, sql_value)

