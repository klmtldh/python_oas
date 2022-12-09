from main_lib import *
from oas_pptx import *
from oas_docx import *
import oas_n_day
import datetime  # 用于日期操作
import time
import sys

def sfb_plan_save_path(file_type):
    # 模块用时计时
    start_time = time.time()
    m_name = '会议'
    if file_type == 'docx':
        f_name = '昆明供电局安全生产分析会安监部材料docx'
    elif file_type == 'pptx':
        f_name = '昆明供电局安全生产分析会安监部材料pptx'
    io_name = '输出'
    file_map = read_file_map(m_name, f_name, io_name)

    elapsed_time = "{0}模块用时{1}秒".format(sys._getframe().f_code.co_name, time.time() - start_time)
    print(elapsed_time)

    return file_map[0]


def sfb_meeting_update():
    # 模块用时计时
    start_time = time.time()

    # 更新“昆明供电局安全监管重点工作计划”到数据库
    m_name = '工作计划'
    f_name = '昆明供电局安全监管重点工作计划'
    io_name = '输入'
    file_map = read_file_map(m_name, f_name, io_name)
    update_excel(file_map[0], file_map[1], file_map[3])
    elapsed_time = "{0}模块用时{1}秒".format(sys._getframe().f_code.co_name, time.time() - start_time)
    print(elapsed_time)


def sfb_meeting_backup():
    # 将“昆明供电局安全监管重点工作计划”数据库备份到本地
    m_name = '工作计划'
    f_name = '昆明供电局安全监管重点工作计划（数据库备份）'
    io_name = '输出'
    file_map = read_file_map(m_name, f_name, io_name)
    backup_mysql(file_map[3], file_map[0], file_map[1])


def write_pptx():
    now_day = oas_n_day.NDay(datetime.date.today())

    op = OasPptx(r'D:\JGY\600-Data\004-auxiliary辅助文件\南方电网logo（16，9）模板.pptx')

    first_page_text_list = ['{0}目标指标总体情况及{1}安全生产工作情况'.format(now_day.get_chinese_date()[0],
                                                             now_day.get_chinese_date()[4]),
                            '安全监管部（应急指挥中心）',
                            now_day.get_chinese_date()[3]]

    op.alter_csg_logo_first_page(first_page_text_list)
    contents_page_text_list = ['一、安全生产目标指标情况',
                               '二、上月布置工作推进情况',
                               '三、本月重点工作推进情况',
                               '四、下月需要重点关注的工作'
                               ]
    op.text_n('目录', font_bold=True)
    op.text_n(contents_page_text_list, textbox_top=12)
    op.page()
    op.text('1.1总体回顾')

    op.save_pptx(sfb_plan_save_path('pptx')+str(datetime.date.today())+'.pptx')
    print(sfb_plan_save_path('pptx')+str(datetime.date.today())+'.pptx')


# 计划分析模块
def sfb_meeting_main():
    # 用Excel文件更新MySQL数据库
    #sfb_meeting_update()
    # 将MySQL中的表备份到本地Excel
    #sfb_meeting_backup()
    # 统一查询库
    sql = "select * from 昆明供电局安全监管重点工作计划"
    # Word文档自动生成
    #write_docx(sql)
    # PPT文档自动生成
    write_pptx()
