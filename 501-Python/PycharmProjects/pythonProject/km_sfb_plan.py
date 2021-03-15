from main_lib import *
from oas_pptx import *
from oas_docx import *
import oas_pandas_analysis
import oas_func
import oas_n_day
import datetime  # 用于日期操作
import time
import sys

def sfb_plan_save_path(file_type):
    # 模块用时计时
    start_time = time.time()
    m_name = '工作计划'
    if file_type == 'docx':
        f_name = '昆明供电局安全监管重点工作计划执行情况docx'
    elif file_type == 'pptx':
        f_name = '昆明供电局安全监管重点工作计划执行情况pptx'
    io_name = '输出'
    file_map = read_file_map(m_name, f_name, io_name)

    elapsed_time = "{0}模块用时{1}秒".format(sys._getframe().f_code.co_name, time.time() - start_time)
    print(elapsed_time)

    return file_map[0]


def sfb_plan_update():
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


def sfb_plan_backup():
    # 将“昆明供电局安全监管重点工作计划”数据库备份到本地
    m_name = '工作计划'
    f_name = '昆明供电局安全监管重点工作计划（数据库备份）'
    io_name = '输出'
    file_map = read_file_map(m_name, f_name, io_name)
    backup_mysql(file_map[3], file_map[0], file_map[1])


def sfb_plan_piovt_table(df, piovt_index, piovt_columns, piovt_values, table_columns, all_name=None):
    """
    piovt_index 透视表需要参与分类的列
    piovt_index 透视表需要统计的列
    piovt_values 透视表需要计算的列
    table_columns 透视表重新命名列名
    all_name  All更改名字
    """

    # 模块用时计时
    start_time = time.time()
    if all_name:
        # 建立基础透视表
        df_table = df.pivot_table(index=piovt_index,
                                  columns=piovt_columns,
                                  values=piovt_values,
                                  aggfunc=len,
                                  fill_value=0,
                                  margins=True
                                  )
        # 将index变为列
        df_table.reset_index(inplace=True)
        print(list(df_table))
        # 将更改列名
        if len(list(df_table)) < len(table_columns):
            df_table[('业务领域', '否')] = 0
            if len(table_columns) == 4:
                df_table = df_table[[(table_columns[0], ''), ('业务领域', '否'), ('业务领域', '是'), ('业务领域', 'All')]]
            else:
                df_table = df_table[[(table_columns[0], ''), (table_columns[1], ''), ('业务领域', '否'), ('业务领域', '是'), ('业务领域', 'All')]]
        print(df_table)
        df_table.columns = table_columns
        for item in piovt_index:
            df_table.loc[len(df_table) - 1, item] = all_name
        # 计算完成率
        df_table['完成率'] = df_table['是'] / df_table['合计']
        # 设置完成率为百分数
        df_table['完成率'] = df_table['完成率'].apply(lambda x: '%.2f%%' % (x * 100))
        df_table_piovt_all = []
        if len(piovt_index) > 1:
            for item in piovt_index:
                if item != piovt_index[-1]:
                    df_table_piovt_index = df.pivot_table(index=item,
                                                          columns=piovt_columns,
                                                          values=piovt_values,
                                                          aggfunc=len,
                                                          fill_value=0,
                                                          margins=True
                                                          )
                    # 将index变为列
                    df_table_piovt_index.reset_index(inplace=True)
                    if len(list(df_table_piovt_index)) < 4:
                        df_table_piovt_index[('业务领域', '否')] = 0
                        df_table_piovt_index = df_table_piovt_index[[(item, ''), ('业务领域', '否'), ('业务领域', '是'), ('业务领域', 'All')]]

                    print(df_table_piovt_index)
                    df_table_piovt_index.columns = [item, '否', '是', '合计']
                    # 将All名称按照给定名字更改
                    df_table_piovt_index.loc[len(df_table_piovt_index) - 1, item] = all_name
                    # 计算完成率
                    df_table_piovt_index['完成率'] = df_table_piovt_index['是'] / df_table_piovt_index['合计']
                    # 设置完成率为百分数
                    df_table_piovt_index['完成率'] = df_table_piovt_index['完成率'].apply(lambda x: '%.2f%%' % (x * 100))
                    df_table_piovt_all.append(df_table_piovt_index)
            df_table_piovt = pd.concat(df_table_piovt_all, ignore_index=True)
            df_table = pd.concat([df_table, df_table_piovt], ignore_index=True)
            df_table = df_table[df_table[piovt_index[-1]] != all_name]
            df_table.reset_index(drop=True, inplace=True)
            for i in range(len(df_table)):
                if isinstance(df_table[piovt_index[-1]][i], float):
                    df_table.loc[i, piovt_index[-1]] = df_table[piovt_index[-2]][i]
    else:
        # 建立基础透视表
        df_table = df.pivot_table(index=piovt_index,
                                  columns=piovt_columns,
                                  values=piovt_values,
                                  aggfunc=len,
                                  fill_value=0,
                                  margins=True
                                  )
        # 将index变为列
        df_table.reset_index(inplace=True)
        if len(list(df_table)) < len(table_columns):
            df_table[('业务领域', '否')] = 0
            if len(table_columns) == 4:
                df_table = df_table[[(table_columns[0], ''), ('业务领域', '否'), ('业务领域', '是'), ('业务领域', 'All')]]
            else:
                df_table = df_table[[(table_columns[0], ''), (table_columns[1], ''), ('业务领域', '否'), ('业务领域', '是'), ('业务领域', 'All')]]
        print(df_table)
        # 将All名称按照给定名字更改
        df_table.columns = table_columns
        # 去除All行
        df_table = df_table.head(len(df_table)-1)
        # 计算完成率
        df_table['完成率'] = df_table['是'] / df_table['合计']
        # 设置完成率为百分数
        df_table['完成率'] = df_table['完成率'].apply(lambda x: '%.2f%%' % (x * 100))

    elapsed_time = "{0}模块用时{1}秒".format(sys._getframe().f_code.co_name, time.time() - start_time)
    print(elapsed_time)
    return df_table

def pandas_rat(df):
    start_time = time.time()
    df_table_department = sfb_plan_piovt_table(df,
                                               piovt_index=['责任科室'],
                                               piovt_columns=['是否完成'],
                                               piovt_values=['业务领域'],
                                               table_columns=['责任科室', '否', '是', '合计'],
                                               all_name='安监部')

    df_table_duty_person = sfb_plan_piovt_table(df,
                                                piovt_index=['责任人员'],
                                                piovt_columns=['是否完成'],
                                                piovt_values=['业务领域'],
                                                table_columns=['责任人员', '否', '是', '合计'],
                                                all_name='安监部')

    df_table_all = sfb_plan_piovt_table(df,
                                        piovt_index=['责任科室', '责任人员'],
                                        piovt_columns=['是否完成'],
                                        piovt_values=['业务领域'],
                                        table_columns=['责任科室', '责任人员', '否', '是', '合计'],
                                        all_name='安监部')


    # df_table['科室否'] = df_table.groupby('责任科室')['否'].transform('sum')
    # df_table['科室是'] = df_table.groupby('责任科室')['是'].transform('sum')
    # df_table['科室合计'] = df_table.groupby('责任科室')['合计'].transform('sum')
    # df_table['科室完成率'] = df_table['科室是'] / df_table['科室合计']
    # df_table['科室完成率'] = df_table['科室完成率'].apply(lambda x: '%.2f%%' % (x * 100))

    elapsed_time = "{0}模块用时{1}秒".format(sys._getframe().f_code.co_name, time.time() - start_time)
    print(elapsed_time)
    return df_table_department, df_table_duty_person, df_table_all


def pandas_rat_text(df_list):
    now_day = oas_n_day.NDay(datetime.date.today())
    title_text = ['', '正在执行中', '到期']
    text_list = []
    for i, item in enumerate(df_list):
        all_text = '2021年昆明供电局安全监管部（应急指挥中心）共有{0}条计划{17}，截至{1}，完成{2}条，未完成{3}条，完成率{4};'\
    '其中：安监科共有{5}条计划，完成{6}条，未完成{7}条，完成率{8}。'\
    '应急科共有{9}条计划，完成{10}条，未完成{11}条，完成率{12}。'\
    '党支部共有{13}条计划，完成{14}条，未完成{15}条，完成率{16}。'.format(item.loc[item['责任科室'] == '安监部', '合计'].values[0],
                                                   now_day.get_chinese_date()[3],
                                                   item.loc[item['责任科室'] == '安监部', '是'].values[0],
                                                   item.loc[item['责任科室'] == '安监部', '否'].values[0],
                                                   item.loc[item['责任科室'] == '安监部', '完成率'].values[0],
                                                   item.loc[item['责任科室'] == '安监科', '合计'].values[0],
                                                   item.loc[item['责任科室'] == '安监科', '是'].values[0],
                                                   item.loc[item['责任科室'] == '安监科', '否'].values[0],
                                                   item.loc[item['责任科室'] == '安监科', '完成率'].values[0],
                                                   item.loc[item['责任科室'] == '应急科', '合计'].values[0],
                                                   item.loc[item['责任科室'] == '应急科', '是'].values[0],
                                                   item.loc[item['责任科室'] == '应急科', '否'].values[0],
                                                   item.loc[item['责任科室'] == '应急科', '完成率'].values[0],
                                                   item.loc[item['责任科室'] == '党支部', '合计'].values[0],
                                                   item.loc[item['责任科室'] == '党支部', '是'].values[0],
                                                   item.loc[item['责任科室'] == '党支部', '否'].values[0],
                                                   item.loc[item['责任科室'] == '党支部', '完成率'].values[0],
                                                   title_text[i]
                                                   )
        text_list.append(all_text)
    return text_list


def write_docx(sql):
    now_day = oas_n_day.NDay(datetime.date.today())
    doc = OasDocx()
    doc.hd('昆明供电局2021年安全监管重点工作计划执行情况', font_size=22)
    correct_tex = "昆明供电局2021年安全监管重点工作计共有{0}条数据，以下数据不完善：{1}"\
        .format(analysis_null(sql)[0], analysis_null(sql)[3])
    doc.par('一、数据总体情况', bold=True)
    doc.par(correct_tex)

    doc.par('二、未安排事项', bold=True)
    no_arrange_list = analysis_no_arrange(sql)
    i = 1
    for item in no_arrange_list:
        doc.par('（{0}）未安排{0}'.format(digital_to_chinese(i)))
        x = item.split('\n')
        for y in x:
            doc.par(y)
        i += 1
    doc.par('三、周期事项事项', bold=True)
    s_columns = ['开始时间', '完成时间', '是否完成']
    h_columns = ['关键任务', '行动计划',  '开始时间', '完成时间', '是否完成', '责任科室', '责任人员', '配合人员', '目前推进情况']
    #analysis_list = analysis_section(sql, s_columns, h_columns, '否', '2021-1-1')
    df_all = analysis_all(sql, ['开始时间', '完成时间'], '2021-1-1')
    df_table = []
    for i in range(len(df_all)):
        df_table.append(pandas_rat(df_all[i])[0])
    all_text = pandas_rat_text(df_table)
    doc.par(all_text[0])
    doc.d_table(df_table[0])
    doc.par(all_text[1])
    doc.d_table(df_table[1])
    doc.par(all_text[2])
    doc.d_table(df_table[2])

    doc.par('四、个人正在开展事项', bold=True)
    df_table = []
    for i in range(len(df_all)):
        df_table.append(pandas_rat(df_all[i])[1])
    doc.par('个人正在开展事项', bold=True)
    doc.d_table(df_table[1])
    doc.par('个人超期事项', bold=True)
    doc.d_table(df_table[2])
    dm = DataFrameBothMysql()
    df = dm.select_mysql(sql)
    duty_name = list(set(df['责任人员'].tolist()))
    duty_name.sort()
    s_columns = ['开始时间', '完成时间', '是否完成', '责任人员']
    j = 1
    for name in duty_name:
        analysis_name_list = analysis_name(sql, name, s_columns, h_columns, '否', '2020-1-1')
        doc.par('（{0}）{1}'.format(oas_func.number_to_chinese_number(j), name))
        if analysis_name_list == []:
            doc.par('无异常')
        else:
            i = 1
            for item in analysis_name_list:
                doc.par('{0}.未完成事项{1}'.format(i, oas_func.number_to_chinese_number(i)))
                x = item.split('\n')
                for y in x:
                    doc.par(y)
                i += 1

        j += 1

    doc.save_docx(sfb_plan_save_path('docx'))

def write_pptx(sql):
    now_day = oas_n_day.NDay(datetime.date.today())

    op = OasPptx(r'D:\JGY\600-Data\004-auxiliary辅助文件\南方电网logo（16，9）模板.pptx')
    text_list = ['昆明供电局2021年安全监管重点工作计划执行情况',
                 '安全监管部（应急指挥中心）',
                 now_day.get_chinese_date()[3]]
    op.alter_csg_logo_first_page(text_list)
    correct_tex = "昆明供电局2021年安全监管重点工作计共有{0}条数据，以下数据不完善：{1}" \
        .format(analysis_null(sql)[0], analysis_null(sql)[3])
    op.text_n('一、数据总体情况', font_bold=True)
    op.text_n(correct_tex, textbox_top=12)
    op.page()
    op.text('二、周期事项事项')
    s_columns = ['开始时间', '完成时间', '是否完成']
    h_columns = ['关键任务', '行动计划', '开始时间', '完成时间', '是否完成', '责任科室', '责任人员', '配合人员', '目前推进情况']
    analysis_list = analysis_section(sql, s_columns, h_columns, '否', '2021-1-1')
    df_all = analysis_all(sql, ['开始时间', '完成时间'], '2021-1-1')
    df_table = []
    for i in range(len(df_all)):
        df_table.append(pandas_rat(df_all[i])[0])
    all_text = pandas_rat_text(df_table)
    op.text_n(all_text[0], textbox_top=12)
    df_table_chart = []
    for i in range(len(df_all)):
        df_table_duty_person = sfb_plan_piovt_table(df_all[i],
                                                    piovt_index=['责任人员'],
                                                    piovt_columns=['是否完成'],
                                                    piovt_values=['业务领域'],
                                                    table_columns=['责任人员', '否', '是', '合计']
                                                    )
        df_table_chart.append(df_table_duty_person)
    op.chart(df_table_chart[0][['责任人员', '否', '是', '合计']], chart_class='bar', top=35)
    op.page()
    op.text('全部计划科室完成情况表')
    op.p_table(df_table[0])
    op.page()
    op.text('全部计划个人完成情况表')
    op.p_table(df_table_chart[0])
    op.page()
    op.text('二、周期事项事项')
    op.text_n(all_text[1], textbox_top=12)

    op.chart(df_table_chart[1][['责任人员', '否', '是', '合计']], chart_class='bar', top=35)
    op.page()
    op.text('执行中计划科室完成情况表')
    op.p_table(df_table[1])
    op.page()
    op.text('执行中计划个人完成情况表')
    op.p_table(df_table_chart[1])
    op.page()
    op.text('二、周期事项事项')
    op.text_n(all_text[2], textbox_top=12)
    op.chart(df_table_chart[2][['责任人员', '否', '是', '合计']], chart_class='bar', top=35)
    op.page()
    op.text('到期计划科室完成情况表')
    op.p_table(df_table[2])
    op.page()
    op.text('到期计划个人完成情况表')
    op.p_table(df_table_chart[2])
    op.page()
    duty_name = list(set(df_all[2]['责任人员'].tolist()))
    duty_name.sort()
    op.text('到期个人计划明细')
    op.text_n(duty_name, textbox_top=12)
    op.page()

    s_columns = ['关键任务', '行动计划', '开始时间', '完成时间', '是否完成', '责任科室', '责任人员', '配合人员', '目前推进情况']
    j = 1
    for name in duty_name:
        analysis_name_df = df_all[2][(df_all[2]['责任人员'] == name)&(df_all[2]['是否完成'] == '否')][s_columns]
        pa = PandasAnalysis(analysis_name_df)
        analysis_name_list = pa.df_iter(analysis_name_df, num=False)

        if analysis_name_list == []:
            op.text_n('（' + oas_func.number_to_chinese_number(j) + '）' + name)
            op.text_n('无异常', textbox_top=12)
            op.page()
        else:
            i = 1
            for item in analysis_name_list:
                op.text_n('（{0}）{1}'.format(oas_func.number_to_chinese_number(j), name))
                op.text_n('{0}.未完成事项{1}'.format(i, oas_func.number_to_chinese_number(i)), textbox_top=12)
                x = item.split('\n')
                z = []
                for k, y in enumerate(x):
                    z.append(y)

                op.text_n(z, textbox_top=22)
                i += 1
                op.page()

        j += 1

    dm = DataFrameBothMysql()
    df = dm.select_mysql(sql)
    duty_name = list(set(df['责任人员'].tolist()))
    duty_name.sort()
    s_columns = ['开始时间', '完成时间', '是否完成', '责任人员']
    op.text('正在执行中个人计划明细')
    duty_name_list = ['（'+oas_func.number_to_chinese_number(i+1)+'）'+item for i, item in enumerate(duty_name)]
    op.text_n(duty_name_list, textbox_top=12)
    op.page()
    j = 1
    for name in duty_name:
        analysis_name_list = analysis_name(sql, name, s_columns, h_columns, '否', '2020-1-1')
        #op.text_n('（{0}）{1}'.format(digital_to_chinese(j), name))
        if not analysis_name_list:
            op.text_n('（' + oas_func.number_to_chinese_number(j) + '）' + name)
            op.text_n('无异常', textbox_top=12)
            op.page()
        else:
            i = 1
            for item in analysis_name_list:
                op.text_n('（{0}）{1}'.format(oas_func.number_to_chinese_number(j), name))
                op.text_n('{0}.未完成事项{1}'.format(i, digital_to_chinese(i)),textbox_top=12)
                x = item.split('\n')
                z = []
                for k, y in enumerate(x):
                    z.append(y)

                op.text_n(z, textbox_top=22)
                i += 1
                op.page()

        j += 1
    op.save_pptx(sfb_plan_save_path('pptx'))


# 计划分析模块
def sfb_plan_main():
    # 用Excel文件更新MySQL数据库
    sfb_plan_update()
    # 将MySQL中的表备份到本地Excel
    sfb_plan_backup()
    # 统一查询库
    sql = "select * from 昆明供电局安全监管重点工作计划"
    # Word文档自动生成
    write_docx(sql)
    # PPT文档自动生成
    write_pptx(sql)

if __name__ == '__main__':
    sfb_plan_main()
