import os
import re
import time
import configparser
import pandas as pd


# 用于获取电脑名称
def get_pc_name():
    return os.environ['COMPUTERNAME']


# 用于获取配置文件路径，区分办公室和个人电脑
def ini_path():
    if get_pc_name() == 'LAPTOP-HF9P6H1P':
        in_pt = r'D:\JGY\600-Data\001-ini配置文件\孔令配置.ini'
    else:
        in_pt = r'Z:\600数据库\000配置文件\昆明局安监部配置文件.ini'
    return in_pt

# 获取模块路径，电脑路径，Sheet，MySQL数据库及表格
def read_file_map(module_name, file_name, in_or_out):
    # 模块用时计时
    start_time = time.time()
    # 读取配置文件ini
    config = configparser.ConfigParser(interpolation=configparser.ExtendedInterpolation())
    config.read(ini_path(), encoding='utf-8')
    path = config.get('文件路径', '工作输入文件地图目录')
    name = os.path.join(path, '昆明供电局安监部文件地图.xlsx')
    ef = pd.ExcelFile(name, engine='openpyxl')
    df = ef.parse('记录')
    if get_pc_name() == 'LAPTOP-HF9P6H1P':
        excel_path = ''.join(
            df[
                (df['模块'] == module_name) &
                (df['文件名称'] == file_name) &
                (df['方式'] == in_or_out)
            ]['个人电脑路径'].values)
    else:
        excel_path = ''.join(
            df[
                (df['模块'] == module_name) &
                (df['文件名称'] == file_name) &
                (df['方式'] == in_or_out)
                ]['服务器路径'].values)

    excel_sheet = ''.join(
        df[
            (df['模块'] == module_name) &
            (df['文件名称'] == file_name) &
            (df['方式'] == in_or_out)
            ]['Sheet名称'].values)

    mysql_db = ''.join(
        df[
            (df['模块'] == module_name) &
            (df['文件名称'] == file_name) &
            (df['方式'] == in_or_out)
            ]['MySQL数据库名称'].values)
    mysql_table = ''.join(
        df[
            (df['模块'] == module_name) &
            (df['文件名称'] == file_name) &
            (df['方式'] == in_or_out)
            ]['MySQL表名称'].values)
    # 模块用时
    elapsed_time = "read_file_map模块用时{0}秒".format(time.time()-start_time)
    print(elapsed_time)
    return excel_path, excel_sheet, mysql_db, mysql_table


def elapsed_time(func):
    """
    作者：孔令
    功能：计算函数消耗时间；为装饰函数。
    参数：func 类型：函数。
    返回：打印被装饰函数运行时间的函数。
    """
    def elapsed(*args, **kwargs):
        start_time = time.time()
        ret = func(*args, **kwargs)
        elapsed_time = time.time() - start_time
        print(f"{func.__name__} time elapsed: {elapsed_time:.5f}")
        return ret
    return elapsed


def number_to_chinese_number(num, flag='complex', financial='no'):
    """
    功能：阿拉伯数字转中文数字;
    参数：num Type int or str为输入阿拉伯数字;
          flag 为输出方式：'complex'为输出带单位中文数字，不能超过10亿；
                           'simple'为不带单位中文数字，可以任意大；
                           'week' 为星期中文数字，范围0-7；0和7认为是周日。
          financial str 是否开启财务计算，num可以为小数。
    返回：chinese_number，Type str。
    """
    chinese_number = ''
    if financial == 'no':
        num_dict = {'1': '一', '2': '二', '3': '三', '4': '四', '5': '五', '6': '六', '7': '七', '8': '八', '9': '九',
                    '0': '〇', }
        nums = list(str(num))

        if flag == 'complex':
            index_dict = {1: '', 2: '十', 3: '百', 4: '千', 5: '万', 6: '十', 7: '百', 8: '千', 9: '亿', 10: '十',
                          11: '百', 12: '千', 13: '万', 14: '亿', 15: '十', 16: '百', 17: '千', 18: '万', 19: '亿'}
            try:
                nums_index = [x for x in range(1, len(nums) + 1)][-1::-1]

                for index, item in enumerate(nums):
                    chinese_number = "".join((chinese_number, num_dict[item], index_dict[nums_index[index]]))
                chinese_number = re.sub("〇[十百千〇]*", "〇", chinese_number)
                chinese_number = re.sub("〇亿", "亿", chinese_number)
                chinese_number = re.sub("〇万", "万", chinese_number)
                chinese_number = re.sub("亿万", "亿〇", chinese_number)
                chinese_number = re.sub("〇〇", "〇", chinese_number)
                if chinese_number != '〇':
                    chinese_number = re.sub("〇\\b", "", chinese_number)
            except Exception as e:
                print(str(e))
                return '输入不能大于亿亿亿，期待下次改进。'

        elif flag == 'simple':
            chinese_number_list = list(map(lambda x: num_dict[x], nums))
            chinese_number = ''.join(chinese_number_list)
        elif flag == 'week':
            num_dict = {'0': '日', '1': '一', '2': '二', '3': '三', '4': '四', '5': '五', '6': '六', '7': '日'}
            if int(num) <= 7:
                num_str = str(num)
                chinese_number = num_dict[num_str]
            else:
                chinese_number = None
    elif financial == 'yes':
        number_arab = '0123456789'
        number_str = '零壹贰叁肆伍陆柒捌玖'
        # 数值字典numDic,和阿拉伯数字是简单的一一对应关系
        number_dict = dict(zip(number_arab, number_str))

        unit_arab = (-2, -1, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
        unit_str = '角分元拾佰仟万亿拾佰仟万亿拾佰仟万亿'
        unit_dict = dict(zip(unit_arab, unit_str))
        number_str = str(num)
        integer_part_number = ''
        decimal_part_number = ''
        try:
            if '.' in number_str:
                integer_part, decimal_part = number_str.split('.')
            else:
                integer_part = number_str
                decimal_part = '0'
            decimal_part = decimal_part[0:2]
            integer_part = list(integer_part)

            decimal_part = list(decimal_part)

            nums_index = [x for x in range(1, len(integer_part) + 1)][-1::-1]
            for index, item in enumerate(integer_part):
                integer_part_number = "".join((integer_part_number, number_dict[item], unit_dict[nums_index[index]]))
            integer_part_number = re.sub("零[拾佰仟零]*", "零", integer_part_number)
            integer_part_number = re.sub("零亿", "亿", integer_part_number)
            integer_part_number = re.sub("零万", "万", integer_part_number)
            integer_part_number = re.sub("亿万", "亿零", integer_part_number)
            integer_part_number = re.sub("零零", "零", integer_part_number)
            integer_part_number = re.sub("零元", "元", integer_part_number)
            if integer_part_number != '零':
                integer_part_number = re.sub("零\\b", "", integer_part_number)

            if len(decimal_part) == 1:
                decimal_part_number = number_dict[decimal_part[0]] + '角'
                decimal_part_number = re.sub("零角", "整", decimal_part_number)
            elif len(decimal_part) == 2:

                nums_index = [x for x in range(1, len(decimal_part) + 1)][-1::-1]
                for index, item in enumerate(decimal_part):
                    decimal_part_number = "".join(
                        (decimal_part_number, number_dict[item], unit_dict[-nums_index[index]]))
                decimal_part_number = re.sub("零角", "零", decimal_part_number)
                decimal_part_number = re.sub("零分", "整", decimal_part_number)
            chinese_number = integer_part_number + decimal_part_number
        except Exception as e:
            print(str(e))
            return '输入不能大于亿亿亿，期待下次改进。'
    return chinese_number


# 实现删除list中指定list
def list_del_list(o_list, d_list):
    r_list = []
    for i in o_list:
        if i not in d_list:
            r_list.append(i)
    return r_list


def split_date_to_year_month_day(df, date_time_columns):
    for item in date_time_columns:
        df[item] = pd.to_datetime(df[item])
        df[item+'年'] = df[item].dt.year
        df[item+'月'] = df[item].dt.month
        df[item + '日'] = df[item].dt.date
    return df


if __name__ == '__main__':

    print(number_to_chinese_number.__doc__)
    print(elapsed_time.__doc__)
    w = number_to_chinese_number(3332)

    print(w)
