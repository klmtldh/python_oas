import locale
import time  # 用于时间操作
from dateutil.relativedelta import relativedelta  # 用于获取时间间隔
import datetime  # 用于日期操作
import calendar  # 用于日历操作
from dateutil import rrule  # 用于时间间隔
import re


class NDay(object):
    """
    此类用于获取给定日期所在月第一天和最后一天，或任意日的间隔一个月的开始和结束日期；
    实例变量只需要给定任意日期。
    the_n_day方法参数n为需要返回的开始日,m为需要返回的间隔月。
    例如：返回给定日期datetime所在月的第一天和最后一天
    d=nDay（datetime）
    d.the_n
    """

    def __init__(self, date_time: datetime.datetime):
        if isinstance(date_time, str):
            self.date_time = datetime.datetime.strptime(date_time, '%Y-%m-%d').date()
        else:
            self.date_time = date_time

    def get_chinese_date(self):
        dt = self.date_time
        now_year = dt.strftime('%Y年')
        now_month = dt.strftime('%m月')
        now_day = dt.strftime('%d日')
        now_year_month_day = dt.strftime('%Y年%m月%d日')
        now_year_month = dt.strftime('%Y年%m月')
        now_month_day = dt.strftime('%m月%d日')
        now_week = '星期{}'.format(self.digital_to_week(dt.strftime('%w')))
        now_n_week = '{}第{}周'.format(now_year, dt.strftime('%U'))
        now_year_c = '{}年'.format(self.digital_to_chinese(dt.strftime('%Y')))
        now_month_c = '{}月'.format(self.digital_to_chinese(dt.strftime('%m')))
        now_day = dt.strftime('%d日')
        now_year_month_day = dt.strftime('%Y年%m月%d日')
        now_year_month = dt.strftime('%Y年%m月')
        now_month_day = dt.strftime('%m月%d日')
        return now_year, now_month, now_day, now_year_month_day, now_year_month, now_month_day, now_week, now_n_week, now_year_c

    def digital_to_week(self, num):
        num_dict = {'0': '日', '1': '一', '2': '二', '3': '三', '4': '四', '5': '五', '6': '六', '7': '日'}
        if int(num) <= 7:
            num_str = str(num)
            week = num_dict[num_str]
        else:
            week = '输入错误'
        return week

    def digital_to_chinese(self, num):
        num_dict = {'1': '一', '2': '二', '3': '三', '4': '四', '5': '五', '6': '六', '7': '七', '8': '八', '9': '九',
                    '0': '〇', }
        index_dict = {1: '', 2: '', 3: '', 4: '', 5: '', 6: '', 7: '', 8: '', 9: ''}
        nums = list(str(num))
        nums_index = [x for x in range(1, len(nums) + 1)][-1::-1]
        str_ = ''
        for index, item in enumerate(nums):
            str_ = "".join((str_, num_dict[item], index_dict[nums_index[index]]))
        str_ = re.sub("〇[十百千〇]*", "〇", str_)
        str_ = re.sub("〇万", "万", str_)
        str_ = re.sub("亿万", "亿〇", str_)
        str_ = re.sub("〇〇", "〇", str_)
        str_ = re.sub("\\b", "", str_)
        return str_

    # 获取输入日期所在月份的第一天，最后一天和第n天，依据m可以是输入日期的前后n个月
    def get_n_day(self, n=1, m=0):
        # this_month_start = datetime.datetime(self.date_time.year, self.date_time.month, 1)
        this_month_nday = datetime.datetime(self.date_time.year, self.date_time.month, n)  # +datetime.timedelta(days=n)
        this_month_end = datetime.datetime(self.date_time.year, self.date_time.month,
                                           calendar.monthrange(self.date_time.year, self.date_time.month)[1])
        # n_month_start=this_month_start +relativedelta(months=m)
        n_month_end = this_month_end + relativedelta(months=m)
        n_month_nday = this_month_nday + relativedelta(months=m)

        return n_month_nday, n_month_end

    # 获取输入日期所在周的第一天，最后一天和第n天，依据m可以是输入日期的前后n个周
    def get_current_week(self, date_time, n):
        monday, sunday = date_time, date_time
        one_day = datetime.timedelta(days=1)
        while monday.weekday() != 0:
            monday -= one_day
        while sunday.weekday() != 6:
            sunday += one_day
        # 返回当前的星期一和星期天的日期
        week_n = monday + datetime.timedelta(days=n)

        return monday, sunday, week_n


if __name__ == '__main__':
    nd = NDay('1970-5-10')
    print(nd.get_chinese_date())
