# -*- coding: utf-8 -*-
"""
Created on Fri Feb 12 20:00:17 2021

@author: Kong Ling

E-Mail:klmtldh@163.com

QQ:1170233

Wechat:klmtldh
"""


import km_sfb_person
import km_sfb_plan
import km_sfb_promptly_problem
import km_sfb_spa_meeting
import time


if __name__ == '__main__':
    # 模块用时计时
    start_time = time.time()
    # 昆明供电局安监部分析模块
    #km_sfb_person.sfb_person_main()
    # 昆明供电局安全监管重点工作计划分析模块
    #km_sfb_plan.sfb_plan_main()
    # 昆明供电局安全监管重点工作计划分析模块
    #km_sfb_promptly_problem.sfb_promptly_problem_main()
    # 昆明供电局安全生产分析会材料
    km_sfb_spa_meeting.sfb_meeting_main()
    # 模块用时
    elapsed_time = "程序用时{0}秒".format(time.time() - start_time)
    print(elapsed_time)
