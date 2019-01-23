#! /usr/bin/env python3
# -*- coding:utf-8 -*-

# @Date    : 2018-08-03 16:39:09
# @Author  : Hume (102734075@qq.com)
# @Link    : https://humecry.wordpress.com/
# @Version : 1.1
# @Description: 汇总月度报告所有数据, 导出时间在6分钟内. 如遇错误, 可执行单个文件

# 引入公共函数
from common import *
# 引入自建模块
import _1_salesAnalysis
import _2_shoppingCardSales
import _3_memberSales
import _4_passengerFlow
import _5_unusual
import _6_negativeProfit
import _7_order
import _8_arrivalRate

# 进度条
max_steps = 324
process_bar = ShowProcess(max_steps, '恭喜, 已完成!')

writer = pd.ExcelWriter('月度报告数据源.xlsx', engine='xlsxwriter')

# 1.消费数据分析
shopping = _1_salesAnalysis.shoppingCardMember(process_bar)
_1_salesAnalysis.excel_output(writer, shopping)

# 2.购物卡消费统计: 购物卡消费趋势, 购物卡消费店别排名
thisYearDf = _2_shoppingCardSales.shoppingCard(STORES, -6, 0, process_bar)
lastYearDf = _2_shoppingCardSales.shoppingCard(STORES, -18, -12, process_bar)
	# 购物卡消费趋势
trend = _2_shoppingCardSales.shoppingTrend(thisYearDf, lastYearDf)
	# 购物卡消费店别排名
rank = _2_shoppingCardSales.shoppingCardRank(thisYearDf)
_2_shoppingCardSales.excel_output(writer, trend, rank)

# 3.会员消费统计
sales = _3_memberSales.memberSales(process_bar)
_3_memberSales.excel_output(writer, sales)

# 4.会员客流统计
flows = _4_passengerFlow.passengerFlow(process_bar)
_4_passengerFlow.excel_output(writer, flows)

# 5.异常数据统计
abnormal = _5_unusual.unusual(process_bar)	
_5_unusual.excel_output(writer, abnormal)

# 6.负毛利统计
profits = _6_negativeProfit.negativeProfit(process_bar)
_6_negativeProfit.excel_output(writer, profits)

# 7.订单统计
orders = _7_order.order(process_bar)
_7_order.excel_output(writer, orders)

# 8.供货商到货率统计
arrival = _8_arrivalRate.arrivalRate(STORES, -12, process_bar)
_8_arrivalRate.excel_output(writer, arrival)

# 保存Excel
writer.save()

time.sleep(30)