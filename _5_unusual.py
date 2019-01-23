#! /usr/bin/env python3
# -*- coding:utf-8 -*-

# @Date    : 2018-08-06 17:48:27
# @Author  : Hume (102734075@qq.com)
# @Link    : https://humecry.wordpress.com/
# @Version : 1.0
# @Description: 5.异常数据统计

from common import *
from conf import *

'''
异常数据统计
'''
def unusual(process_bar):
	# 连接000号服务器的数据库
	conn = pyodbc.connect(S000)
	unusualCategories = ['未销售', '高库存', '即将删除', '畅销缺货']
	unusualDf = []
	for index, item in enumerate(unusualCategories):
		df = []
		for i in range(-2, 1, 1):
			sql = '''
					SELECT
						b.lqm AS [分店],
						a.c AS [{statisticDate}]
					FROM
						pos.dbo._sjfc AS a
						LEFT JOIN pos.dbo.jlqtab AS b ON ( a.sn = b.c )
					WHERE
						a.t = {index} --分类
						AND a.d = '{statisticDate}' --统计日期截止日
						AND b.lb = 2 --店别分类
			'''.format(statisticDate=lastMonthDate(i), index=index)
			df.append(pd.read_sql(sql, conn, '分店'))

			process_bar.show_process()

		df1 = df.pop(0)
		df1.reset_index(inplace=True)
		# 多表拼接
		for value in df:
			df1 = df1.join(value, on='分店')
		# 去掉新华园店, 塔埔小店
		df1.drop(index=5, inplace=True)
		df1.drop(index=8, inplace=True)
		# 计算环比
		a = df1.apply(lambda x:
					  (x[lastMonthDate(0)] - x[lastMonthDate(-1)])/x[lastMonthDate(-1)],
					  axis=1)
		df1.insert(4, '环比', a)
		df1.columns = [
			['', item, item, item, item],
			['分店', lastMonthDate(-3, 'end'), lastMonthDate(-2, 'end'), lastMonthDate(-1, 'end'), '环比']
		]
		unusualDf.append(df1)
	unusualDf1 = unusualDf.pop(0)
	# 多表拼接
	for value in unusualDf:
		unusualDf1 = pd.merge(unusualDf1, value)
	return unusualDf1

def excel_output(writer, abnormal):
	workbook = writer.book
	abnormal.to_excel(writer, sheet_name='各店异常数据报告', header=False)
	worksheet = writer.sheets['各店异常数据报告']
	common = {
		'font_name': '微软雅黑',
		'font_size': 16,
		'text_wrap': True,
		'align': 'center',
		'valign': 'vcenter',
		'border': 1,
		'border_color': '#FFFFFF',
	}
	cell_format = workbook.add_format({**common, 'font_color': '#FFFFFF', 'bold': True, 'bg_color': '#9bbb59'})
	cell_format_num = workbook.add_format({**common, 'num_format': '00', 'bg_color': '#d8e4bc'})
	cell_format_num_even = workbook.add_format({**common, 'num_format': '00', 'bg_color': '#ebf1de'})
	cell_format_percent = workbook.add_format({**common, 'num_format': '0%', 'bg_color': '#d8e4bc'})
	cell_format_percent_even = workbook.add_format({**common, 'num_format': '0%', 'bg_color': '#ebf1de'})
	cell_format_green = workbook.add_format({
		**common,
		'font_color': 'green',
		'bold': True,
	})
	cell_format_red = workbook.add_format({
		**common,
		'font_color': 'red',
		'bold': True,
	})
	# 设置标题
	worksheet.merge_range(0, 0, 1, 0, '店别', cell_format)
	worksheet.merge_range(0, 1, 0, 4, '未销售', cell_format)
	worksheet.merge_range(0, 5, 0, 8, '高库存', cell_format)
	worksheet.merge_range(0, 9, 0, 12, '畅销缺货', cell_format)
	worksheet.merge_range(0, 13, 0, 16, '即将删除', cell_format)
	for i in range(4):
		for j in range(3):
			worksheet.write(1, 4*i+j+1, lastMonthDate(-3+j)[4:6]+'月', cell_format)
		worksheet.write(1, 4*i+4, '环比', cell_format)
	# 写入数据
	for col_num, col in enumerate(abnormal.columns.values):
		for row_num, value in enumerate(abnormal[col]):
			if col_num == 0:
				worksheet.write(row_num+2, col_num, value, cell_format)
			elif col_num in (4, 8, 12, 16):
				if (row_num + 1) % 2:
					worksheet.write(row_num+2, col_num, value, cell_format_percent)
				else:
					worksheet.write(row_num+2, col_num, value, cell_format_percent_even)
			else:
				if (row_num + 1) % 2:
					worksheet.write(row_num+2, col_num, value, cell_format_num)
				else:
					worksheet.write(row_num+2, col_num, value, cell_format_num_even)
	# 清除多余数据
	for row_num in range(8):
		worksheet.write_blank(row_num, 17, '', workbook.add_format())
	# 设置列宽
	worksheet.set_column('A:A', 16)
	worksheet.set_column('B:Q', 8.3)
	# 设置行高
	for i in range(9):
		worksheet.set_row(i, 50)
	# 设置条件格式
	for i in range(4):
		worksheet.conditional_format(
			2, 4*i+4, 8, 4*i+4,
			{
				'type': 'cell',
				'criteria': '<',
				'value': 0,
				'format': cell_format_green
			}
		)
		worksheet.conditional_format(
			2, 4*i+4, 8, 4*i+4,
			{
				'type': 'cell',
				'criteria': '>',
				'value': 0,
				'format': cell_format_red
			}
	)

def main():
	# 进度条
	max_steps = 12
	process_bar = ShowProcess(max_steps, 'OK')
	# 5.异常数据统计
	abnormal = unusual(process_bar)	
	# 保存为Excel文件名称
	writer = pd.ExcelWriter(PATH + '5.异常数据统计.xlsx', engine='xlsxwriter')
	excel_output(writer, abnormal)
	writer.save()

if __name__ == '__main__':
	main()