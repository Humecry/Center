#! /usr/bin/env python3
# -*- coding:utf-8 -*-

# @Date    : 2018-08-07 09:08:39
# @Author  : Hume (102734075@qq.com)
# @Link    : https://humecry.wordpress.com/
# @Version : 1.0
# @Description: 7.订单数据分析RPS占比

from common import *

'''
订单统计
'''
def order(process_bar):
	conn = pyodbc.connect(S008)
	orderDf = pd.DataFrame()
	for i in range(-3, 0 , 1):
		sql = '''
				SELECT
					lqm [店别],
					SUM ( CASE WHEN dfs % 10 = 9 THEN 1 ELSE 0 END ) [RPS下单{startDate}],
					SUM ( CASE WHEN dfs % 10 = 2 THEN 1 ELSE 0 END ) [档期下单{startDate}],
					SUM ( CASE WHEN dfs % 10 = 0 THEN 1 ELSE 0 END ) [日常下单{startDate}]
				FROM
					pos.dbo.jdhtab k,
					pos.dbo.jbrtab,
					pos.dbo.jlqtab a 
				WHERE
					dg# = bx# 
					AND zdr > 0 
					AND dg# NOT IN ( 1077, 1078 ) 
					AND a.lb= 2 
					AND c = dsn 
					AND dht BETWEEN '{startDate}' AND '{nextStartDate}' 
					AND dsn NOT IN ( 3, 4, 8 ) 
					AND dg# >= 0 
					AND dg# < 10000 
					AND zdr > 0 
				GROUP BY
					lqm,
					f0 
				ORDER BY
					f0
		'''.format(startDate=lastMonthDate(i), nextStartDate=lastMonthDate(i+1))
		df = pd.read_sql(sql, conn)

		process_bar.show_process()

		orderDf = pd.merge(orderDf, df) if not orderDf.empty else df
	orderDf.set_index('店别', inplace=True)
	for i in range(3):
		proportionDf = orderDf.apply(lambda x: x[3*i]/(x[3*i]+x[3*i+1]+x[3*i+2]), axis=1)
		orderDf[lastMonthDate(-3+i)[4:6]+'月占比'] = proportionDf
	proportionDf = orderDf.apply(lambda x: x[7]/(x[6]+x[7]+x[8]), axis=1)
	orderDf[lastMonthDate(-1)[4:6]+'月档期占比'] = proportionDf
	orderDf = orderDf.stack().unstack(0)[6:]
	return orderDf

def excel_output(writer, orders):
	workbook = writer.book
	orders.to_excel(writer, sheet_name='订单数据分析RPS占比')
	worksheet = writer.sheets['订单数据分析RPS占比']
	col = [{'header': '内容'}]
	for value in orders.columns.values:
		col.append({'header': value})
	worksheet.add_table('A1:' + chr(65+orders.columns.values.size) + '8', {'columns': col, 'style': 'Table Style Medium 11'})
	common = {
		'font_name': '微软雅黑',
		'text_wrap': True,
		'align': 'center',
		'valign': 'vcenter'
	}
	cell_format = workbook.add_format(common)
	cell_format_num = workbook.add_format({**common, 'num_format': '00'})
	cell_format_percent = workbook.add_format({**common, 'num_format': '0.0%'})
	index = [
		'RPS下单',
		'档期下单',
		'日常下单',
		lastMonthDate(-3)[4:6]+'月占比',
		lastMonthDate(-2)[4:6]+'月占比',
		lastMonthDate(-1)[4:6]+'月占比',
		lastMonthDate(-1)[4:6]+'月档期占比'
	]
	# 写入第一列数据
	for row_num in range(7):
		worksheet.write(row_num+1, 0, index[row_num], cell_format)
	# 设置数据格式
	for i in range(3):
		worksheet.set_row(i+1, None, cell_format_num)
	for i in range(4):
		worksheet.set_row(i+4, None, cell_format_percent)
	# 设置列宽
	worksheet.set_column('A:A', 12)
	
	# 画图
	chart = workbook.add_chart({'type': 'column'})
	font_common = {
		'name': '微软雅黑',
		'size': 14,
	}
	chart.add_series({
		'name': '订单数据分析RPS占比!$A$2',
		'categories': '=订单数据分析RPS占比!$B$1:$H$1',
		'values':     '=订单数据分析RPS占比!$B$2:$H$2',
		# 'data_labels': {'value': True, 'position': 'center', 'font': font_common},
	})
	chart.add_series({
		'name': '订单数据分析RPS占比!$A$3',
		'categories': '=订单数据分析RPS占比!$B$1:$H$1',
		'values':     '=订单数据分析RPS占比!$B$3:$H$3',
		# 'data_labels': {'value': True, 'position': 'center', 'font': font_common},
	})
	chart.add_series({
		'name': '订单数据分析RPS占比!$A$4',
		'categories': '=订单数据分析RPS占比!$B$1:$H$1',
		'values':     '=订单数据分析RPS占比!$B$4:$H$4',
		# 'data_labels': {'value': True, 'position': 'center', 'font': font_common},
	})
	line_chart = workbook.add_chart({'type': 'line'})
	line_chart.add_series({
		'name': '订单数据分析RPS占比!$A$5',
		'categories': '=订单数据分析RPS占比!$B$1:$H$1',
		'values':     '=订单数据分析RPS占比!$B$5:$H$5',
		# 'data_labels': {'value': True, 'position': 'center', 'font': {**font_common, 'color': '#10243f'}},
		'y2_axis': True,
	})
	line_chart.add_series({
		'name': '订单数据分析RPS占比!$A$6',
		'categories': '=订单数据分析RPS占比!$B$1:$H$1',
		'values':     '=订单数据分析RPS占比!$B$6:$H$6',
		# 'data_labels': {'value': True, 'font': {**font_common, 'color': '#376092'}},
		'y2_axis': True,
	})
	line_chart.add_series({
		'name': '订单数据分析RPS占比!$A$7',
		'categories': '=订单数据分析RPS占比!$B$1:$H$1',
		'values':     '=订单数据分析RPS占比!$B$7:$H$7',
		# 'data_labels': {'value': True, 'font': {**font_common, 'color': '#984807'}},
		'y2_axis': True,
	})
	line_chart.add_series({
		'name': '订单数据分析RPS占比!$A$8',
		'categories': '=订单数据分析RPS占比!$B$1:$H$1',
		'values':     '=订单数据分析RPS占比!$B$8:$H$8',
		# 'data_labels': {'value': True, 'font': {**font_common, 'color': '#255399'}},
		'y2_axis': True,
	})
	# 设置y轴
	chart.set_y_axis({
		'name': '单',
		'display_units_visible': False,
		'num_font': font_common,
		'num_format': '0',
	})
	line_chart.set_y2_axis({
		'num_font': font_common,
		'num_format': '0%',
	})
	chart.set_table({
		'font': {
			**font_common,
			'bold': True,
		},
		'show_keys': True
	})
	chart.combine(line_chart)
	chart.set_legend({'font': {**font_common, 'bold': True}, 'position': 'top'})
	chart.set_style(34)
	chart.set_size({'width': 1125, 'height': 660})
	worksheet.insert_chart('A10', chart)

def main():
	# 进度条
	max_steps = 3
	process_bar = ShowProcess(max_steps, 'OK')
	# 7.订单统计
	orders = order(process_bar)
	# 保存为Excel文件名称
	writer = pd.ExcelWriter(PATH + '7.订单统计.xlsx', engine='xlsxwriter')
	excel_output(writer, orders)
	writer.save()

if __name__ == '__main__':
	main()