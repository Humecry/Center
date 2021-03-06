#! /usr/bin/env python3
# -*- coding:utf-8 -*-

# @Date    : 2018-08-06 17:39:13
# @Author  : Hume (102734075@qq.com)
# @Link    : https://humecry.wordpress.com/
# @Version : 1.1
# @Description: 3.会员消费统计

from common import *

'''
各店会员销售占比
'''
def memberSales(process_bar):
	# 3个月业绩
	conn = pyodbc.connect(S008)
	conn2 = pyodbc.connect(S003)
	salesDf = pd.DataFrame()
	for i in range(-3, 0, 1):
		#全店业绩
		sql = '''
				SELECT
					zsn [店号],
					SUM( zxs * ( 1- ( 1- sfl ) * 0 ) ) [全店业绩{startDate}]
				FROM
					pos.dbo.s_xszb,
					pos.dbo.yblb,
					pos.dbo.jlqtab
				WHERE
					zrq = bly
					AND zsn < 10000
					AND lb = 2
					AND c = zsn % 1000
					AND zrq BETWEEN '{startDate}'	AND '{endDate}'
					AND zsn / 1000 % 10 < 2
					AND zsn % 1000 IN (1, 2, 6, 7, 9, 88, 188)
				GROUP BY
					zsn,
					zsn % 1000,
					zsn / 1000 % 10
				ORDER BY
					zsn % 1000
		'''.format(startDate=lastMonthDate(i), endDate=lastMonthDate(i, 'end'))
		df1 = pd.read_sql(sql, conn)

		process_bar.show_process()

	#     会员销售额
		sql = '''
				--除绿苑高百其他分店
				SELECT
					a.ykd [店号],
					SUM( a.je ) [会员销售额{startDate}] 
				FROM
					pos.dbo.jhjltab AS a
					LEFT JOIN pos.dbo.jhdjtab AS b ON a.k#=b.h#
					LEFT JOIN pos.dbo.jmgtab AS c ON  a.bg%10000=c.gz
					LEFT JOIN pos.dbo.jbrtab AS d ON a.hs = d.bx#
				WHERE
					a.k#>0
					AND a.bg%10000 NOT BETWEEN 6000 AND 6004
					AND a.rt BETWEEN '{startDate}' AND '{nextStartDate}'
				GROUP BY
					a.ykd
				UNION ALL
				--绿苑高百
				SELECT
					1006 [店号],
					SUM( a.je ) [会员销售额{startDate}] 
					FROM
					pos.dbo.jhjltab AS a
					LEFT JOIN pos.dbo.jhdjtab AS b ON a.k#=b.h#
					LEFT JOIN pos.dbo.jmgtab AS c ON  a.bg%10000=c.gz
					LEFT JOIN pos.dbo.jbrtab AS d ON a.hs = d.bx#
				WHERE
					a.k#>0
					AND a.bg%10000 BETWEEN 6000 AND 6004
					AND a.rt BETWEEN '{startDate}' AND '{nextStartDate}'
				GROUP BY
					a.ykd
		'''.format(startDate=lastMonthDate(i), nextStartDate=lastMonthDate(i+1))
		df2 = pd.read_sql(sql, conn2)

		process_bar.show_process()

		df = pd.merge(df1, df2)
		salesDf = pd.merge(salesDf, df) if not salesDf.empty else df
	for i in range(3):
		proportionDf = salesDf.apply(lambda x: x[2*i+2]/x[2*i+1], axis=1)
		salesDf[lastMonthDate(-3+i)[4:6]+'月占比'] = proportionDf
	salesDf.set_index('店号', inplace=True)
	salesDf = salesDf.stack().unstack(0)
	salesDf.reset_index(inplace=True)
	salesDf.drop([0, 1, 2, 3], inplace=True)
	salesDf.columns = ['店别', '海富', '东孚', '绿苑', '绿苑百货', '灌口', '塔埔', '万达', '华森']
	salesDf.loc[4, '店别'] = '全店业绩'
	salesDf.loc[5, '店别'] = '会员销售额'
	salesDf.set_index('店别', inplace=True)
	return salesDf

def excel_output(writer, sales):
	workbook = writer.book
	sales.to_excel(writer, sheet_name='各店会员销售占比')
	worksheet = writer.sheets['各店会员销售占比']
	col = [{'header': '内容'}]
	for value in sales.columns.values:
		col.append({'header': value})
	worksheet.add_table('A1:I6', {'columns': col, 'style': 'Table Style Medium 11'})
	common = {
		'font_name': '微软雅黑',
		'text_wrap': True,
		'align': 'center',
		'valign': 'vcenter',
	}
	cell_format = workbook.add_format(common)
	cell_format_num = workbook.add_format({
		**common,
		'num_format': '00',
	})
	cell_format_percent = workbook.add_format({
		**common,
		'num_format': '0%',
	})
	worksheet.set_row(0, None, cell_format)
	worksheet.set_row(1, None, cell_format_num)
	worksheet.set_row(2, None, cell_format_num)
	worksheet.set_row(3, None, cell_format_percent)
	worksheet.set_row(4, None, cell_format_percent)
	worksheet.set_row(5, None, cell_format_percent)
	worksheet.set_column('A:I', 11)
	
	# 画图
	chart = workbook.add_chart({'type': 'column'})
	font_common = {
		'name': '微软雅黑',
		'size': 14,
	}
	chart.add_series({
		'name': '各店会员销售占比!$A$2',
		'categories': '=各店会员销售占比!$B$1:$I$1',
		'values':     '=各店会员销售占比!$B$2:$I$2',
		# 'data_labels': {'value': True, 'position': 'center', 'font': font_common},
	})
	chart.add_series({
		'name': '各店会员销售占比!$A$3',
		'categories': '=各店会员销售占比!$B$1:$I$1',
		'values':     '=各店会员销售占比!$B$3:$I$3',
		# 'data_labels': {'value': True, 'position': 'center', 'font': font_common},
	})

	line_chart = workbook.add_chart({'type': 'line'})
	line_chart.add_series({
		'name': '各店会员销售占比!$A$4',
		'categories': '=各店会员销售占比!$B$1:$I$1',
		'values':     '=各店会员销售占比!$B$4:$I$4',
		# 'data_labels': {'value': True, 'position': 'center', 'font': {**font_common, 'color': '#10243f'}},
		'y2_axis': True,
	})
	line_chart.add_series({
		'name': '各店会员销售占比!$A$5',
		'categories': '=各店会员销售占比!$B$1:$I$1',
		'values':     '=各店会员销售占比!$B$5:$I$5',
		# 'data_labels': {'value': True, 'font': {**font_common, 'color': '#376092'}},
		'y2_axis': True,
	})
	line_chart.add_series({
		'name': '各店会员销售占比!$A$6',
		'categories': '=各店会员销售占比!$B$1:$I$1',
		'values':     '=各店会员销售占比!$B$6:$I$6',
		# 'data_labels': {'value': True, 'font': {**font_common, 'color': '#984807'}},
		'y2_axis': True,
	})
	# 设置y轴
	chart.set_y_axis({
		'name': '万元',
		'display_units': 'ten_thousands',
		'display_units_visible': False,
		'num_font': font_common,
		'num_format': '0',
	})
	line_chart.set_y2_axis({
		'num_font': font_common,
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
	max_steps = 6
	process_bar = ShowProcess(max_steps, 'OK')
	# 3.会员消费统计
	sales = memberSales(process_bar)
	# 保存为Excel文件名称
	writer = pd.ExcelWriter(PATH + '3.会员消费统计.xlsx', engine='xlsxwriter')
	excel_output(writer, sales)
	writer.save()

if __name__ == '__main__':
	main()