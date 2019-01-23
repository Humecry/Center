#! /usr/bin/env python3
# -*- coding:utf-8 -*-

# @Date    : 2018-08-07 09:03:59
# @Author  : Hume (102734075@qq.com)
# @Link    : https://humecry.wordpress.com/
# @Version : 1.0
# @Description: 6.负毛利统计

from common import *

'''
负毛利统计
'''
def negativeProfit(process_bar):
	conn = pyodbc.connect(S008)
	negativeProfitDf = pd.DataFrame()
	for i in range(-6, 0, 1):
		# 负毛利
		sql = '''
				SELECT
					[门店编号],
					[门店] AS [店别],
					SUM ( [毛利] ) AS [毛利]
				FROM
					(
					SELECT
						e.c AS [门店编号],
						e.lqm AS [门店],
						SUM ( a.qc - a.qi ) AS [毛利] 
					FROM
						pos.dbo.bzsrbak AS a
						LEFT JOIN pos.dbo.jlqtab AS e ON a.sn = e.c --分店
					WHERE
						e.lb= 2 
						AND a.rq >= '{startDate}' 
						AND a.rq < '{nextStartDate}' 
						AND a.g IN ( 1101, 1106, 1205, 1206, 1207 ) 
					GROUP BY
						a.s,
						e.lqm,
						e.c
					HAVING
						SUM ( a.qc - a.qi ) < 0
					) AS d
				GROUP BY
					[门店编号],
					[门店]
				ORDER BY
					[门店编号]
		'''.format(startDate=lastMonthDate(i), nextStartDate=lastMonthDate(i+1))
		df1 = pd.read_sql(sql, conn)

		process_bar.show_process()

		# 业绩
		sql = '''
				SELECT
					lqm [店别],
					SUM( zxs * ( 1- ( 1- sfl ) * 0 ) ) [业绩]
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
					-- AND zsn % 1000 NOT IN ( 3, 8 )
					AND zsn % 1000 IN (1, 2, 6, 7, 9, 88, 188)
				GROUP BY
					lqm,
					zsn % 1000
				ORDER BY
					zsn % 1000
		'''.format(startDate=lastMonthDate(i), endDate=lastMonthDate(i, 'end'))
		df2 = pd.read_sql(sql, conn)

		process_bar.show_process()

		df = pd.merge(df1, df2)
		df.loc[7] = df.apply(lambda x: x.sum())
		df['店别'].iat[7] = '合计'
		proportion = df.apply(lambda x: abs(x[2]/x[3]), axis=1)
		df['占业绩比'] = proportion
		df.columns = [['', '', lastMonthDate(i), lastMonthDate(i), lastMonthDate(i)], ['门店编号', '店别', '负毛利', '业绩', '占业绩比']]
		negativeProfitDf = pd.merge(negativeProfitDf, df) if not negativeProfitDf.empty else df
	negativeProfitDf.set_index(('', '店别'), inplace=True)
	negativeProfitDf.drop(['业绩'], inplace=True, axis=1, level=1)
	negativeProfitDf.drop(['门店编号'], inplace=True, axis=1, level=1)
	return negativeProfitDf

def excel_output(writer, profits):
	workbook = writer.book
	profits.to_excel(writer, sheet_name='各店负毛利', header=False)
	worksheet = writer.sheets['各店负毛利']
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
	cell_format_percent = workbook.add_format({**common, 'num_format': '0.00%', 'bg_color': '#d8e4bc'})
	cell_format_percent_even = workbook.add_format({**common, 'num_format': '0.00%', 'bg_color': '#ebf1de'})
	cell_format_red = workbook.add_format({
		**common,
		'font_color': 'red',
		'bold': True,
	})
	# 设置标题
	worksheet.merge_range(0, 0, 1, 0, '店别', cell_format)
	for i in range(6):
		worksheet.merge_range(0, 2*i+1, 0, 2*i+2, lastMonthDate(-6+i)[4:6]+'月', cell_format)
		worksheet.write(1, 2*i+1, '负毛利额', cell_format)
		worksheet.write(1, 2*i+2, '占比业绩', cell_format)
	# 写入数据
	for row_num, value in enumerate(profits.index.values):
		worksheet.write(row_num+2, 0, value, cell_format)
	for col_num, col in enumerate(profits.columns.values):
		for row_num, value in enumerate(profits[col]):
			if col_num % 2 == 0:
				if (row_num + 1) % 2:
					worksheet.write(row_num+2, col_num+1, value, cell_format_num)
				else:
					worksheet.write(row_num+2, col_num+1, value, cell_format_num_even)
			else:
				if (row_num + 1) % 2:
					worksheet.write(row_num+2, col_num+1, value, cell_format_percent)
				else:
					worksheet.write(row_num+2, col_num+1, value, cell_format_percent_even)
	# 设置列宽
	worksheet.set_column('A:A', 16)
	worksheet.set_column('B:M', 11.3)
	# 设置行高
	for i in range(10):
		worksheet.set_row(i, 44.25)
	# 设置条件格式
	for i in range(6):
		worksheet.conditional_format(
			2, 2*i+2, 8, 2*i+2,
			{
				'type': 'top',
				'value': 3,
				'format': cell_format_red
			}
		)
def main():
	# 进度条
	max_steps = 12
	process_bar = ShowProcess(max_steps, 'OK')
	# 6.负毛利统计
	profits = negativeProfit(process_bar)
	# 保存为Excel文件名称
	writer = pd.ExcelWriter(PATH + '6.负毛利统计.xlsx', engine='xlsxwriter')
	excel_output(writer, profits)
	writer.save()

if __name__ == '__main__':
	main()