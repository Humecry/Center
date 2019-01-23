#! /usr/bin/env python3
# -*- coding:utf-8 -*-

# @Date    : 2018-08-07 09:10:55
# @Author  : Hume (102734075@qq.com)
# @Link    : https://humecry.wordpress.com/
# @Version : 1.0
# @Description: 8.供应商到货率统计

from common import *

'''
供应商到货率统计
'''
def arrivalRate(stores, num, process_bar):
	arrivalRateDfs = []
	for value in stores:
		conn = pyodbc.connect(value[0])
		arrivalRateDf = pd.DataFrame(data={'店别': [value[1]]})
	#     每月
		for j in range(num, 0, 1):
			sql = '''
						SELECT
							SUM( ysl ), --已收量
							SUM( dsl ) -- 订量
						FROM
							pos.dbo.jdhtab,
							pos.dbo.jbrtab,
							pos.dbo.jdstab,
							pos.dbo.jsptab
						WHERE
							zdr > 0
							AND dh# = dd#
							AND f0 = f1 % 10000
							AND bx# = dg#
							AND art < getdate ( ) AND dsl > 0
							AND dsp = sx#
							AND jlr / 1000000000 = 0
							AND dht BETWEEN '{startDate}' AND '{nextStartDate}'
			 '''.format(startDate=lastMonthDate(j), nextStartDate=lastMonthDate(j+1))
			df = pd.read_sql(sql, conn)

			process_bar.show_process()

			df.columns = [[lastMonthDate(j)+'已收量', lastMonthDate(j)+'订量']]
			arrivalRateDf = pd.merge(arrivalRateDf, df, left_index=True, right_index=True)
	#     总计
		sql = '''
				SELECT
					SUM( ysl ), --已收量
					SUM( dsl ) -- 订量
				FROM
					pos.dbo.jdhtab,
					pos.dbo.jbrtab,
					pos.dbo.jdstab,
					pos.dbo.jsptab
				WHERE
					zdr > 0
					AND dh# = dd#
					AND f0 = f1 % 10000
					AND bx# = dg#
					AND art < getdate ( ) AND dsl > 0
					AND dsp = sx#
					AND jlr / 1000000000 = 0
					AND dht BETWEEN '{startDate}' AND '{nextStartDate}'
		'''.format(startDate=lastMonthDate(num), nextStartDate=lastMonthDate(0))
		df = pd.read_sql(sql, conn)

		process_bar.show_process()

		df.columns = ['已收量总计', '订量总计']
		arrivalRateDf = pd.merge(arrivalRateDf, df, left_index=True, right_index=True)
		arrivalRateDfs.append(arrivalRateDf)
	arrivalRateDf = pd.concat(arrivalRateDfs)
	arrivalRateDf.reset_index(drop=True, inplace=True)
	arrivalRateDf.loc[7] = arrivalRateDf.apply(lambda x: x.sum())
	arrivalRateDf['店别'].iat[7] = '全公司'
	arrivalRateDf2 = pd.DataFrame(arrivalRateDf['店别'])
	# 求到货率
	for i in range(num, 0, 1):
		df = arrivalRateDf.apply(lambda x: x[2*(i-num)+1]/x[2*(i-num)+2], axis=1)
		arrivalRateDf2[lastMonthDate(i)] = df 
	df = arrivalRateDf.apply(lambda x: x[-2*num+1]/x[-2*num+2], axis=1)
	arrivalRateDf2['总计'] = df
	# 去年同期
	dfs = []
	for value in stores:
		conn = pyodbc.connect(value[0])
		df2 = pd.DataFrame(data={'店别': [value[1]]})
		sql = '''
				SELECT
					SUM( ysl )/SUM( dsl )
				FROM
					pos.dbo.jdhtab,
					pos.dbo.jbrtab,
					pos.dbo.jdstab,
					pos.dbo.jsptab
				WHERE
					zdr > 0
					AND dh# = dd#
					AND f0 = f1 % 10000
					AND bx# = dg#
					AND art < getdate ( ) AND dsl > 0
					AND dsp = sx#
					AND jlr / 1000000000 = 0
					AND dht BETWEEN '{startDate}' AND '{nextStartDate}'
		'''.format(startDate=lastMonthDate(num-12), nextStartDate=lastMonthDate(-12))
		df = pd.read_sql(sql, conn)

		process_bar.show_process()

		df.columns = ['去年同期']
		df['店别'] = value[1]
		dfs.append(df)
	df = pd.concat(dfs)
	#     总计
	sql = '''
			SELECT
				SUM( ysl )/SUM( dsl )
			FROM
				pos.dbo.jdhtab,
				pos.dbo.jbrtab,
				pos.dbo.jdstab,
				pos.dbo.jsptab
			WHERE
				zdr > 0
				AND dh# = dd#
				AND f0 = f1 % 10000
				AND bx# = dg#
				AND art < getdate ( ) AND dsl > 0
				AND dsp = sx#
				AND jlr / 1000000000 = 0
				AND dht BETWEEN '{startDate}' AND '{nextStartDate}'
	'''.format(startDate=lastMonthDate(num-12), nextStartDate=lastMonthDate(-12))
	df2 = pd.read_sql(sql, conn)

	process_bar.show_process()

	df2.columns = ['去年同期']
	df2['店别'] = '全公司'
	df = pd.concat([df, df2])
	df = pd.merge(arrivalRateDf2, df)
	df['同比'] = df.apply(lambda x: x['总计']-x['去年同期'], axis=1)
	df.set_index('店别', inplace=True)
	return df

def excel_output(writer, arrival):
	workbook = writer.book
	arrival.to_excel(writer, sheet_name='供应商到货率')
	worksheet = writer.sheets['供应商到货率']
	col = [{'header': '店别'}]
	for index, value in enumerate(arrival.columns.values):
		if index <= len(arrival.columns.values) - 4:
			value = lastMonthDate(3-len(arrival.columns.values)+index)[4:6] + '月'
		col.append({'header': value})
	worksheet.add_table('A1:' + chr(65+arrival.columns.values.size) + '9', {'columns': col, 'style': 'Table Style Medium 11'})
	common = {
		'font_name': '微软雅黑',
		'size': 16,
		'text_wrap': True,
		'align': 'center',
		'valign': 'vcenter',
	}
	cell_format = workbook.add_format({**common,'bold': True})
	cell_format_percent = workbook.add_format({**common, 'num_format': '0.00%'})
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
	# 重新写入索引
	for row_num, value in enumerate(arrival.index.values):
		worksheet.write(row_num+1, 0, value, cell_format)
	# 设置列宽
	worksheet.set_column('A:A', 8.38)
	worksheet.set_column('B:L', 10.13)
	worksheet.set_column('M:O', 11.25)
	# 设置行高
	for i in range(9):
		if i == 0:
			worksheet.set_row(0, 50, cell_format)
		else:
			worksheet.set_row(i, 50, cell_format_percent)
	# 设置条件格式
	col_num = len(arrival.columns.values)
	worksheet.conditional_format(
		1, col_num, 7, col_num,
		{
			'type': 'bottom',
			'value': 3,
			'format': cell_format_green
		}
	)
	worksheet.conditional_format(
		1, col_num-2, 7, col_num-2,
		{
			'type': 'top',
			'value': 3,
			'format': cell_format_red
		}
	)
def main():
	# 进度条
	max_steps = 92
	process_bar = ShowProcess(max_steps, 'OK')
	# 8.供应商到货率统计
	arrival = arrivalRate(STORES, -12, process_bar)
	# 保存为Excel文件名称
	writer = pd.ExcelWriter(PATH + '8.供应商到货率统计.xlsx', engine='xlsxwriter')
	excel_output(writer, arrival)
	writer.save()

if __name__ == '__main__':
	main()