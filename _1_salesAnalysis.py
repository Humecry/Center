#! /usr/bin/env python3
# -*- coding:utf-8 -*-

# @Date    : 2018-08-03 16:44:00
# @Author  : Hume (102734075@qq.com)
# @Link    : https://humecry.wordpress.com/
# @Version : 1.1
# @Description: 1.消费数据分析

from common import *

'''
消费数据分析
'''
def shoppingCardMember(process_bar):
	# 统计日期
	countDate = {'start':[], 'end':[], 'nextStart':[], 'last3start':[], 'lastHalfEnd':[]}
	# 上个月同期，上上个月日期，上个月日期
	for i in (-13, -2, -1):
		# 上月月第一天
		countDate['start'].append(lastMonthDate(i))
		# 上月最后一天
		countDate['end'].append(lastMonthDate(i, 'end'))
		# 本月第一天
		countDate['nextStart'].append(lastMonthDate(i+1))
		# 上月的前三个月的第一天
		countDate['last3start'].append(lastMonthDate(i-2))
		# 上月的半年前的第一天
		countDate['lastHalfEnd'].append(lastMonthDate(i-6, 'end') + ' 23:59')
	countDate['end'] = tuple(countDate['end'])
	# 连接003号金卡店的数据库
	conn = pyodbc.connect(S003)

	# 购物卡余额与会员卡零钱包
	# fd字段为分店代号
	sql1 = '''
			SELECT
				CONVERT(INT, ye + 0.5) AS [购物卡余额],
				CONVERT(INT, lqye + 0.5) AS [会员卡零钱包] 
			FROM
				pos.dbo.jckrtab 
			WHERE
				fd = 3 
				AND krq IN {countDate}
	'''.format(countDate=countDate['end'])
	df1 = pd.read_sql(sql1, conn)

	process_bar.show_process()

	# 充值额,消费会员的消费额与会员数，半年未消费
	df2 = []
	df3 = []
	df4 = []
	df5 = []
	zipData = zip(countDate['start'],
				  countDate['end'],
				  countDate['nextStart'],
				  countDate['last3start'],
				  countDate['lastHalfEnd'])
	for (i, j, k, t, l) in zipData:
		# 充值额
		sql2 = '''
				SELECT 
					CONVERT(INT, SUM ( tzje ) + SUM ( kkje ) + 0.5) AS [充值额] 
				FROM
					pos.dbo.jckrtab 
				WHERE
					fd < 10000 
					AND krq BETWEEN '{start}' 
					AND '{end}'
		 '''.format(start=i, end=j)
		df2.append(pd.read_sql(sql2, conn))

		process_bar.show_process()

		# 最近一个月有消费会员
		sql3 = '''
				SELECT
					CONVERT(INT, SUM ( je ) + 0.5) AS [会员消费额(一个月有消费)],
					COUNT( DISTINCT k# ) AS [会员数(一个月有消费)]
				FROM
					pos.dbo.jhjltab
				WHERE
					rt BETWEEN '{start}' AND '{nextStart}'
		'''.format(start=i, nextStart=k)
		df3.append(pd.read_sql(sql3, conn))

		process_bar.show_process()

		# 最近三个月有消费会员
		sql4 = '''
				SELECT
					CONVERT(INT, SUM ( je ) + 0.5) AS [会员消费额(三个月有消费)],
					COUNT( DISTINCT k# ) AS [会员数(三个月有消费)]
				FROM
					pos.dbo.jhjltab
				WHERE
					rt BETWEEN '{last3start}' AND '{nextStart}'
		 '''.format(last3start=t, nextStart=k)
		df4.append(pd.read_sql(sql4, conn))

		process_bar.show_process()

		# 半年前未消费会员
		sql5 = '''
				SELECT
					COUNT( DISTINCT h# ) AS [会员数(超过半年未消费)]
				FROM
					pos.dbo.jhdjtab
				WHERE
					csrq BETWEEN '20010101'
					AND '{lastHalfEnd}'
		'''.format(lastHalfEnd=l)
		df5.append(pd.read_sql(sql5, conn))

		process_bar.show_process()

	df2 = pd.concat(df2, ignore_index=True)
	df3 = pd.concat(df3, ignore_index=True)
	df4 = pd.concat(df4, ignore_index=True)
	df5 = pd.concat(df5, ignore_index=True)
	# 多表拼接
	cardMemberDf = df1.join([df2, df3, df4, df5])
	# 增加统计日期列
	cardMemberDf.insert(0, '统计日期', countDate['end'])
	# 设置统计日期为索引
	cardMemberDf.set_index('统计日期', inplace=True)
	# 求环比
	a = cardMemberDf.apply(lambda x: (x[lastMonthDate(-1, 'end')] - x[lastMonthDate(-2, 'end')])/x[lastMonthDate(-2, 'end')])
	# 求同比
	b = cardMemberDf.apply(lambda x: (x[lastMonthDate(-1, 'end')] - x[lastMonthDate(-13, 'end')])/x[lastMonthDate(-13, 'end')])
	cardMemberDf.loc['环比'] = a
	cardMemberDf.loc['同比'] = b
	cardMemberDf.insert(2, '乐海卡余额', 0)
	# 列排序
	cardMemberDf = cardMemberDf[[
					'购物卡余额',
					'充值额',
					'乐海卡余额',
					'会员卡零钱包',
					'会员消费额(一个月有消费)',
					'会员数(一个月有消费)',
					'会员消费额(三个月有消费)',
					'会员数(三个月有消费)',
					'会员数(超过半年未消费)']]
	# 行列转置
	cardMemberDf = cardMemberDf.stack().unstack(0)
	# 列排序
	cardMemberDf = cardMemberDf[[lastMonthDate(-1, 'end'), lastMonthDate(-2, 'end'), '环比', lastMonthDate(-13, 'end'), '同比']]
	cardMemberDf.rename(
							columns={
										cardMemberDf.columns[0]: '本期' + lastMonthDate(-1, 'end')[4:6] + '月',
										cardMemberDf.columns[1]: '上期' + lastMonthDate(-2, 'end')[4:6] + '月',
										cardMemberDf.columns[3]: '去年同期' + lastMonthDate(-13, 'end')[4:6] + '月'
										},
							inplace=True)
	return cardMemberDf

def excel_output(writer, shopping):
	workbook = writer.book
	# 购物卡/会员数据分析
	shopping.to_excel(writer, sheet_name='购物卡会员数据分析')
	worksheet = writer.sheets['购物卡会员数据分析']
	col = [{'header': '内容'}]
	for value in shopping.columns.values:
		col.append({'header': value})
	worksheet.add_table('A1:F10', {'columns': col, 'style': 'Table Style Medium 11'})
	# 单元格格式
	common = {
		'font_name': '微软雅黑',
		'font_size': 16,
		'text_wrap': True,
		'align': 'center',
		'valign': 'vcenter'
	}
	cell_format = workbook.add_format({
		**common,
		'bold': True,
	})
	cell_format_num = workbook.add_format({
		**common,
		'num_format': '00',
	})
	cell_format_percent = workbook.add_format({
		**common,
		'num_format': '0.00%',
	})
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
	# 设置条件格式
	worksheet.conditional_format(
		'D2:D10',
		{
			'type': 'cell',
			'criteria': '<',
			'value': 0,
			'format': cell_format_green
		}
	)
	worksheet.conditional_format(
		'F2:F10',
		{
			'type': 'cell',
			'criteria': '<',
			'value': 0,
			'format': cell_format_green
		}
	)
	worksheet.conditional_format(
		'D2:D10',
		{
			'type': 'cell',
			'criteria': '>',
			'value': 0,
			'format': cell_format_red
		}
	)
	worksheet.conditional_format(
		'F2:F10',
		{
			'type': 'cell',
			'criteria': '>',
			'value': 0,
			'format': cell_format_red
		}
	)
	# 重新写入索引
	for row_num, value in enumerate(shopping.index.values):
		worksheet.write(row_num+1, 0, value, cell_format)
	# 设置列宽
	worksheet.set_column('A:A', 38, cell_format)
	worksheet.set_column('B:C', 26, cell_format_num)
	worksheet.set_column('E:E', 26, cell_format_num)
	worksheet.set_column('D:D', 20, cell_format_percent)
	worksheet.set_column('F:F', 20, cell_format_percent)
	# 设置行高
	for i in range(10):
		if i == 0:
			worksheet.set_row(i, 52)
		else:
			worksheet.set_row(i, 44)

def ppt_output(prs, shopping):
	# 1.消费数据分析
	slide_layout = prs.slide_layouts[1]
	slide = prs.slides.add_slide(slide_layout)
	title = slide.shapes.title
	# 设置标题位置
	title.top = Inches(0.2)
	title.width = Inches(8)
	title.height = Inches(1)
	# 标题文本
	title.text = "  一. 购物卡/会员数据分析"
	# 标题格式
	ppt_layout(title, bold=True, color=(255, 255, 255))
	rows = shopping.index.size
	cols = shopping.columns.size
	left, top, width, height = Inches(0), Inches(1.34), Inches(13.33), Inches(5.3)
	# 设置PPT表格样式
	CT_Table.new_tbl = partial(CT_Table.new_tbl, tableStyleId=TABLE_STYLE_ID)
	# 新建PPT表格
	table = slide.shapes.add_table(rows+1, cols+1, left, top, width, height).table
	# 写入列名
	for index, value in enumerate(shopping.columns.values):
		ppt_layout(table.cell(0, index+1), value, bold=True, color=(255, 255, 255))
	# 写入行标题
	for index, value in enumerate(shopping.index.values):
		ppt_layout(table.cell(index+1, 0), value, bold=True)
	# 写入数据
	for r in range(rows):
		for c in range(cols):
			if c == 0 or c == 1 or c == 3:
				ppt_layout(table.cell(r+1, c+1), '{:.0f}'.format(shopping.iat[r, c]))
			else:
				if shopping.iat[r, c] < 0:
					color=(0, 128, 58)
				elif shopping.iat[r, c] > 0:
					color=(255, 0, 0)
				else:
					color=(0, 0, 0)
				ppt_layout(table.cell(r+1, c+1), '{:.2%}'.format(shopping.iat[r, c]), bold=True, color=color)

def main():
	# 进度条
	max_steps = 13
	process_bar = ShowProcess(max_steps, 'OK')
	# 1.消费数据分析
	shopping = shoppingCardMember(process_bar)

	# 导出Excel
	writer = pd.ExcelWriter(PATH + '1.消费数据分析.xlsx', engine='xlsxwriter')
	excel_output(writer, shopping)
	writer.save()
	
	# 导出PPT
	# prs = Presentation('模板.pptx')
	# ppt_output(prs, shopping)
	# prs.save('1.消费数据分析.pptx')

if __name__ == '__main__':
	main()