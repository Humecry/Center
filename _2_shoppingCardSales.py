#! /usr/bin/env python3
# -*- coding:utf-8 -*-

# @Date    : 2018-08-03 17:45:03
# @Author  : Hume (102734075@qq.com)
# @Link    : https://humecry.wordpress.com/
# @Version : 1.1
# @Description: 2.购物卡消费统计: 购物卡消费趋势, 购物卡消费店别排名

from common import *

'''
购物卡消费额与全店业绩
'''
def shoppingCard(stores, start, end, process_bar):
	# 乐海卡6个月消费金额
	dfs = []
	for value in stores:
		cardDf = pd.DataFrame()
		conn = pyodbc.connect(value[0])
		storeDf = pd.DataFrame(data={'店别': [value[1]]})
		for j in range(start, end, 1):
			# 储卡
			sql = '''
					SELECT
						SUM ( je )
					FROM
						pos.dbo.sydbak,
						pos.dbo.sybbak
					WHERE
						s#= b#
						AND b_ % 100 = 2
						AND ct BETWEEN '{startDate}'
						AND '{nextStartDate}'
						AND zcd % 1000 = 0
			'''.format(startDate=lastMonthDate(j), nextStartDate=lastMonthDate(j+1))
			df = pd.read_sql(sql, conn)

			process_bar.show_process()

			df = pd.merge(storeDf, df, left_index=True, right_index=True)
			# 乐海卡
			sql2 = '''
					SELECT
						SUM ( je )
					FROM
						pos.dbo.sydbak,
						pos.dbo.sybbak
					WHERE
						s#= b#
						AND b_ % 100 IN ( 3, 13, 15 ) 
						AND ct BETWEEN '{startDate}'
						AND '{nextStartDate}'
						AND zcd % 1000 = 0
			'''.format(startDate=lastMonthDate(j), nextStartDate=lastMonthDate(j+1))
			df2 = pd.read_sql(sql2, conn)

			process_bar.show_process()

			df2 = pd.merge(df, df2, left_index=True, right_index=True)
			df2.columns = [
				['', '消费金额', '消费金额'],
				['店别', lastMonthDate(j) + '储卡', lastMonthDate(j) + '乐海卡']
			]
			cardDf = pd.merge(cardDf, df2) if not cardDf.empty else df2
		dfs.append(cardDf)
	leHaiCardDf = pd.concat(dfs)
	leHaiCardDf.reset_index(drop=True, inplace=True)
	# 全店6个月业绩
	performanceDf = pd.DataFrame()
	conn = pyodbc.connect(S008)
	storeDf = pd.DataFrame(data={'店别': [value[1]]})
	for i in range(start, end, 1):
		sql = '''
				SELECT
					lqm [店别],
					SUM( zxs * ( 1- ( 1- sfl ) * 0 ) ) [销售额]
				FROM
					pos.dbo.s_xszb,
					pos.dbo.yblb,
					pos.dbo.jlqtab
				WHERE
					zrq = bly
					AND zsn < 10000
					AND lb = 2
					AND c = zsn % 1000
					AND zrq BETWEEN '{startDate}'   AND '{endDate}'
					AND zsn / 1000 % 10 < 2
					AND zsn % 1000 IN (1, 2, 6, 7, 9, 88, 188)
				GROUP BY
					lqm,
					zsn % 1000
				ORDER BY
					zsn % 1000
		'''.format(startDate=lastMonthDate(i), endDate=lastMonthDate(i, 'end'))
		df = pd.read_sql(sql, conn)

		process_bar.show_process()

		df2 = pd.DataFrame(df['销售额'])
		df2.columns = [[''], [lastMonthDate(i)+'全店业绩']]
		leHaiCardDf = pd.merge(leHaiCardDf, df2, left_index=True, right_index=True)
	leHaiCardDf.loc[7] = leHaiCardDf.apply(lambda x: x.sum())
	leHaiCardDf['', '店别'].iat[7] = '全公司'
	return leHaiCardDf

'''
购物卡消费趋势
'''
def shoppingTrend(thisYearDf, lastYearDf):
	months = ['月份']
	for i in range(-6, 0, 1):
		months.append(lastMonthDate(i)[4:6]+'月')
	# 6个月业绩
	performance = thisYearDf.iloc[7:8, 13:19]
	performance.insert(0, '月份', '业绩')
	performance.columns = months
	# 6个月储卡消费
	valueCard = thisYearDf.iloc[7:8, [1, 3, 5, 7, 9, 11]]
	valueCard.insert(0, '月份', '储卡消费额')
	valueCard.columns = months
	# 6个月乐海卡消费
	leHaiCard = thisYearDf.iloc[7:8, [2, 4, 6, 8, 10, 12]]
	leHaiCard.insert(0, '月份', '乐海卡消费额')
	leHaiCard.columns = months
	# 去年同期储卡消费
	valueCardLastYear = lastYearDf.iloc[7:8, [1, 3, 5, 7, 9, 11]]
	valueCardLastYear.insert(0, '月份', '去年储卡消费额')
	valueCardLastYear.columns = months
	# 去年同期业绩
	performanceLastYear = lastYearDf.iloc[7:8, 13:19]
	performanceLastYear.insert(0, '月份', '去年业绩')
	performanceLastYear.columns = months
	# 建立索引
	shoppingCardTrendDf = pd.concat([performance, valueCard, leHaiCard, valueCardLastYear, performanceLastYear])
	shoppingCardTrendDf.set_index('月份', inplace=True)
	# 储卡占比
	valueCardProportion = shoppingCardTrendDf.apply(lambda x: x['储卡消费额']/x['业绩'])
	shoppingCardTrendDf.loc['储卡占比'] = valueCardProportion
	# 乐海卡占比
	leHaiCardProportion = shoppingCardTrendDf.apply(lambda x: x['乐海卡消费额']/x['业绩'])
	shoppingCardTrendDf.loc['乐海卡占比'] = leHaiCardProportion
	# 去年同期储卡消费额
	valueCardLastYear = shoppingCardTrendDf.apply(lambda x: x['去年储卡消费额']/x['去年业绩'])
	shoppingCardTrendDf.loc['去年占比'] = valueCardLastYear
	# 业绩
	sales = shoppingCardTrendDf.apply(lambda x: int(x['业绩']+0.5))
	shoppingCardTrendDf.loc['业绩'] = sales
	# 储卡消费额
	valueCard = shoppingCardTrendDf.apply(lambda x: int(x['储卡消费额']+0.5))
	shoppingCardTrendDf.loc['储卡消费额'] = valueCard
	# 去年储卡消费额
	valueCardLastYear = shoppingCardTrendDf.apply(lambda x: int(x['去年储卡消费额']+0.5))
	shoppingCardTrendDf.loc['去年储卡消费额'] = valueCardLastYear
	shoppingCardTrendDf.drop(['乐海卡消费额', '去年业绩'], inplace=True)
	index = ['业绩', '储卡占比', '乐海卡占比', '储卡消费额', '去年储卡消费额', '去年占比']
	shoppingCardTrendDf = shoppingCardTrendDf.reindex(index)
	return shoppingCardTrendDf

'''
购物卡消费店别排名
'''
def shoppingCardRank(thisYearDf):
	# 店别
	shopDf = thisYearDf.loc[:, [('', '店别')]]
	# 上月储卡消费
	valueCardDf = thisYearDf.iloc[:, [11]]
	# 上月乐海卡消费
	leHaiCardDf = thisYearDf.iloc[:, [12]]
	shoppingCardRankDf = pd.merge(shopDf, valueCardDf, left_index=True, right_index=True)
	shoppingCardRankDf = pd.merge(shoppingCardRankDf, leHaiCardDf, left_index=True, right_index=True)
	for i in range(3):
		# 前三个月储卡业绩占比
		valueCardProportion = thisYearDf.apply(lambda x: x[2*i+7]/x[16+i], axis=1)
		shoppingCardRankDf['储卡业绩占比',lastMonthDate(-3+i)[4:6]+'月'] = valueCardProportion
	for i in range(3):
		# 前三个月储卡+乐海卡业绩占比
		cardProportion = thisYearDf.apply(lambda x: (x[2*i+7]+x[2*i+8])/x[16+i], axis=1)
		shoppingCardRankDf['储卡+乐海卡业绩占比', lastMonthDate(-3+i)[4:6]+'月'] = cardProportion
	linkRelativeDf = shoppingCardRankDf.apply(lambda x: x['储卡+乐海卡业绩占比'][2]-x['储卡+乐海卡业绩占比'][1], axis=1)
	shoppingCardRankDf['储卡+乐海卡环比'] = linkRelativeDf
	# 去掉全公司合计
	shoppingCardRankDf = shoppingCardRankDf.iloc[:7, 0:10]
	# 上月储卡+乐海卡业绩占比排名
	rankDf = shoppingCardRankDf.loc[:, [('储卡+乐海卡业绩占比', lastMonthDate(-1)[4:6]+'月')]].rank(method='first', ascending=False)
	shoppingCardRankDf['综合占比排名'] = rankDf
	return shoppingCardRankDf

def excel_output(writer, trend, rank):
	workbook = writer.book
	'''
	购物卡消费趋势
	'''
	trend.to_excel(writer, sheet_name='购物卡消费趋势')
	worksheet = writer.sheets['购物卡消费趋势']
	col = [{'header': '项目'}]
	for value in trend.columns.tolist():
		col.append({'header': value})
	worksheet.add_table('A1:G7', {'columns': col, 'style': 'Table Style Medium 11'})
	common = {
		'font_name': '微软雅黑',
		'text_wrap': True,
		'align':'center',
		'valign': 'vcenter',
	}
	cell_format = workbook.add_format(common)
	cell_format_num = workbook.add_format({
		**common,
		'num_format': '00',
	})
	cell_format_percent = workbook.add_format({
		**common,
		'num_format': '0.00%',
	})
	worksheet.set_row(0, None, cell_format)
	for i in (1, 4, 5):
		worksheet.set_row(i, None, cell_format_num)
	for i in (2, 3, 6):
		worksheet.set_row(i, None, cell_format_percent)
	# 设置列宽
	worksheet.set_column('A:G', 16)
	# 画图
	chart = workbook.add_chart({'type': 'column'})
	font_common = {
		'name': '微软雅黑',
		'size': 14
	}
	chart.add_series({
		'name': '购物卡消费趋势!$A$2',
		'categories': '=购物卡消费趋势!$B$1:$G$1',
		'values':     '=购物卡消费趋势!$B$2:$G$2',
		# 'data_labels': {'value': True, 'position': 'center', 'font': font_common},
	})
	chart.add_series({
		'name': '购物卡消费趋势!$A$5',
		'categories': '=购物卡消费趋势!$B$1:$G$1',
		'values':     '=购物卡消费趋势!$B$5:$G$5',
		# 'data_labels': {'value': True, 'position': 'center', 'font': font_common},
	})
	chart.add_series({
		'name': '购物卡消费趋势!$A$6',
		'categories': '=购物卡消费趋势!$B$1:$G$1',
		'values':     '=购物卡消费趋势!$B$6:$G$6',
		# 'data_labels': {'value': True, 'position': 'center', 'font': font_common},
	})

	line_chart = workbook.add_chart({'type': 'line'})
	line_chart.add_series({
		'name': '购物卡消费趋势!$A$3',
		'categories': '=购物卡消费趋势!$B$1:$G$1',
		'values':     '=购物卡消费趋势!$B$3:$G$3',
		# 'data_labels': {'value': True, 'position': 'center', 'font': {**font_common, 'color': '#10243f'}},
		'y2_axis': True,
	})
	line_chart.add_series({
		'name': '购物卡消费趋势!$A$4',
		'categories': '=购物卡消费趋势!$B$1:$G$1',
		'values':     '=购物卡消费趋势!$B$4:$G$4',
		# 'data_labels': {'value': True, 'font': {**font_common, 'color': '#376092'}},
		'y2_axis': True,
	})
	line_chart.add_series({
		'name': '购物卡消费趋势!$A$7',
		'categories': '=购物卡消费趋势!$B$1:$G$1',
		'values':     '=购物卡消费趋势!$B$7:$G$7',
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
	'''
	购物卡消费店别排名
	'''
	rank.to_excel(writer, sheet_name='购物卡消费店别排名', header=False)
	worksheet = writer.sheets['购物卡消费店别排名']
	common = {
		'font_name': '微软雅黑',
		'font_size': 16,
		'text_wrap': True,
		'align': 'center',
		'valign': 'vcenter',
		'border': 1,
		'border_color': '#FFFFFF',
	}
	cell_format = workbook.add_format({
		**common,
		'font_size': 20,
		'font_color': '#FFFFFF',
		'bold': True,
		'bg_color': '#9bbb59' ,
	})
	cell_format_num = workbook.add_format({
		**common,
		'bg_color': '#d8e4bc',
		'num_format': '0',
	})
	cell_format_num_even = workbook.add_format({
		**common,
		'bg_color': '#ebf1de',
		'num_format': '0',
	})
	cell_format_percent = workbook.add_format({
		**common,
		'bg_color': '#d8e4bc',
		'num_format': '0.00%',
	})
	cell_format_percent_even = workbook.add_format({
		**common,
		'bg_color': '#ebf1de',
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
	# 设置标题
	worksheet.merge_range(0, 0, 2, 0, '店别', cell_format)
	worksheet.merge_range(0, 1, 1, 2, '消费金额', cell_format)
	worksheet.merge_range(0, 3, 0, 8, '业绩占比', cell_format)
	worksheet.merge_range(1, 3, 1, 5, '储卡', cell_format)
	worksheet.merge_range(1, 6, 1, 8, '储卡+乐海卡', cell_format)
	worksheet.merge_range(0, 9, 2, 9, '储卡+乐海卡环比', cell_format)
	worksheet.merge_range(0, 10, 2, 10, '综合占比排名', cell_format)    
	worksheet.write(2, 1, '储卡', cell_format)
	worksheet.write(2, 2, '乐海卡', cell_format)
	for i in range(3):
		worksheet.write(2, i+3, lastMonthDate(-3+i)[4:6]+'月', cell_format)
		worksheet.write(2, i+6, lastMonthDate(-3+i)[4:6]+'月', cell_format)
	# 设置数字格式
	for col_num, col in enumerate(rank.columns.values):
		for row_num, value in enumerate(rank[col]):
			if col_num == 0:
				worksheet.write(row_num+3, col_num, value, cell_format)
			elif col_num in (1, 2, 10):
				if (row_num + 1) % 2:
					worksheet.write(row_num+3, col_num, value, cell_format_num)
				else:
					worksheet.write(row_num+3, col_num, value, cell_format_num_even)
			else:
				if (row_num + 1) % 2:
					worksheet.write(row_num+3, col_num, value, cell_format_percent)
				else:
					worksheet.write(row_num+3, col_num, value, cell_format_percent_even)
	for row_num in range(10):
		worksheet.write_blank(row_num, 11, '', workbook.add_format())
	# 设置条件格式
	worksheet.conditional_format(
		'J4:J10',
		{
			'type': 'cell',
			'criteria': '<',
			'value': 0,
			'format': cell_format_green
		}
	)
	worksheet.conditional_format(
		'J4:J10',
		{
			'type': 'cell',
			'criteria': '>',
			'value': 0,
			'format': cell_format_red
		}
	)
	# 设置列宽
	worksheet.set_column('A:A', 10)
	worksheet.set_column('B:C', 16)
	worksheet.set_column('D:J', 14)
	worksheet.set_column('K:K', 12)
	# 设置行高
	for i in range(10):
		if i in (0, 1, 2):
			worksheet.set_row(i, 36)
		else:
			worksheet.set_row(i, 48)

def ppt_output(prs, trend, rank):
	slide_layout = prs.slide_layouts[1]
	# 1.购物卡/会员数据分析
	slide = prs.slides.add_slide(slide_layout)
	title = slide.shapes.title
	# 设置标题位置
	title.top = Inches(0.2)
	title.width = Inches(8)
	title.height = Inches(1)
	# 标题文本
	title.text = "购物卡消费店别排名"
	# 标题格式
	ppt_layout(title, bold=True, color=(255, 255, 255))
	rows = rank.index.size
	cols = rank.columns.size
	left, top, width, height = Inches(0), Inches(1.34), Inches(13.33), Inches(6.15)
	# 设置PPT表格样式
	CT_Table.new_tbl = partial(CT_Table.new_tbl, tableStyleId=TABLE_STYLE_ID)
	# 新建PPT表格
	table = slide.shapes.add_table(rows+3, cols, left, top, width, height).table
	table.first_row = False
	# 合并单元格
	mergeCellsVertically(table=table, start_row_idx=0, end_row_idx=2, col_idx=0)
	mergeCellsVertically(table=table, start_row_idx=0, end_row_idx=2, col_idx=9)
	mergeCellsVertically(table=table, start_row_idx=0, end_row_idx=2, col_idx=10)
	mergeCellsHorizontally(table=table, row_idx=0, start_col_idx=3, end_col_idx=8)
	mergeCellsHorizontally(table=table, row_idx=1, start_col_idx=3, end_col_idx=5)
	mergeCellsHorizontally(table=table, row_idx=1, start_col_idx=6, end_col_idx=8)
	mergeCells(table=table, start_row_idx=0, end_row_idx=1, start_col_idx=1, end_col_idx=2)
	# 写入列名
	ppt_layout(table.cell(0, 0), '店别', bold=True)
	ppt_layout(table.cell(0, 1), '消费金额', bold=True)
	ppt_layout(table.cell(0, 3), '业绩占比', bold=True)
	ppt_layout(table.cell(0, 9), '储卡+乐海卡环比', bold=True)
	ppt_layout(table.cell(0, 10), '综合占比排名', bold=True)
	ppt_layout(table.cell(1, 3), '储卡', bold=True)
	ppt_layout(table.cell(1, 6), '储卡+乐海卡', bold=True)
	ppt_layout(table.cell(2, 1), '储卡', bold=True)
	ppt_layout(table.cell(2, 2), '乐海卡', bold=True)
	for i in range(3):
		ppt_layout(table.cell(2, i+3), lastMonthDate(i-3)[4:6]+'月', bold=True)
		ppt_layout(table.cell(2, i+6), lastMonthDate(i-3)[4:6]+'月', bold=True)
	# 单元格填充颜色
	table.cell(1, 3).fill.solid()
	table.cell(1, 3).fill.fore_color.rgb = RGBColor(248, 186, 147)
	table.cell(1, 6).fill.solid()
	table.cell(1, 6).fill.fore_color.rgb = RGBColor(248, 186, 147)
	# 写入数据
	for r in range(rows):
		for c in range(cols):
			cell = table.cell(r+3, c)
			value = rank.iat[r, c]
			if c == 0:
				ppt_layout(cell, value, bold=True)
			elif c in (1, 2, 10):
				value = '{:.0f}'.format(rank.iat[r, c])
				ppt_layout(cell, value)
			elif c == 9:
				value = '{:.2%}'.format(rank.iat[r, c])
				if rank.iat[r, c] < 0:
					ppt_layout(cell, value, bold=True, color=(0, 128, 58))
				elif rank.iat[r, c] > 0:
					ppt_layout(cell, value, bold=True, color=(255, 0, 0))
				else:
					ppt_layout(cell, value, bold=True)
			else:
				value = '{:.2%}'.format(rank.iat[r, c])
				ppt_layout(cell, value)

def main():
	# 进度条
	max_steps = 180
	process_bar = ShowProcess(max_steps, 'OK')
	# 2.购物卡消费统计: 购物卡消费趋势, 购物卡消费店别排名
	thisYearDf = shoppingCard(STORES, -6, 0, process_bar)
	lastYearDf = shoppingCard(STORES, -18, -12, process_bar)
	# 购物卡消费趋势
	trend = shoppingTrend(thisYearDf, lastYearDf)
	# 购物卡消费店别排名
	rank = shoppingCardRank(thisYearDf)

	# 导出Excel
	writer = pd.ExcelWriter(PATH + '2.购物卡消费统计.xlsx', engine='xlsxwriter')
	excel_output(writer, trend, rank)
	writer.save()

	# 导出PPT
	# prs = Presentation('模板.pptx')
	# ppt_output(prs, trend, rank)
	# prs.save('2.购物卡消费统计.pptx')

if __name__ == '__main__':
	main()