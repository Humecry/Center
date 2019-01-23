# Center
## 配置文件
新建conf.py文件, 保存在同级目录里
```python
# 数据库配置
S000 = (
		'DRIVER={SQL Server};'
		'SERVER=192.168.105.1;'
		'UID=sa_;'
		'PWD=')

S001 = (
		'DRIVER={SQL Server};'
		'SERVER=192.168.2.1;'
		'UID=sa_;'
		'PWD=')

S002 = (
		'DRIVER={SQL Server};'
		'SERVER=192.168.40.1;'
		'UID=sa_;'
		'PWD=')

S003 = (
		'DRIVER={SQL Server};'
		'SERVER=192.168.2.2;'
		'UID=sa_;'
		'PWD=')

S006 = (
		'DRIVER={SQL Server};'
		'SERVER=192.168.2.3;'
		'UID=sa_;'
		'PWD=')

S007 = (
		'DRIVER={SQL Server};'
		'SERVER=192.168.50.1;'
		'UID=sa_;'
		'PWD=')

S008 = (
		'DRIVER={SQL Server};'
		'SERVER=192.168.103.1;'
		'UID=oa;'
		'PWD=oa')

S009 = (
		'DRIVER={SQL Server};'
		'SERVER=192.168.60.1;'
		'UID=sa_;'
		'PWD=')

S088 = (
		'DRIVER={SQL Server};'
		'SERVER=192.168.92.1;'
		'UID=sa_;'
		'PWD=')

S188 = (
		'DRIVER={SQL Server};'
		'SERVER=192.168.105.2;'
		'UID=sa_;'
		'PWD=')

STORES = [
	(S001, '海富'),
	(S002, '东孚'),
	(S006, '绿苑'),
	(S007, '灌口'),
	(S009, '塔埔'),
	(S088, '万达'),
	(S188, '华森'),
]

# Excel导出路径
PATH = 'excel/'
# PPT表格样式id
TABLE_STYLE_ID = '{284E427A-3D55-4303-BF80-6455036E1DE7}'
```