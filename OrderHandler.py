import datetime
import  pandas as pd
import re
import os
import sys
import yaml

def newRow(oIndex, product, quantity, consignee, phoneNo, addr, addr_abbr, sortkey):
    return {'跟团号': oIndex, '商品': product, '数量': quantity, '收货人': consignee, '联系电话': phoneNo, '地址': addr_abbr, '备注':'', '详细地址': addr, '排序键':sortkey}

def writeSheet(writer, data, sheetName, colNames):
	ordered_data = sorted(data, key=lambda book: book['排序键'], reverse=False)
	for i in range(len(ordered_data)):
		if i == 0 or ordered_data[i]['地址'] == '':
			continue
		if ordered_data[i]['地址'] == ordered_data[i-1]['地址']:
			ordered_data[i]['备注'] = '地址同上'

	df = pd.DataFrame(ordered_data)
	df.index = df.index + 1
	df.to_excel(writer, sheetName,columns=colNames, index_label='序号')
	workbook=writer.book
	workbook.add_format({"border":1,"border_color": "#000000", 'align': 'center'})
	
	cellFormat = workbook.add_format({'border':1, 'align':'left', 'valign':'vcenter'})

	worksheet = writer.sheets[sheetName]
	worksheet.set_row(0, 30)
	worksheet.set_column(0, 0, 5, cellFormat)
	worksheet.set_column(1, 1, 5, cellFormat)
	worksheet.set_column(2, 2, 25, cellFormat)
	worksheet.set_column(3, 3, 5, cellFormat)
	worksheet.set_column(4, 5, 16, cellFormat)
	worksheet.set_column(6, 6, 30, cellFormat)
	# worksheet.set_column(7, 7, 15, cellFormat)
	# worksheet.set_column(8, 8, 8, cellFormat)

	for i in range(len(data)):
		worksheet.set_row( i+1, 22)


def writeSummarySheet(writer, data, sheetName):
	df = pd.DataFrame(data)
	df.index = df.index + 1
	df.to_excel(writer, sheetName, index=False)
	workbook=writer.book
	workbook.add_format({"border":1,"border_color": "#000000", 'align': 'center'})
	
	cellFormat = workbook.add_format({'border':1, 'align':'left', 'valign':'vcenter'})

	worksheet = writer.sheets[sheetName]
	worksheet.set_row(0, 30)
	worksheet.set_column(0, 0, 30, cellFormat)
	worksheet.set_column(1, 1, 15, cellFormat)
	worksheet.set_column(2, 2, 5, cellFormat)

	for i in range(len(data)):
		worksheet.set_row( i+1, 22)

# ----------------------------------------
# Arguments
# ----------------------------------------
OrderFileDir = os.path.abspath(sys.argv[1])
RiskAddrFilePath = os.path.abspath(sys.argv[2])
OutputDirPath = os.path.abspath(sys.argv[3])
print("Order dir: ", OrderFileDir)
print("Risk address file:", RiskAddrFilePath)
print("Output dir:", OutputDirPath)

with open('config.yml', 'r') as file:
    configuration = yaml.safe_load(file)

print(configuration)
for group in configuration['groups']:
	print(group['name'])

existOutputDir = os.path.isdir(OutputDirPath)
if not existOutputDir:
	os.mkdir(OutputDirPath)

# ----------------------------------------
#  Load order files
# ----------------------------------------
OutputFileName = '配送单_{0}.xlsx'.format(datetime.datetime.now().strftime('%Y-%m-%d'))
OutputFilePath = os.path.join(OutputDirPath, OutputFileName)

# ----------------------------------------
#  Load risky building info
# ---------------------------------------- 
riskAddrs = []
riskAddrFile = pd.read_excel(RiskAddrFilePath)
riskBlocks = riskAddrFile['弄']
riskBuildings = riskAddrFile['楼栋']
for i in range(len(riskBlocks)):
	riskAddrs.append('-'.join([str(riskBlocks[i]), str(riskBuildings[i])]))

orders_363 = []
orders_828 = []
orders_1280 = []
orders_risk = []
orders_err = []
orders_summary = []
orderFiles = os.listdir(OrderFileDir)
for orderFile in orderFiles:
	orderPath = os.path.join(OrderFileDir, orderFile)
	orderFile = pd.read_excel(orderPath)
	indexes = orderFile['跟团号'].values
	if(len(indexes) == 0):
		continue
	products = orderFile['商品'].values
	quantities = orderFile['商品种类数'].values
	consignees = orderFile['收货人'].values
	phoneNos = orderFile['联系电话'].values
	addrs = orderFile['详细地址'].values

	orders_summary.append({'商品': products[0], '描述': '总计', '数量': len(indexes)})

	counter_363 = 0
	counter_828 = 0
	counter_1280 = 0
	counter_err = 0
	counter_err = 0
	counter_risk = 0

	for i in range(len(indexes)):
		sortkey = i
		addr = addrs[i]
		addr_tokens = re.findall(r'\d+', addr)
		addr = addr[7:]
		if len(addr_tokens) == 3 or len(addr_tokens) == 4:
			if len(addr_tokens) == 4 and len(re.findall(r'\d+弄\d+幢\d+号\d+', addr)) == 1:
				addr_abb = '-'.join([addr_tokens[0],addr_tokens[2],addr_tokens[3]])
				isRisk = '-'.join([addr_tokens[0],addr_tokens[2]]) in riskAddrs
				sortkey = int(''.join([addr_tokens[2],addr_tokens[3]]))
			elif len(addr_tokens) == 3:
				addr_abb = '-'.join(addr_tokens)
				isRisk = '-'.join([addr_tokens[0],addr_tokens[1]]) in riskAddrs
				sortkey = int(''.join([addr_tokens[1], addr_tokens[2]]))
			else :
				orders_err.append(newRow(indexes[i], products[i], quantities[i], consignees[i], phoneNos[i], addr,'', sortkey))
				counter_err += int(quantities[i])
				continue

			if isRisk:
				orders_risk.append(newRow(indexes[i], products[i], quantities[i],consignees[i], phoneNos[i], addr, addr_abb, sortkey))
				counter_risk += int(quantities[i])
			elif addr_tokens[0] == '363':
				orders_363.append(newRow(indexes[i], products[i], quantities[i],consignees[i], phoneNos[i], addr, addr_abb, sortkey))
				counter_363 += int(quantities[i])
			elif addr_tokens[0] == '828':
				orders_828.append(newRow(indexes[i], products[i], quantities[i],consignees[i], phoneNos[i], addr,addr_abb, sortkey))
				counter_828 += int(quantities[i])
			elif addr_tokens[0] == '1280':
				orders_1280.append(newRow(indexes[i], products[i], quantities[i],consignees[i], phoneNos[i], addr,addr_abb, sortkey))
				counter_1280 += int(quantities[i])
			else :
				orders_err.append(newRow(indexes[i], products[i], quantities[i],consignees[i], phoneNos[i], addr, '', sortkey))
				counter_err += int(quantities[i])
		else:
			orders_err.append(newRow(indexes[i], products[i], quantities[i], consignees[i], phoneNos[i], addr,'', sortkey))
			counter_err += int(quantities[i])
	
	orders_summary.append({'商品': '', '描述': '363弄', '数量': counter_363})
	orders_summary.append({'商品': '', '描述': '828弄', '数量': counter_828})
	orders_summary.append({'商品': '', '描述': '1280弄', '数量': counter_1280})
	orders_summary.append({'商品': '', '描述': '封控楼栋', '数量': counter_risk})
	orders_summary.append({'商品': '', '描述': '地址错误', '数量': counter_err})
	orders_summary.append({'商品': '', '描述': '', '数量': ''})
	
		
writer = pd.ExcelWriter(OutputFilePath, engine='xlsxwriter')
writeSheet(writer, orders_363, '363弄', ['跟团号', '商品', '数量', '收货人','地址','备注'])
writeSheet(writer, orders_828, '828弄', ['跟团号', '商品', '数量','收货人','地址','备注'])
writeSheet(writer, orders_1280, '1280弄', ['跟团号', '商品', '数量', '收货人','地址','备注'])
writeSheet(writer, orders_risk, '封控楼栋', ['跟团号', '商品', '数量', '收货人','地址','备注'])
writeSheet(writer, orders_err, '地址错误', ['跟团号', '商品', '数量', '收货人','详细地址','备注'])
writeSummarySheet(writer, orders_summary, '订单数量汇总')

writer.save()