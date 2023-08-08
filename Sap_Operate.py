import sys
import chicon  # 引用图标
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from Sap_Operate_Ui import *
from Get_Data import *
from File_Operate import *
from PDF_Operate import *
from Sap_Operate_Ui import Ui_MainWindow


class MyMainWindow(QMainWindow, Ui_MainWindow):
	def __init__(self, parent=None):
		super(MyMainWindow, self).__init__(parent)
		self.setupUi(self)

		self.actionExport.triggered.connect(self.exportConfig)
		self.actionImport.triggered.connect(self.importConfig)
		self.actionExit.triggered.connect(MyMainWindow.close)
		self.actionHelp.triggered.connect(self.showVersion)
		self.actionAuthor.triggered.connect(self.showAuthorMessage)
		self.pushButton_11.clicked.connect(self.sapOperate)
		self.pushButton_12.clicked.connect(self.textBrowser.clear)
		self.pushButton_20.clicked.connect(self.textBrowser_2.clear)
		self.pushButton_16.clicked.connect(self.getFileUrl)
		self.pushButton_18.clicked.connect(self.getODMDataFileUrl)
		self.pushButton_23.clicked.connect(self.getCombineFileUrl)
		self.pushButton_24.clicked.connect(self.getLogFileUrl)
		self.pushButton_17.clicked.connect(self.odmDataToSap)
		self.pushButton_19.clicked.connect(self.odmCombineData)
		self.pushButton_25.clicked.connect(self.orderMergeProject)
		self.pushButton_36.clicked.connect(self.splitOdmData)
		self.pushButton_34.clicked.connect(self.textBrowser_3.clear)
		self.pushButton_35.clicked.connect(self.pdfOperate)
		self.pushButton_33.clicked.connect(self.getFiles)
		self.lineEdit_15.textChanged.connect(self.lineEditChange)
		# self.pushButton_17.clicked.connect(self.odmDataToSap1)
		self.doubleSpinBox_2.valueChanged.connect(self.getAmountVat)
		# self.pushButton_12.clicked.connect(self.lineEdit.clear)
		# self.pushButton_12.clicked.connect(self.lineEdit.clear)
		# self.pushButton_12.clicked.connect(self.lineEdit.clear)
		self.checkBox_9.toggled.connect(lambda: self.pdfNameRule('Invoice No'))
		self.checkBox_10.toggled.connect(lambda: self.pdfNameRule('Company Name'))
		self.checkBox_12.toggled.connect(lambda: self.pdfNameRule('Order No'))
		self.checkBox_11.toggled.connect(lambda: self.pdfNameRule('Project No'))
		self.filesUrl = []

	def getConfig(self):
		# 初始化，获取或生成配置文件
		global configFileUrl
		global desktopUrl
		global now
		global last_time
		global today
		global oneWeekday
		global fileUrl

		date = datetime.datetime.now() + datetime.timedelta(days=1)
		now = int(time.strftime('%Y'))
		last_time = now - 1
		today = time.strftime('%Y.%m.%d')
		oneWeekday = (datetime.datetime.now() + datetime.timedelta(days=7)).strftime('%Y.%m.%d')
		desktopUrl = os.path.join(os.path.expanduser("~"), 'Desktop')
		configFileUrl = '%s/config' % desktopUrl
		configFile = os.path.exists('%s/config_sap.csv' % configFileUrl)
		# print(desktopUrl,configFileUrl,configFile)
		if not configFile:  # 判断是否存在文件夹如果不存在则创建为文件夹
			reply = QMessageBox.question(self, '信息', '确认是否要创建配置文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
			if reply == QMessageBox.Yes:
				if not os.path.exists(configFileUrl):
					os.makedirs(configFileUrl)
				MyMainWindow.createConfigContent(self)
				MyMainWindow.getConfigContent(self)
				self.textBrowser.append("创建并导入配置成功")
			else:
				exit()
		else:
			MyMainWindow.getConfigContent(self)

	def getConfigContent(self):
		# 配置文件
		csvFile = pd.read_csv('%s/config_sap.csv' % configFileUrl, names=['A', 'B', 'C'])
		global configContent
		global username
		global role
		configContent = {}
		username = list(csvFile['A'])
		number = list(csvFile['B'])
		role = list(csvFile['C'])
		for i in range(len(username)):
			configContent['%s' % username[i]] = number[i]
		# a = len(configContent)
		# if (int(configContent['config_num']) != len(configContent)) or (len(configContent) != 40):
		# 	reply = QMessageBox.question(self, '信息', 'config文件配置缺少一些参数，是否重新创建并获取新的config文件', QMessageBox.Yes | QMessageBox.No,
		# 								 QMessageBox.Yes)
		# 	if reply == QMessageBox.Yes:
		# 		MyMainWindow.createConfigContent(self)
		# 		MyMainWindow.getConfigContent(self)
		MyMainWindow.csItem(self)
		MyMainWindow.salesItem(self)
		MyMainWindow.getDefaultInformation(self)
		MyMainWindow.getInvoiceMsg(self)

		try:
			self.textBrowser_2.append("配置获取成功")
		except AttributeError:
			QMessageBox.information(self, "提示信息", "已获取配置文件内容", QMessageBox.Yes)
		else:
			pass

	def createConfigContent(self):
		global monthAbbrev
		months = "JanFebMarAprMayJunJulAugSepOctNovDec"
		n = time.strftime('%m')
		pos = (int(n) - 1) * 3
		monthAbbrev = months[pos:pos + 3]

		configContent = [
			# ['SAP登入信息', '内容', '备注'],
			# ['订单类型', 'DR', '根据Site自定义'],
			# ['销售组织', '0486', '根据Site自定义'],
			# ['分销渠道', '01', '根据Site自定义'],
			# ['销售办事处', '>601', '根据Site自定义'],
			# ['销售组', '240', '根据Site自定义'],、
			['特殊开票', '内容', '备注'],
			['SAP_Date_URL', 'N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\收样\\3.Sap\\ODM Data - XM', '文件数据路径'],
			['Invoice_File_URL', 'N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\收样\\3.Sap\\ODM Data - XM\\2.特殊开票', '特殊开票文件路径'],
			['Invoice_File_Name', '特殊开票要求2022.xlsx', '特殊开票文件名称'],
			['SAP登入信息', '内容', '备注'],
			['Login_msg', 'DR-0486-01->601-240', '订单类型-销售组织-分销渠道-销售办事处-销售组'],
			['Hourly Rate', '金额', '备注'],
			["Hourly Rate(PC)", 315, '每年更新'],
			['Hourly Rate(CHM)', 342, '每年更新'],
			['Hourly Rate(PHY)', 342, '每年更新'],
			['成本中心', '编号', '备注'],
			['CS_Selected', 1, '是否默认被选中,1选中，0未选中'],
			['PHY_Selected', 1, '是否默认被选中,1选中，0未选中'],
			['CHM_Selected', 1, '是否默认被选中,1选中，0未选中'],
			['CS_Cost_Center', '48601240', 'CS成本中心'],
			['CHM_Cost_Center', '48601293', 'CHM成本中心'],
			['PHY_Cost_Center', '48601294', 'PHY成本中心'],
			['计划成本', '数值', '备注'],
			['Plan_Cost_Parameter', 0.9, '实际的90%，预留10%利润'],
			['Significant_Digits', 0, '保留几位有效数值'],
			['实验室成本比例', '数值', '备注'],
			['CHM_Cost_Parameter', 0.3, '给到CHM30%'],
			['PHY_Cost_Parameter', 0.3, '给到PHY30%'],
			['DATA A数据填写', '判断依据', '备注'],
			['Data_A_E1', '5010815347;5010427355;5010913488;5010685589;5010829635;5010817524', 'Data A录E1,新添加用;隔开即可'],
			['Data_A_Z2', '5010908478;5010823259', 'Data A录Z2,新添加用;隔开即可'],
			['SAP操作', '内容', '备注'],
			['NVA01_Selected', 1, '是否默认被选中,1选中，0未选中'],
			['NVA02_Selected', 1, '是否默认被选中,1选中，0未选中'],
			['NVF01_Selected', 0, '是否默认被选中,1选中，0未选中'],
			['NVF03_Selected', 0, '是否默认被选中,1选中，0未选中'],
			['DataB_Selected', 1, '是否默认被选中,1选中，0未选中'],
			['Plan_Cost_Selected', 25, '每月超过几号自动选中（不包含）'],
			['Save_Selected', 1, '是否默认被选中,1选中，0未选中'],
			['管理操作', '内容', '备注'],
			['Invoice_No_Selected', 1, '是否默认被选中,1选中，0未选中'],
			['Invoice_Start_Num', 4, 'Invoice的起始数字'],
			['Invoice_Num', 9, 'Invoice的总位数'],
			['Company_Name_Selected', 1, '是否默认被选中,1选中，0未选中'],
			['Order_No_Selected', 0, '是否默认被选中,1选中，0未选中'],
			['Order_Start_Num', 7, 'Order的起始数字'],
			['Order_Num', 9, 'Order的总位数'],
			['Project_No_Selected', 0, '是否默认被选中,1选中，0未选中'],
			['PDF_Name', 'Invoice No + Company Name', 'PDF文件名称默认规则'],
			['PDF_Files_Import_URL', desktopUrl, 'PDF文件导入路径'],
			['PDF_Files_Export_URL', 'N:\\XM Softlines\\1. Project\\3. Finance\\02. WIP', 'PDF文件导出路径'],
			['名称', '编号', '角色'],
			['chen, frank', '6375108', 'CS'],
			['chen, frank', '6375108', 'Sales'],
		]
		config = np.array(configContent)
		df = pd.DataFrame(config)
		df.to_csv('%s/config_sap.csv' % configFileUrl, index=0, header=0, encoding='utf_8_sig')
		self.textBrowser_2.append("配置文件创建成功")
		QMessageBox.information(self, "提示信息",
								"默认配置文件已经创建好，\n如需修改请在用户桌面查找config文件夹中config_sap.csv，\n将相应的文件内容替换成用户需求即可，修改后记得重新导入配置文件。",
								QMessageBox.Yes)

	def exportConfig(self):
		# 重新导出默认配置文件
		reply = QMessageBox.question(self, '信息', '确认是否要创建默认配置文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
		if reply == QMessageBox.Yes:
			MyMainWindow.createConfigContent(self)
		else:
			QMessageBox.information(self, "提示信息", "没有创建默认配置文件，保留原有的配置文件", QMessageBox.Yes)

	def importConfig(self):
		# 重新导入配置文件
		reply = QMessageBox.question(self, '信息', '确认是否要导入配置文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
		if reply == QMessageBox.Yes:
			MyMainWindow.getConfigContent(self)
		else:
			QMessageBox.information(self, "提示信息", "没有重新导入配置文件，将按照原有的配置文件操作", QMessageBox.Yes)

	def getDefaultInformation(self):
		# 默认登录界面信息
		try:
			loginMsgList = configContent['Login_msg'].split('-')
			self.lineEdit_10.setText(loginMsgList[0])
			self.lineEdit_11.setText(loginMsgList[1])
			self.lineEdit_12.setText(loginMsgList[2])
			self.lineEdit_13.setText(loginMsgList[3])
			self.lineEdit_14.setText(loginMsgList[4])
			# 每小时成本
			self.doubleSpinBox_5.setValue(float(format(float(configContent['Hourly Rate(PC)']), '.2f')))
			self.doubleSpinBox_6.setValue(float(format(float(configContent['Hourly Rate(CHM)']), '.2f')))
			self.doubleSpinBox_8.setValue(float(format(float(configContent['Hourly Rate(PHY)']), '.2f')))
			# 成本中心
			self.checkBox_13.setChecked(int(configContent['CS_Selected']))
			self.checkBox_14.setChecked(int(configContent['CHM_Selected']))
			self.checkBox_15.setChecked(int(configContent['PHY_Selected']))
			self.lineEdit_18.setText(configContent['CS_Cost_Center'])
			self.lineEdit_19.setText(configContent['CHM_Cost_Center'])
			self.lineEdit_20.setText(configContent['PHY_Cost_Center'])
			# 计划成本
			self.doubleSpinBox_7.setValue(float(format(float(configContent['Plan_Cost_Parameter']), '.2f')))
			self.spinBox_5.setValue(int(configContent['Significant_Digits']))
			# 实验室分配比例
			self.doubleSpinBox_9.setValue(float(format(float(configContent['CHM_Cost_Parameter']), '.2f')))
			self.doubleSpinBox_10.setValue(float(format(float(configContent['PHY_Cost_Parameter']), '.2f')))
			# DATA A选择
			self.lineEdit_21.setText(configContent['Data_A_E1'])
			self.lineEdit_22.setText(configContent['Data_A_Z2'])
			# SAP操作
			self.checkBox.setChecked(int(configContent['NVA01_Selected']))
			self.checkBox_2.setChecked(int(configContent['NVA02_Selected']))
			self.checkBox_3.setChecked(int(configContent['NVF01_Selected']))
			self.checkBox_4.setChecked(int(configContent['NVF03_Selected']))
			self.checkBox_7.setChecked(int(configContent['DataB_Selected']))
			self.checkBox_6.setChecked(int(configContent['Save_Selected']))
			if int(configContent['Plan_Cost_Selected']) < int(today.split('.')[-1]):
				self.checkBox_8.setChecked(True)
			# admin操作
			self.checkBox_9.setChecked(int(configContent['Invoice_No_Selected']))
			self.spinBox.setValue(int(configContent['Invoice_Start_Num']))
			self.spinBox_2.setValue(int(configContent['Invoice_Num']))
			self.checkBox_10.setChecked(int(configContent['Company_Name_Selected']))
			self.checkBox_12.setChecked(int(configContent['Order_No_Selected']))
			self.spinBox_3.setValue(int(configContent['Order_Start_Num']))
			self.spinBox_4.setValue(int(configContent['Order_Num']))
			self.checkBox_11.setChecked(int(configContent['Project_No_Selected']))
			self.lineEdit_17.setText(configContent['PDF_Name'])
		except Exception as msg:
			self.textBrowser_2.append("错误信息：%s" % msg)
			self.textBrowser_2.append('----------------------------------')
			app.processEvents()
			reply = QMessageBox.question(self, '信息', '错误信息：%s。\n是否要重新创建配置文件' % msg, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
			if reply == QMessageBox.Yes:
				MyMainWindow.createConfigContent(self)
				self.textBrowser.append("创建并导入配置成功")
				self.textBrowser_2.append('----------------------------------')
				app.processEvents()





	def csItem(self):
		self.comboBox_2.clear()
		self.comboBox_2.addItem('')
		nameList = username
		i = 0
		for each in nameList:
			if role[i] == 'CS':
				self.comboBox_2.addItem(each)
			i += 1
			app.processEvents()

	def salesItem(self):
		self.comboBox_3.clear()
		self.comboBox_3.addItem('')
		self.comboBox_3.addItem('')
		nameList = username
		i = 0
		for each in nameList:
			if role[i] == 'Sales':
				self.comboBox_3.addItem(each)
			i += 1
			app.processEvents()

	def showAuthorMessage(self):
		# 关于作者
		QMessageBox.about(self, "关于",
						  "人生苦短，码上行乐。\n\n\n        ----Frank Chen")

	def showVersion(self):
		# 关于作者
		QMessageBox.about(self, "版本",
						  "V 22.01.11\n\n\n 2022-04-26")

	def getAmountVat(self):
		amount = float(self.doubleSpinBox_2.text())
		self.doubleSpinBox_4.setValue(amount*1.06)

	def getGuiData(self):
		guiData = {}
		guiData['sapNo'] = self.lineEdit.text()
		guiData['projectNo'] = self.lineEdit_2.text()
		guiData['materialCode'] = self.comboBox_4.currentText()
		guiData['currencyType'] = self.comboBox.currentText()
		guiData['exchangeRate'] = float(self.doubleSpinBox.text())
		guiData['globalPartnerCode'] = self.lineEdit_3.text()
		guiData['csName'] = self.comboBox_2.currentText()
		guiData['salesName'] = self.comboBox_3.currentText()
		guiData['amount'] = float(self.doubleSpinBox_2.text())
		guiData['cost'] = float(self.doubleSpinBox_3.text())
		guiData['amountVat'] = float(self.doubleSpinBox_4.text())
		guiData['csHourlyRate'] = float(self.doubleSpinBox_5.text())
		guiData['chmHourlyRate'] = float(self.doubleSpinBox_6.text())
		guiData['phyHourlyRate'] = float(self.doubleSpinBox_8.text())
		guiData['longText'] = self.lineEdit_4.text()
		guiData['shortText'] = self.lineEdit_5.text()
		guiData['planCostRate'] = float(self.doubleSpinBox_7.text())
		guiData['significantDigits'] = int(self.spinBox_5.text())
		guiData['chmCostRate'] = float(self.doubleSpinBox_9.text())
		guiData['phyCostRate'] = float(self.doubleSpinBox_10.text())
		guiData['dataAE1'] = self.lineEdit_21.text().split(';')
		guiData['dataAZ2'] = self.lineEdit_22.text().split(';')
		guiData['invoiceStsrtNum'] = int(self.spinBox.text())
		guiData['invoiceBits'] = int(self.spinBox_2.text())
		guiData['orderStsrtNum'] = int(self.spinBox_3.text())
		guiData['orderBits'] = int(self.spinBox_4.text())
		guiData['pdfName'] = self.lineEdit_17.text()
		return guiData
	def getAdminGuiData(self):
		guiAdminData = {}
		guiAdminData['invoiceStsrtNum'] = int(self.spinBox.text())
		guiAdminData['invoiceBits'] = int(self.spinBox_2.text())
		guiAdminData['orderStsrtNum'] = int(self.spinBox_3.text())
		guiAdminData['orderBits'] = int(self.spinBox_4.text())
		guiAdminData['pdfName'] = self.lineEdit_17.text()
		return guiAdminData

	def sapOperate(self):
		logMsg = {}
		logMsg['Remark'] = ''
		logMsg['orderNo'] = ''
		logMsg['Proforma No.'] = ''
		try:
			SapGuiAuto = win32com.client.GetObject("SAPGUI")
			if not type(SapGuiAuto) == win32com.client.CDispatch:
				return

			application = SapGuiAuto.GetScriptingEngine
			if not type(application) == win32com.client.CDispatch:
				SapGuiAuto = None
				return

			connection = application.Children(0)
			if not type(connection) == win32com.client.CDispatch:
				application = None
				SapGuiAuto = None
				return

			session = connection.Children(0)
			if not type(session) == win32com.client.CDispatch:
				connection = None
				application = None
				SapGuiAuto = None
				return


			flag = 1
			guiData= MyMainWindow.getGuiData(self)
			# guiData['sapNo'] = self.lineEdit.text()
			# guiData['projectNo'] = self.lineEdit_2.text()
			# guiData['materialCode'] = self.comboBox_4.currentText()
			# guiData['currencyType'] = self.comboBox.currentText()
			# guiData['exchangeRate'] = float(self.doubleSpinBox.text())
			# guiData['globalPartnerCode'] = self.lineEdit_3.text()
			# guiData['csName'] = self.comboBox_2.currentText()
			if guiData['csName'] != '':
				csCode = configContent[guiData['csName']]
			# guiData['salesName'] = self.comboBox_3.currentText()
			if guiData['salesName'] != '':
				salesCode = configContent[guiData['salesName']]
			orderNo = ''
			proformaNo = ''
			# guiData['amount'] = float(self.doubleSpinBox_2.text())
			# guiData['cost'] = float(self.doubleSpinBox_3.text())
			# guiData['amountVat'] = float(self.doubleSpinBox_4.text())
			# guiData['longText'] = self.lineEdit_4.text()
			# guiData['shortText'] = self.lineEdit_5.text()
			if guiData['sapNo'] == '' or guiData['projectNo'] == '' or guiData['materialCode'] == '' or guiData['currencyType'] == '' or guiData['exchangeRate'] == '' or guiData['globalPartnerCode'] == '' or guiData['csName'] == '' or guiData['amount'] == 0.00 or guiData['amountVat'] == 0.00:
				self.textBrowser.append("有关键信息未填")
				logMsg['Remark'] = '有关键信息未填'
				self.textBrowser.append("'Project No.', 'CS', 'Sales', 'Currency', 'GPC Glo. Par. Code', 'Material Code','SAP No.', 'Amount', 'Amount with VAT', 'Exchange Rate'都是必须填写的")
				self.textBrowser.append('----------------------------------')
				app.processEvents()
				QMessageBox.information(self, "提示信息", "有关键信息未填", QMessageBox.Yes)

			else:
				revenue = guiData['amountVat'] / 1.06
				# plan cost
				# planCost = revenue * guiData['exchangeRate'] * 0.9 - guiData['cost']
				planCost = revenue * guiData['exchangeRate']
				revenueForCny = revenue * guiData['exchangeRate']
				if ('405' in guiData['materialCode']) and (("A2" in guiData['materialCode']) or ("D2" in guiData['materialCode']) or ("D3" in guiData['materialCode'])):
					# DataB-CHM成本
					chmCost = format((revenueForCny - guiData['cost']) * guiData['chmCostRate'] * 0.5, '.2f')
					# DataB-PHY成本
					phyCost = format((revenueForCny - guiData['cost']) * guiData['phyCostRate'] * 0.5, '.2f')
					# Item1000 的revenue
					chmRe = format(revenue * 0.5, '.2f')
					# Item2000 的revenue
					phyRe = format(revenue * 0.5, '.2f')
					# plan cost总算法
					# chmCsCostAccounting = format(planCost * 0.5 * (1 - 0.3  - (1 - guiData['planCostRate'] )) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
					# chmLabCostAccounting = format(planCost * 0.5 * 0.3 / guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])
					# phyCsCostAccounting = format(planCost * 0.5 * (1 - 0.3  - (1 - guiData['planCostRate'] )) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
					# phyLabCostAccounting = format(planCost * 0.5 * 0.3 / guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])

					# plan cost，理论上（revenue-total cost）*0.9*0.5，实际上SFL省略了0.9的计算（金额不大）

					# CS的Item1000-Cost
					chmCsCostAccounting = format((revenueForCny*guiData['planCostRate'] - guiData['cost']) * 0.5 * (1 - guiData['chmCostRate']) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
					# CHM的Item1000-Cost
					chmLabCostAccounting = format((revenueForCny*guiData['planCostRate'] - guiData['cost']) * 0.5 * guiData['chmCostRate'] / guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])
					# CS的Item2000-Cost
					phyCsCostAccounting = format((revenueForCny*guiData['planCostRate'] - guiData['cost']) * 0.5 * (1 - guiData['phyCostRate']) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
					# PHY的Item2000-Cost
					phyLabCostAccounting = format((revenueForCny*guiData['planCostRate'] - guiData['cost']) * 0.5 * guiData['phyCostRate'] / guiData['phyHourlyRate'], '.%sf' % guiData['significantDigits'])
				elif ('441' in guiData['materialCode']) and (("A2" in guiData['materialCode'] or ("D2" in guiData['materialCode']) or ("D3" in guiData['materialCode']))):
					# DataB-CHM成本
					chmCost = format((revenueForCny - guiData['cost']) * guiData['chmCostRate'] * 0.8, '.2f')
					# DataB-PHY成本
					phyCost = format((revenueForCny - guiData['cost']) * guiData['phyCostRate'] * 0.2, '.2f')
					# Item1000 的revenue
					chmRe = format(revenue * 0.8, '.2f')
					# Item2000 的revenue
					phyRe = format(revenue * 0.2, '.2f')
					# plan cost总算法
					# chmCsCostAccounting = format(planCost * 0.8 * (1 - 0.3  - (1 - guiData['planCostRate'] )) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
					# chmLabCostAccounting = format(planCost * 0.8 * 0.3 / guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])
					# phyCsCostAccounting = format(planCost * 0.2 * (1 - 0.3  - (1 - guiData['planCostRate'] )) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
					# phyLabCostAccounting = format(planCost * 0.2 * 0.3 / guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])

					# CS的Item1000-Cost
					chmCsCostAccounting = format((revenueForCny*guiData['planCostRate'] - guiData['cost']) * 0.8 * (1 - guiData['chmCostRate']) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
					# CHM的Item1000-Cost
					chmLabCostAccounting = format((revenueForCny*guiData['planCostRate'] - guiData['cost']) * 0.8 * guiData['chmCostRate'] / guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])
					# CS的Item2000-Cost
					phyCsCostAccounting = format((revenueForCny*guiData['planCostRate'] - guiData['cost']) * 0.2 * (1 - guiData['phyCostRate']) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
					# PHY的Item2000-Cost
					phyLabCostAccounting = format((revenueForCny*guiData['planCostRate'] - guiData['cost']) * 0.2 * guiData['phyCostRate'] / guiData['phyHourlyRate'], '.%sf' % guiData['significantDigits'])
				else:
					chmCost = format((revenueForCny - guiData['cost']) * guiData['chmCostRate'], '.2f')
					phyCost = format((revenueForCny - guiData['cost']) * guiData['phyCostRate'], '.2f')
					chmRe = format(revenue, '.2f')
					phyRe = format(revenue, '.2f')
					# plan cost总算法
					# csCostAccounting = format(planCost * (1 - 0.3  - (1 - guiData['planCostRate'] )) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
					# labCostAccounting = format(planCost * 0.3 / guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])
					if 'T75' in guiData['materialCode']:
						labCostRate = guiData['chmCostRate']
						labHourlyRate = guiData['chmHourlyRate']
					else:
						labCostRate = guiData['phyCostRate']
						labHourlyRate = guiData['phyHourlyRate']

					csCostAccounting = format((revenueForCny*guiData['planCostRate'] - guiData['cost']) * (1 - labCostRate) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
					labCostAccounting = format((revenueForCny*guiData['planCostRate'] - guiData['cost']) * labCostRate / labHourlyRate, '.%sf' % guiData['significantDigits'])

				# 	用未税金额算销售，SAP和ODM会有差别
				# if ('405' in guiData['materialCode']) and (("A2" in guiData['materialCode']) or ("D2" in guiData['materialCode'])):
				# 	chmCost = format((guiData['amount'] * guiData['exchangeRate'] - guiData['cost']) * 0.3 * 0.5, '.2f')
				# 	phyCost = format((guiData['amount'] * guiData['exchangeRate'] - guiData['cost']) * 0.3 * 0.5, '.2f')
				# 	chmRe = format(guiData['amount'] * 0.5, '.2f')
				# 	phyRe = format(guiData['amount'] * 0.5, '.2f')
				#
				# elif ('441' in guiData['materialCode']) and (("A2" in guiData['materialCode'] or ("D2" in guiData['materialCode']))):
				# 	chmCost = format((guiData['amount'] * guiData['exchangeRate'] - guiData['cost']) * 0.3 * 0.8, '.2f')
				# 	phyCost = format((guiData['amount'] * guiData['exchangeRate'] - guiData['cost']) * 0.3 * 0.2, '.2f')
				# 	chmRe = format(guiData['amount'] * 0.8, '.2f')
				# 	phyRe = format(guiData['amount'] * 0.2, '.2f')
				# else:
				# 	chmCost = format((guiData['amount'] * guiData['exchangeRate'] - guiData['cost']) * 0.3, '.2f')
				# 	phyCost = format((guiData['amount'] * guiData['exchangeRate'] - guiData['cost']) * 0.3, '.2f')
				# 	chmRe = guiData['amount']
				# 	phyRe = guiData['amount']



				messageFlag = 1
				if self.checkBox_5.isChecked():
					if guiData['salesName'] == '':
						reply = QMessageBox.question(self, '信息', 'Sales未填，是否继续', QMessageBox.Yes | QMessageBox.No,
													 QMessageBox.Yes)
						if reply == QMessageBox.Yes:
							messageFlag = 1
						else:
							messageFlag = 2
				if guiData['salesName'] != '' or messageFlag == 1:
					self.textBrowser.append("Sap No.:%s" % guiData['sapNo'])
					self.textBrowser.append("Project No.:%s" % guiData['projectNo'])
					self.textBrowser.append("Material Code:%s" % guiData['materialCode'])
					self.textBrowser.append("Global Partner Code:%s" % guiData['globalPartnerCode'])
					self.textBrowser.append("CS Name:%s" % guiData['csName'])
					self.textBrowser.append("Sales Name:%s" % guiData['salesName'])
					self.textBrowser.append("Amount:%s" % guiData['amount'])
					self.textBrowser.append("Cost:%s" % guiData['cost'])
					self.textBrowser.append("Currency Type:%s" % guiData['currencyType'])
					self.textBrowser.append("CHM Cost:%s" % chmCost)
					self.textBrowser.append("PHY Cost:%s" % phyCost)
					self.textBrowser.append("CHM Amount:%s" % chmRe)
					self.textBrowser.append("PHY Amount:%s" % phyRe)
					app.processEvents()
					csCostCenter = self.lineEdit_18.text()
					chmCostCenter = self.lineEdit_19.text()
					phyCostCenter = self.lineEdit_20.text()
					if self.checkBox.isChecked():
						orderType = self.lineEdit_10.text()
						salesOrganization = self.lineEdit_11.text()
						distributionChannels = self.lineEdit_12.text()
						salesOffice = self.lineEdit_13.text()
						salesGroup = self.lineEdit_14.text()
						session.findById("wnd[0]/tbar[0]/okcd").text = "/nva01"
						session.findById("wnd[0]").sendVKey(0)
						session.findById("wnd[0]/usr/ctxtVBAK-AUART").text = orderType
						session.findById("wnd[0]/usr/ctxtVBAK-VKORG").text = salesOrganization
						session.findById("wnd[0]/usr/ctxtVBAK-VTWEG").text = distributionChannels
						session.findById("wnd[0]/usr/ctxtVBAK-VKBUR").text = salesOffice
						session.findById("wnd[0]/usr/ctxtVBAK-VKGRP").text = salesGroup
						session.findById("wnd[0]").sendVKey(0)
						session.findById(
							"wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").text = guiData['sapNo']
						session.findById(
							"wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").caretPosition = 6
						session.findById("wnd[0]").sendVKey(0)
						session.findById("wnd[0]").sendVKey(0)
						session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").text = guiData['projectNo']
						session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBKD-BSTDK").text = today
						session.findById(
							"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/ctxtVBKD-FBUDA").text = today
						session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").setFocus()
						session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").caretPosition = 17
						session.findById("wnd[0]").sendVKey(0)
						session.findById("wnd[1]/tbar[0]/btn[0]").press()
						session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press()
						session.findById(
							"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-WAERK").text = guiData['currencyType']

						session.findById(
							"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-WAERK").setFocus()
						session.findById(
							"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-WAERK").caretPosition = 3
						session.findById("wnd[0]").sendVKey(0)
						try:
							session.findById("wnd[1]").sendVKey(0)
						except:
							pass
						else:
							pass
						if guiData['currencyType'] != "CNY":
							session.findById(
								"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBKD-KURSK").text = guiData['exchangeRate']
							session.findById(
								"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBKD-KURSK").setFocus()
							session.findById(
								"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBKD-KURSK").caretPosition = 8
							session.findById("wnd[0]").sendVKey(0)
						# 会计
						session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\06").select()
						session.findById(
							"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\06/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").text = "*"
						session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\06/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").setFocus()
						session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\06/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").caretPosition = 1
						session.findById("wnd[0]").sendVKey(0)
						# 合作伙伴
						session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09").select()

						# 获取文本名称
						fourName = session.findById(
							"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,4]").text
						fiveName = session.findById(
							"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,5]").text

						# # eNum负责雇员位置，gNum送达方位置
						if fourName == '负责雇员' or fourName == 'Employee respons.':
							eNum = 4
							gNum = 5
						else:
							eNum = 5
							gNum = 4

						# 送达方GPC
						# eNum负责雇员位置，gNum送达方位置

						# if fiveName == '送达方' or fiveName == 'Global Partner':
						# 	gNum = 5
						# else:
						# 	gNum = 4
						session.findById(
							"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,%s]" % gNum).key = "ZG"
						session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,%s]" % gNum).text = guiData['globalPartnerCode']
						session.findById(
							"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,%s]" % gNum).setFocus()
						session.findById(
							"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,%s]" % gNum).caretPosition = 8
						# session.findById("wnd[0]").sendVKey(0)

						# fiveName = session.findById(
						# 	"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,5]").text

						# 负责雇员cs
						# eNum负责雇员位置，gNum送达方位置
						# if fourName == '负责雇员' or fourName == 'Employee respons.':
						# 	eNum = 4
						# else:
						# 	eNum = 5

						session.findById(
							"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,%s]" % eNum).text = csCode
						session.findById("wnd[0]").sendVKey(0)

						# 联系人
						session.findById(
							"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,6]").key = "AP"
						session.findById(
							"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,6]").setFocus()
						session.findById(
							"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,6]").caretPosition = 0
						session.findById("wnd[0]").sendVKey(4)
						session.findById("wnd[1]/tbar[0]/btn[0]").press()
						session.findById("wnd[1]/tbar[0]/btn[0]").press()
						session.findById("wnd[0]").sendVKey(0)
						if guiData['salesName'] != '':
							session.findById(
								"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,7]").key = "VE"
							session.findById(
								"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,7]").text = salesCode
							session.findById(
								"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,7]").setFocus()
							session.findById(
								"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,7]").caretPosition = 4
							session.findById("wnd[0]").sendVKey(0)
						# 文本
						session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10").select()
						session.findById(
							"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = guiData['shortText']
						session.findById(
							"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes(
							11, 11)
						session.findById(
							"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cmbLV70T-SPRAS").key = "EN"
						session.findById(
							"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cmbLV70T-SPRAS").setFocus()
						session.findById("wnd[0]").sendVKey(0)
						# DATA A
						session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13").select()
						if 'D2' in guiData['materialCode'] or 'D3' in guiData['materialCode']:
							if guiData['sapNo'] in guiData['dataAE1']:
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR1").key = "E1"
							else:
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR1").key = "Z0"
						elif guiData['sapNo'] in guiData['dataAZ2']:
							session.findById(
								"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR1").key = "Z2"
						else:
							session.findById(
								"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR1").key = "00"
						# DATA B
						session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14").select()
						session.findById(
							"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/ctxtVBAK-ZZAUART").text = "WO"
						session.findById(
							"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/ctxtVBAK-ZZUNLIMITLIAB").text = "N"
						session.findById(
							"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/ctxtZAUFTD-VORAUS_AUFENDE").text = oneWeekday
						if revenueForCny >= 35000:
							session.findById(
								"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/txtZAUFTD-AUFTRAGSWERT").text = format(revenueForCny, '.2f')
						# 是否要添加cost 
						if self.checkBox_7.isChecked():
							if 'A2' in guiData['materialCode'] or 'D2' in guiData['materialCode'] or 'D3' in guiData['materialCode']:
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,0]").text = chmCostCenter
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,1]").text = phyCostCenter
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/ctxtTABD-KOSTL[0,0]").text = chmCostCenter
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/ctxtTABD-KOSTL[0,1]").text = phyCostCenter
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,0]").text = chmCost
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,1]").text = phyCost
							elif 'T20' in guiData['materialCode']:
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,0]").text = phyCostCenter
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/ctxtTABD-KOSTL[0,0]").text = phyCostCenter
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,0]").text = phyCost
							else:
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,0]").text = chmCostCenter
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/ctxtTABD-KOSTL[0,0]").text = chmCostCenter
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,0]").text = chmCost

						session.findById("wnd[0]").sendVKey(0)
						session.findById(
							"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,1]").setFocus()
						session.findById(
							"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,1]").caretPosition = 8

						if self.checkBox_2.isChecked() or self.checkBox_6.isChecked():
							try:
								session.findById("wnd[0]/tbar[0]/btn[3]").press()
								session.findById("wnd[0]/tbar[0]/btn[3]").press()
								session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
							except:
								try:
									session.findById("wnd[0]/tbar[0]/btn[3]").press()
									session.findById("wnd[0]/tbar[0]/btn[3]").press()
									session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
									session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
								except:
									self.textBrowser.append("跳过保存")
								else:
									pass
							else:
								pass
							# session.findById("wnd[0]/tbar[0]/btn[3]").press()
							# session.findById("wnd[0]/tbar[0]/btn[3]").press()
							# session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()

					if self.checkBox_2.isChecked():
							# try:
							# 	session.findById("wnd[0]/tbar[0]/btn[3]").press()
							# 	session.findById("wnd[0]/tbar[0]/btn[3]").press()
							# 	session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
							# except:
							# 	try:
							# 		session.findById("wnd[0]/tbar[0]/btn[3]").press()
							# 		session.findById("wnd[0]/tbar[0]/btn[3]").press()
							# 		session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
							# 		session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
							# 	except:
							# 		self.textBrowser.append("已保存")
							# 	else:
							# 		pass
							# else:
							# 	pass
							session.findById("wnd[0]/tbar[0]/okcd").text = "/NVA02"
							session.findById("wnd[0]").sendVKey(0)
							orderNo = session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text
							self.textBrowser.append("Order No.:%s" % orderNo)
							app.processEvents()
							logMsg['orderNo'] = orderNo
							session.findById("wnd[0]").sendVKey(0)
							if 'A2' in guiData['materialCode']:
								if '405' in guiData['materialCode']:
									session.findById(
										"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").text = "T75-405-00"
									session.findById(
										"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,1]").text = "T20-405-00"
								else:
									session.findById(
										"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").text = "T75-441-00"
									session.findById(
										"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,1]").text = "T20-441-00"
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/txtVBAP-ZMENG[2,0]").text = "1"
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/txtVBAP-ZMENG[2,1]").text = "1"
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtVBAP-ZIEME[3,0]").text = "pu"
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtVBAP-ZIEME[3,1]").text = "pu"
								session.findById("wnd[0]").sendVKey(0)

								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/txtVBAP-ZMENG[2,1]").setFocus()
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/txtVBAP-ZMENG[2,1]").caretPosition = 16
								session.findById("wnd[0]").sendVKey(2)
								session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06").select()
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text = phyRe
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").setFocus()
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").caretPosition = 16
								session.findById("wnd[0]").sendVKey(0)
								sapAmountVatStr = session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text
								sapAmountVat = float(sapAmountVatStr.replace(',', ''))

								session.findById("wnd[0]/tbar[0]/btn[3]").press()

								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").setFocus()
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").caretPosition = 10
								session.findById("wnd[0]").sendVKey(2)
								session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06").select()
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text = chmRe
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").setFocus()
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").caretPosition = 16
								session.findById("wnd[0]").sendVKey(0)
								sapAmountVatStr = session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text

								sapAmountVat += float(sapAmountVatStr.replace(',', ''))
								sapAmountVat = format(sapAmountVat,'.2f')
								sapAmountVat = re.sub(r"(\d)(?=(\d\d\d)+(?!\d))", r"\1,", sapAmountVat)


								if self.checkBox_8.isChecked() or revenueForCny >= 35000:

									session.findById("wnd[0]/tbar[0]/btn[3]").press()
									if revenueForCny >= 1000:
										# 这个是Item2000的
										session.findById(
											"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,1]").setFocus()
										session.findById(
											"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,1]").caretPosition = 10
										session.findById("wnd[0]/mbar/menu[3]/menu[7]").select()
										session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
										session.findById("wnd[1]/tbar[0]/btn[0]").press()
										# cs
										if self.checkBox_13.isChecked():
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,0]").text = "E"
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,0]").text = csCostCenter
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,0]").text = "T01AST"
											# 录金额
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").text = round(
												float(phyCsCostAccounting), 0)
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").setFocus()
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").caretPosition = 20
											session.findById("wnd[0]").sendVKey(0)
										# phy
										if self.checkBox_15.isChecked():
											session.findById("wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,1]").text = "E"
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,1]").text = phyCostCenter
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,1]").text = "T01AST"
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").text = round(
												float(phyLabCostAccounting), 0)
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").setFocus()
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").caretPosition = 20
										session.findById("wnd[0]").sendVKey(0)



										# session.findById("wnd[0]").sendVKey(0)
										session.findById("wnd[0]/tbar[0]/btn[3]").press()

										# session.findById("wnd[0]/tbar[0]/btn[11]").press()
										session.findById("wnd[0]/tbar[0]/btn[3]").press()
										session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()


										# Items1000的plan cost
										session.findById(
											"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").setFocus()
										session.findById(
											"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").caretPosition = 10
										session.findById("wnd[0]/mbar/menu[3]/menu[7]").select()
										session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
										session.findById("wnd[1]/tbar[0]/btn[0]").press()
										# cs
										if self.checkBox_13.isChecked():
											session.findById("wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,0]").text = "E"
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,0]").text = csCostCenter
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,0]").text = "T01AST"
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").text = round(
												float(chmCsCostAccounting), 0)
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").setFocus()
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").caretPosition = 19
										# 	chm
										if self.checkBox_14.isChecked():
											session.findById("wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,1]").text = "E"
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,1]").text = chmCostCenter
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,1]").text = "T01AST"
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").text = round(
												float(chmLabCostAccounting), 0)
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").setFocus()
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").caretPosition = 20
										session.findById("wnd[0]").sendVKey(0)
										if guiData['cost'] > 0:
											if self.checkBox_14.isChecked():
												n = 2
											else:
												n = 1
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,%s]" % n).text = "E"
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,%s]" % n).text = csCostCenter
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,%s]" % n).text = "FREMDL"
											session.findById("wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).text = format(guiData['cost'] / 1.06, '.2f')
											session.findById("wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).setFocus()
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).caretPosition = 20
											session.findById("wnd[0]").sendVKey(0)

										# session.findById("wnd[0]/tbar[0]/btn[11]").press()
										session.findById("wnd[0]/tbar[0]/btn[3]").press()
										session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
										# session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
							else:
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").text = guiData['materialCode']
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/txtVBAP-ZMENG[2,0]").text = "1"
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtVBAP-ZIEME[3,0]").text = "pu"
								session.findById("wnd[0]").sendVKey(0)
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").setFocus()
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").caretPosition = 10
								session.findById("wnd[0]").sendVKey(2)
								session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06").select()
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text = format(revenue, '.2f')
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").setFocus()
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").caretPosition = 16
								session.findById("wnd[0]").sendVKey(0)
								sapAmountVat = session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text


								if self.checkBox_8.isChecked() or revenueForCny >= 35000:
									session.findById("wnd[0]/tbar[0]/btn[3]").press()
									if revenueForCny >= 1000:
										session.findById(
											"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").setFocus()
										session.findById(
											"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").caretPosition = 10
										session.findById("wnd[0]/mbar/menu[3]/menu[7]").select()
										session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
										session.findById("wnd[1]/tbar[0]/btn[0]").press()
										if self.checkBox_13.isChecked():
											session.findById("wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,0]").text = "E"
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,0]").text = csCostCenter
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,0]").text = "T01AST"
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").text = round(
												float(csCostAccounting), 0)
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").setFocus()
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").caretPosition = 19

										if self.checkBox_14.isChecked() or self.checkBox_15.isChecked():
											session.findById("wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,1]").text = "E"
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,1]").text = "T01AST"
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").text = round(
												float(labCostAccounting), 0)
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").setFocus()
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").caretPosition = 20

										if 'T75' in guiData['materialCode']:
											if self.checkBox_14.isChecked():
												session.findById(
													"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,1]").text = chmCostCenter
										else:
											if self.checkBox_15.isChecked():
												session.findById(
													"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,1]").text = phyCostCenter

										if guiData['cost'] > 0:
											if self.checkBox_14.isChecked() or self.checkBox_15.isChecked():
												n = 2
											else:
												n = 1
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,%s]" % n).text = "E"
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,%s]" % n).text = csCostCenter
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,%s]" % n).text = "FREMDL"
											session.findById("wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).text = format(guiData['cost'] / 1.06, '.2f')
											session.findById("wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).setFocus()
											session.findById(
												"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).caretPosition = 20
											session.findById("wnd[0]").sendVKey(0)
										# 直接保存
										# session.findById("wnd[0]/tbar[0]/btn[11]").press()
										session.findById("wnd[0]/tbar[0]/btn[3]").press()
										session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()


							if guiData['longText'] != '':
								if self.checkBox_8.isChecked() or revenueForCny >= 35000:
									session.findById(
										"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").setFocus()
									session.findById(
										"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").caretPosition = 10
									session.findById("wnd[0]").sendVKey(2)

								session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09").select()
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = guiData['longText']
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes(
									4, 4)
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cmbLV70T-SPRAS").key = "EN"
								session.findById(
									"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cmbLV70T-SPRAS").setFocus()
								session.findById("wnd[0]").sendVKey(0)

							if self.checkBox_8.isChecked() or revenueForCny >= 35000:
								pass
							else:
								session.findById("wnd[0]/tbar[0]/btn[3]").press()
							amountVatStr = re.sub(r"(\d)(?=(\d\d\d)+(?!\d))", r"\1,", format(guiData['amountVat'],'.2f'))
							self.textBrowser.append("Sap Amount Vat:%s" % sapAmountVat)
							self.textBrowser.append("Amount Vat:%s" % amountVatStr)
							app.processEvents()
							# sapAmountVat在A2是数字，其它为字符串
							if sapAmountVat.strip() != amountVatStr:
								flag = 2
								reply = QMessageBox.question(self, '信息', 'SAP数据与ODM不一致，请确认并修改后再继续！！！',
															 QMessageBox.Yes | QMessageBox.No,
															 QMessageBox.Yes)
								logMsg['Remark'] = 'SAP数据与ODM不一致，请确认并修改后再继续！！！'
								if reply == QMessageBox.Yes:
									flag = 1
							if (self.checkBox_3.isChecked() or self.checkBox_6.isChecked()) and flag == 1:
								try:
									session.findById("wnd[0]/tbar[0]/btn[3]").press()
									session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
									session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
								except:
									try:
										session.findById("wnd[0]/tbar[0]/btn[3]").press()
										session.findById("wnd[0]/tbar[0]/btn[3]").press()
										session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
										session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
									except:
										self.textBrowser.append("跳过保存")
									else:
										pass
								else:
									pass

					if self.checkBox_3.isChecked() and flag == 1:
						try:
							session.findById("wnd[0]/tbar[0]/btn[3]").press()
							session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
							session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
						except:
							try:
								session.findById("wnd[0]/tbar[0]/btn[3]").press()
								session.findById("wnd[0]/tbar[0]/btn[3]").press()
								session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
								session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
							except:
								self.textBrowser.append("已保存")
							else:
								pass
						else:
							pass
						session.findById("wnd[0]/tbar[0]/okcd").text = "/nvf01"
						session.findById("wnd[0]").sendVKey(0)
						session.findById("wnd[0]/tbar[0]/btn[11]").press()

					if self.checkBox_4.isChecked():
						session.findById("wnd[0]/tbar[0]/okcd").text = "/nvf03"
						session.findById("wnd[0]").sendVKey(0)
						proformaNo = session.findById("wnd[0]/usr/ctxtVBRK-VBELN").text
						logMsg['Proforma No.'] = proformaNo
						self.textBrowser.append("Proforma No.:%s" % proformaNo)
						app.processEvents()
						session.findById("wnd[0]/mbar/menu[0]/menu[11]").select()
						session.findById("wnd[1]/tbar[0]/btn[37]").press()

					self.textBrowser.append('SAP操作已完成')
					self.textBrowser.append('----------------------------------')
					app.processEvents()
					if self.checkBox_5.isChecked():
						QMessageBox.information(self, "提示信息", "SAP操作已完成", QMessageBox.Yes)

			return logMsg

		except Exception as msg:
			guiData = MyMainWindow.getGuiData(self)
			self.textBrowser.append('这单%s的数据或者SAP有问题' % guiData['projectNo'])
			self.textBrowser.append('错误信息：%s' % msg)
			logMsg['Remark'] += '错误信息：%s' % msg
			self.textBrowser.append('----------------------------------')
			QMessageBox.information(self, "提示信息", '这份%s的ODM获取数据有问题' % guiData['projectNo'], QMessageBox.Yes)
			return logMsg
			# print(sys.exc_info()[0])

		finally:
			session = None
			connection = None
			application = None
			SapGuiAuto = None

	def getFile(self):
		selectBatchFile = QFileDialog.getOpenFileName(self, '选择ODM导出文件', '%s\\%s' % (configContent['SAP_Date_URL'], today), 'files(*.docx;*.xls*;*.csv)')
		fileUrl = selectBatchFile[0]
		return fileUrl

	def getFileUrl(self):
		fileUrl = MyMainWindow.getFile(self)
		if fileUrl:
			self.lineEdit_6.setText(fileUrl)
			app.processEvents()
		else:
			self.textBrowser.append("请重新选择ODM文件")
			QMessageBox.information(self, "提示信息", "请重新选择ODM文件", QMessageBox.Yes)

	def getODMDataFileUrl(self):
		fileUrl = MyMainWindow.getFile(self)
		if fileUrl:
			self.lineEdit_7.setText(fileUrl)
			app.processEvents()
		else:
			self.textBrowser_2.append("请重新选择ODM文件")
			QMessageBox.information(self, "提示信息", "请重新选择ODM文件", QMessageBox.Yes)

	def getCombineFileUrl(self):
		fileUrl = MyMainWindow.getFile(self)
		if fileUrl:
			self.lineEdit_8.setText(fileUrl)
			app.processEvents()
		else:
			self.textBrowser_2.append("请重新选择ODM文件")
			QMessageBox.information(self, "提示信息", "请重新选择ODM文件", QMessageBox.Yes)

	def getLogFileUrl(self):
		fileUrl = MyMainWindow.getFile(self)
		if fileUrl:
			self.lineEdit_9.setText(fileUrl)
			app.processEvents()
		else:
			self.textBrowser_2.append("请重新选择ODM文件")
			QMessageBox.information(self, "提示信息", "请重新选择ODM文件", QMessageBox.Yes)

	def odmDataToSap(self):
		try:
			fileUrl = self.lineEdit_6.text()
			(filepath, filename) = os.path.split(fileUrl)
			if fileUrl:
				# 下拉框默认选择0
				self.comboBox.setCurrentIndex(0)
				self.comboBox_2.setCurrentIndex(0)
				self.comboBox_3.setCurrentIndex(0)
				self.comboBox_4.setCurrentIndex(0)
				# log文件
				logFileUrl = '%s/log' % filepath
				MyMainWindow.createFolder(self, logFileUrl)
				csvFileType = 'csv'
				logFileName = 'log'
				logDataPath = MyMainWindow.getFileName(self, logFileUrl, logFileName, csvFileType)

				# 获取最终ODM数据
				newData = Get_Data()
				newData.getFileData(fileUrl)
				deleteList = {'Amount': 0}
				newData.deleteTheRows(deleteList)
				headList = newData.getHeaderData()
				# 去除Amount with VAT中数值为空的数据，因为数据sales为空
				newData.fileData = newData.fileData[newData.fileData['Amount with VAT'].notnull()]
				newData.fileData = newData.fileData.reset_index(drop=True)

				if ("PHY Material Code" in headList) and ("CHM Material Code" in headList):
					fillNanColumnKey = {'Material Code': ["PHY Material Code", "CHM Material Code"]}
					newData.fillNanColumn(fillNanColumnKey)
				getFileDataListKey = ['Project No.', 'CS', 'Sales', 'Currency', 'GPC Glo. Par. Code', 'Material Code',
									  'SAP No.', 'Amount', 'Amount with VAT', 'Exchange Rate', 'Total Cost']

				combineKeyFieldsList = ['GPC Glo. Par. Code', 'SAP No.', 'Amount', 'Amount with VAT', 'Total Cost']

				if 'Text' in headList:
					getFileDataListKey.append('Text')
					combineKeyFieldsList.append('Text')
				if 'Long Text' in headList:
					getFileDataListKey.append('Long Text')
					combineKeyFieldsList.append('Long Text')
				# log文件
				combinekeyFields = self.lineEdit_15.text()
				combineKeyFieldsList += combinekeyFields.split(';')
				combineKeyFieldsList.append('Project No.')
				logFile = newData.fileData[combineKeyFieldsList]
				logFile['Order No.'] = ''
				logFile['Remark'] = ''
				logFile['Proforma No.'] = ''
				logFile['Update Time'] = '未开Order'
				if 'Text' not in headList:
					logFile['Text'] = ''
				if 'Long Text' not in headList:
					logFile['Long Text'] = ''

				fileDataList = newData.getFileDataList(getFileDataListKey)
				headerData = newData.getHeaderData()
				n = 0
				for n in range(len(fileDataList['Amount'])):
					if fileDataList['Material Code'][n] == '':
						QMessageBox.information(self, "提示信息", "无Material Code，请检查", QMessageBox.Yes)
						break
					else:
						materialCode = fileDataList['Material Code'][n]
					self.lineEdit_2.setText(fileDataList['Project No.'][n])
					self.lineEdit_3.setText(str(int(fileDataList['GPC Glo. Par. Code'][n])))
					self.textBrowser.append("No.:%s" % (n + 1))
					# if pd.isnull(fileDataList['SAP No.'][n]):
					# # if math.isnan(fileDataList['SAP No.'][n]):
					# 	self.textBrowser.append("没有SAP No.")
					# 	self.lineEdit.setText('')
					# else:
					# 	self.lineEdit.setText(str(int(fileDataList['SAP No.'][n])))
					try:
						self.lineEdit.setText(str(int(fileDataList['SAP No.'][n])))
					except:
						self.lineEdit.setText('')
					else:
						pass
					# materialCodeList = ['', 'T75-441-A2', 'T75-405-A2', 'T20-441-00', 'T20-405-00', 'T75-441-00', 'T75-405-00', 'T75-441-D2', 'T75-405-D2', 'S11-441-10', 'S11-405-10']
					# self.comboBox_4.setCurrentIndex(username.index(materialCode))
					app.processEvents()
					self.comboBox_4.setItemText(int(0), materialCode)

					if fileDataList['CS'][n] in configContent:
						# self.comboBox_2.setCurrentIndex(username.index(fileDataList['CS'][n])+1)
						self.comboBox_2.setItemText(int(0), fileDataList['CS'][n])
					else:
						# self.comboBox_2.setCurrentIndex(0)
						self.comboBox_2.setItemText(int(0), '')
					if fileDataList['Sales'][n] in configContent:
						# self.comboBox_3.setCurrentIndex(username.index(fileDataList['Sales'][n]) + 1)
						self.comboBox_3.setItemText(int(0), fileDataList['Sales'][n])
					else:
						# self.comboBox_3.setCurrentIndex(0)
						self.comboBox_3.setItemText(int(0), '')
					self.comboBox.setItemText(int(0), fileDataList['Currency'][n])
					self.doubleSpinBox_2.setValue(fileDataList['Amount'][n])
					self.doubleSpinBox_4.setValue(fileDataList['Amount with VAT'][n])
					self.doubleSpinBox_3.setValue(fileDataList['Total Cost'][n])
					self.doubleSpinBox.setValue(fileDataList['Exchange Rate'][n])
					if 'Text' in headList:
						try:
							self.lineEdit_5.setText(fileDataList['Text'][n])
						except:
							self.lineEdit_5.setText('Testing Fee')
						else:
							pass
					else:
						self.lineEdit_5.setText('Testing Fee')
					if 'Long Text' in headList:
						try:
							self.lineEdit_4.setText(fileDataList['Long Text'][n])
						except:
							pass
						else:
							pass
					app.processEvents()
					logMsg = MyMainWindow.sapOperate(self)
					# 写log
					logIndex = logFile[(logFile['Project No.'] == fileDataList['Project No.'][n])].index.tolist()[0]
					logFile.loc[logIndex, 'Order No.'] = logMsg['orderNo']
					logFile.loc[logIndex, 'Remark'] = logMsg['Remark']
					logFile.loc[logIndex, 'Proforma No.'] = logMsg['Proforma No.']
					nowDate = datetime.datetime.today()
					logFile.loc[logIndex, 'Update Time'] = nowDate
					logDataFile = logFile.to_csv('%s' % logDataPath, encoding='utf_8_sig')
					self.lineEdit_9.setText(logDataPath)
					if n < len(fileDataList['Amount'])-1:
						if self.checkBox_5.isChecked():
							reply = QMessageBox.question(self, '信息', '是否继续填写下一个Order', QMessageBox.Yes | QMessageBox.No,
														 QMessageBox.Yes)
							if reply == QMessageBox.Yes:
								continue
							else:
								break
					else:
						os.startfile(logFileUrl)
						os.startfile(logDataPath)
						self.textBrowser.append("ODM数据已全部填写完成")
						self.textBrowser.append("log数据:%s" % logDataPath)
						self.textBrowser.append('----------------------------------')
						QMessageBox.information(self, "提示信息", "ODM数据已全部填写完成", QMessageBox.Yes)
			else:
				self.textBrowser.append("请重新选择ODM文件")
				QMessageBox.information(self, "提示信息", "请重新选择ODM文件", QMessageBox.Yes)
		except Exception as msg:
			fileData = self.lineEdit_6.text()
			self.textBrowser.append('这份%s的ODM获取数据有问题' % fileData)
			self.textBrowser.append('错误信息：%s' % msg)
			self.textBrowser.append('----------------------------------')
			QMessageBox.information(self, "提示信息", '这份%s的ODM获取数据有问题' % fileData, QMessageBox.Yes)

	def getFileName(self, fileUrl, fileName, fileType):
		nowTime = time.strftime('%Y-%m-%d %H.%M.%S')
		fileName = fileUrl + '/' + nowTime + ' - ' + fileName + '.' + fileType
		return fileName

	def createFolder(self, url):
		isExists = os.path.exists(url)
		if not isExists:
			os.makedirs(url)

	def lineEditChange(self, url):
		combineKey = self.lineEdit_15.text()
		self.lineEdit_16.setText(combineKey)

	def getInvoiceMsg(self):
		try:
			# 特殊开票文件
			global specialInvoiceMsg
			global invoiceName
			global orderMode
			global invoiceRemarks
			global invoiceField
			global invoiceGroup

			# invoiceFile = pd.read_csv(r'%s/%s' % (configFileUrl, configContent['Invoice_File_Name']), encoding="utf8")
			invoiceFile = pd.read_excel('%s/%s' % (configContent['Invoice_File_URL'], configContent['Invoice_File_Name']))
			invoiceName = list(invoiceFile['Invoice name'])
			orderMode = list(invoiceFile['开order方式'])
			invoiceRemarks = list(invoiceFile['开票要求(特殊的外币开票请参考近期的invoice)'])
			invoiceField = list(invoiceFile['字段'])
			invoiceGroup = list(invoiceFile['组别'])
			orderModeList = list(set(orderMode))
			specialInvoiceMsg = {}
			for each in orderModeList:
				specialInvoiceMsg[each] = {}
				# 保留包含关键字的行
				seachInvoiceFile = pd.DataFrame(invoiceFile[invoiceFile['开order方式'] == each])
				# 嵌套字典
				# invoiceName = {}
				# orderMode = {}
				# invoiceRemarks = {}
				# invoiceField = {}
				# invoiceGroup = {}
				embeddedDict = {}
				embeddedDict['Invoice name'] = list(seachInvoiceFile['Invoice name'])
				embeddedDict['Order Mode'] = list(seachInvoiceFile['开order方式'])
				embeddedDict['Invoice Remarks'] = list(seachInvoiceFile['开票要求(特殊的外币开票请参考近期的invoice)'])
				embeddedDict['Invoice Field'] = list(seachInvoiceFile['字段'])
				embeddedDict['Invoice Group'] = list(seachInvoiceFile['组别'])
				specialInvoiceMsg[each] = embeddedDict

			self.textBrowser_2.append('特殊开票信息获取成功')
			self.textBrowser_2.append('----------------------------------')
		except Exception as msg:
			self.textBrowser_2.append('错误信息：%s' % msg)
			self.textBrowser_2.append('----------------------------------')

	# 分开ODM数据
	def splitOdmData(self):
		try:
			fileUrl = self.lineEdit_7.text()
			(filepath, filename) = os.path.split(fileUrl)
			if fileUrl:
				newData = Get_Data()
				newData.getFileData(fileUrl)
				generalData = newData.fileData[~newData.fileData["Invoices' name (Chinese)"].isin(invoiceName)]
				# 新文件地址
				newFolderUrl = '%s/%s' % (filepath, today)
				newFolder = File_Opetate()
				newFolder.createFolder(newFolderUrl)
				csvFileType = 'csv'
				if generalData.empty:
					pass
				else:
					invoiceFileName = '1.正常合并'
					invoiceFilePath = newFolder.getFileName(newFolderUrl, invoiceFileName, csvFileType)
					generalFile = generalData.to_csv('%s' % invoiceFilePath, encoding='utf_8_sig')
				specialData = newData.fileData[newData.fileData["Invoices' name (Chinese)"].isin(invoiceName)]
				fileNum = 2
				for each in specialInvoiceMsg:
					eachSpecialData = specialData[specialData["Invoices' name (Chinese)"].isin(specialInvoiceMsg[each]['Invoice name'])]
					if eachSpecialData.empty:
						pass
					else:
						invoiceFileName = str(fileNum) + '.' + each
						invoiceFilePath = newFolder.getFileName(newFolderUrl, invoiceFileName, csvFileType)
						specialFile = eachSpecialData.to_csv('%s' % invoiceFilePath, encoding='utf_8_sig')
						fileNum += 1
				os.startfile(newFolderUrl)
				self.textBrowser_2.append('处理好特殊数据')
				self.textBrowser_2.append('----------------------------------')
			else:
				self.textBrowser_2.append('请重新选择ODM文件')
				self.textBrowser_2.append('----------------------------------')
		except Exception as msg:
			self.textBrowser_2.append('错误信息：%s' % msg)
			self.textBrowser_2.append('----------------------------------')


	# 数据透视并合并
	def odmCombineData(self):
		try:
			fileUrl = self.lineEdit_7.text()
			(filepath, filename) = os.path.split(fileUrl)
			if fileUrl:
				newData = Get_Data()
				newData.getFileData(fileUrl)
				# 删除Amount为0的数据
				deleteRowList = {'Amount': 0}
				newData.deleteTheRows(deleteRowList)
				newData.fileData.sort_values(by=["Invoices' name (Chinese)",'CS', 'Sales', 'Currency', 'Material Code', 'Buyer(GPC)', 'Month'], axis=0, ascending=[True, True, True, True, True, True, True], inplace=True)
				# 只保留Order No为空的数据
				newData.fileData = newData.fileData[newData.fileData[['SAP Order No.']].isnull().T.any()]
				# material code将空值填上
				headList = newData.getHeaderData()
				if ("PHY Material Code" in headList) and ("CHM Material Code" in headList):
					fillNanColumnKey = {'Material Code': ["PHY Material Code", "CHM Material Code"]}
					newData.fillNanColumn(fillNanColumnKey)
				# 将联系人空值填上
				newData.fileData['Client Contact Name'].fillna("******", inplace=True)
				# 保存原始数据
				fileUrl = '%s/%s' % (filepath, today)
				MyMainWindow.createFolder(self, fileUrl)
				csvFileType = 'csv'
				odmFileName = '1.ODM Raw Data'
				odmDataPath = MyMainWindow.getFileName(self, fileUrl, odmFileName, csvFileType)
				odmDataFile = newData.fileData.to_csv('%s' % (odmDataPath), encoding='utf_8_sig')
				# 数据透视并保存
				combinekeyFields = self.lineEdit_15.text()
				combineKeyFieldsList = combinekeyFields.split(';')
				pivotTableKey = combineKeyFieldsList
				# pivotTableKey = ['CS', 'Sales', 'Currency', 'Material Code', "Invoices' name (Chinese)", 'Buyer(GPC)', 'Month', 'Exchange Rate']
				valusKey = ['Amount', 'Amount with VAT', 'Total Cost', 'Revenue\n(RMB)']
				pivotTable = newData.pivotTable(pivotTableKey, valusKey)
				combineFileName = '2.Combine'
				combineFileNamePath = MyMainWindow.getFileName(self, fileUrl, combineFileName, csvFileType)
				combineFile = pivotTable.to_csv('%s' % (combineFileNamePath), encoding='utf_8_sig')
				# 读取数据透视数据
				combineData = Get_Data()
				combineData = combineData.getFileData(combineFileNamePath)
				# 删除列
				deleteColumnList = ['Amount', 'Amount with VAT', 'Total Cost', 'Revenue\n(RMB)']
				newData = newData.deleteTheColumn(deleteColumnList)
				# merge数据，combine和原始数据
				onData = combineKeyFieldsList
				# onData = ['CS', 'Sales', 'Currency', 'Material Code', "Invoices' name (Chinese)", 'Buyer(GPC)', 'Month', 'Exchange Rate']
				mergeData = pd.merge(combineData, newData, on=onData, how='right')
				mergeDataName = '3.Merge to Project'
				mergeFileNamePath = MyMainWindow.getFileName(self, fileUrl, mergeDataName, csvFileType)
				mergeFile = mergeData.to_csv('%s' % (mergeFileNamePath), encoding='utf_8_sig')
				self.lineEdit_8.setText(mergeFileNamePath)
				# merge数据去重得到最终数据
				mergeData.drop_duplicates(subset=pivotTableKey, keep='first', inplace=True)
				finalDataName = '4.final'
				finalFileNamePath = MyMainWindow.getFileName(self, fileUrl, finalDataName, csvFileType)
				ascendingList = [True] * len(combineKeyFieldsList)
				mergeData.sort_values(by=combineKeyFieldsList, axis=0, ascending=ascendingList, inplace=True)
				# mergeData.sort_values(by=["Invoices' name (Chinese)", 'CS', 'Sales', 'Currency', 'Material Code', 'Buyer(GPC)', 'Month', 'Exchange Rate'], axis=0, ascending=[True, True, True, True, True, True, True, True], inplace=True)
				finalFile = mergeData.to_csv('%s' % (finalFileNamePath), encoding='utf_8_sig')
				self.textBrowser_2.append('ODM原始数据：%s' % odmDataPath)
				self.textBrowser_2.append('数据透视数据：%s' % combineFileNamePath)
				self.textBrowser_2.append('添加Project No.的数据：%s' % mergeFileNamePath)
				self.textBrowser_2.append('最终的SAP应用数据：%s' % finalFileNamePath)
				self.lineEdit_6.setText(finalFileNamePath)
				self.textBrowser_2.append('ODM数据已处理完成')
				self.textBrowser_2.append('----------------------------------')
				app.processEvents()
				os.startfile(fileUrl)
				os.startfile(finalFileNamePath)
			else:
				self.textBrowser_2.append('请重新选择ODM文件')
				self.textBrowser_2.append('----------------------------------')
		except Exception as msg:
			fileData = self.lineEdit_7.text()
			self.textBrowser_2.append('这份%s的ODM获取数据有问题' % fileData)
			self.textBrowser_2.append('错误信息：%s' % msg)
			self.textBrowser_2.append('----------------------------------')
			app.processEvents()
			# QMessageBox.information(self, "提示信息", '这份%s的ODM获取数据有问题' % fileData, QMessageBox.Yes)

	# 找到project对应的order
	def orderMergeProject(self):
		try:
			combineFileUrl = self.lineEdit_8.text()
			(combineFilepath, combineFilename) = os.path.split(combineFileUrl)
			logFileUrl = self.lineEdit_9.text()
			(logFilepath, logFilename) = os.path.split(logFileUrl)
			if combineFileUrl and logFileUrl:
				csvFileType = 'csv'
				fileUrl = combineFilepath
				combineFile = Get_Data()
				# combineFile.getMergeFileData(combineFileUrl)
				combineFile.getFileData(combineFileUrl)
				logFile = Get_Data()
				# logFile.getMergeFileData(logFileUrl)
				logFile.getFileData(logFileUrl)
				# # 删除列，Project No.保留以便更好的溯源数据
				# deleteColumnList = ['Project No.']
				# logFile = logFile.deleteTheColumn(deleteColumnList)
				# merge数据，combine和原始数据
				mergekeyFields = self.lineEdit_16.text()
				mergekeyFieldsList = mergekeyFields.split(';')
				# 多个字段合并为一列，作为id用于匹配
				# combineKey = ''
				# logKey = ''
				# for each in mergekeyFieldsList:
				# 	combineKey += "+ combineFile.fileData['%s']" % each
				# 	logKey += "+ logFile['%s']" % each
				# combineFile.fileData['ID'] = combineKey
				# logFile['ID'] = logKey
				# combineFile.fileData['ID'] = combineFile.fileData['Amount with VAT'] + combineFile.fileData['CS'] + combineFile.fileData['Currency'] + combineFile.fileData['Material Code'] + combineFile.fileData['GPC Glo. Par. Code'] + combineFile.fileData['SAP No.'] + combineFile.fileData['Exchange Rate']
				# logFile['ID'] = logFile['Amount with VAT'] + logFile['CS'] + logFile['Currency'] + logFile['Material Code'] + logFile['GPC Glo. Par. Code'] + logFile['SAP No.'] + logFile['Exchange Rate']
				# mergeData = pd.merge(combineFile.fileData, logFile, on='ID', how='outer', indicator=True)
				# 原来根据多个字段meger
				# combineFile.fileData['SAP No.'] = combineFile.fileData['SAP No.'].apply(int)
				# logFile['SAP No.'] = logFile['SAP No.'].apply(int)
				onData = mergekeyFieldsList
				mergeData = pd.merge(combineFile.fileData, logFile.fileData, on=onData, how='outer', indicator=True)
				# mergeData = pd.merge(combineFile.fileData, logFile, on=['SAP No.'], how='outer', indicator=True)
				mergeData.sort_values(by=['Order No.'], axis=0, ascending=[True], inplace=True)
				# 保留数据
				leaveDataList = ["_merge", 'Project No._x', 'Order No.', 'Text', 'Long Text', 'Total Cost_x', 'Revenue\n(RMB)', 'SAP No._x', 'Project No._y', 'Remark', 'Update Time']
				leaveDataList += mergekeyFieldsList
				mergeData = mergeData[leaveDataList]
				ascendingList = [True] * len(leaveDataList)
				mergeData.sort_values(by=leaveDataList, axis=0, ascending=ascendingList, inplace=True)

				mergeDataName = '5.Order Merge Project'
				mergeFileNamePath = MyMainWindow.getFileName(self, fileUrl, mergeDataName, csvFileType)
				mergeFile = mergeData.to_csv('%s' % (mergeFileNamePath), encoding='utf_8_sig')
				self.textBrowser_2.append('Order NO 与 Project No合并的数据：%s' % mergeFileNamePath)
				self.textBrowser_2.append('Order Merge Project 数据,根据Order No数据透视算Amount with VAT的平均数值与ODM导出数据算Amount with VAT总值比较大小，有差说明错误。')
				self.textBrowser_2.append('SAP数据已处理完成')
				self.textBrowser_2.append('----------------------------------')
				os.startfile(combineFileUrl)
				os.startfile(mergeFileNamePath)
				os.startfile(fileUrl)
			else:
				self.textBrowser_2.append('请重新选择文件')
				self.textBrowser_2.append('----------------------------------')
		except Exception as msg:
			self.textBrowser_2.append('Order No Merge Project No数据有问题')
			self.textBrowser_2.append('错误信息：%s' % msg)
			self.textBrowser_2.append('----------------------------------')
			app.processEvents()
			# QMessageBox.information(self, "提示信息", '这份%s的ODM获取数据有问题' % fileData, QMessageBox.Yes)

	def pdfNameRule(self, msg):
		guiData = MyMainWindow.getAdminGuiData(self)
		pdfName = guiData['pdfName']
		pdfNameList = pdfName.split(' + ')
		changedPdfName = ''
		if msg in pdfNameList:
			pdfNameList.remove(msg)
			for each in pdfNameList:
				if changedPdfName != '':
					changedPdfName += ' + '
				changedPdfName += each
		else:
			changedPdfName += pdfName
			if changedPdfName != '':
				changedPdfName += ' + '
			changedPdfName += msg
		self.lineEdit_17.setText(changedPdfName)
		return changedPdfName


	def getFiles(self):
		selectBatchFile = QFileDialog.getOpenFileNames(self, '选择文件', '%s' % configContent['PDF_Files_Import_URL'], 'files(*.pdf)')
		self.filesUrl = selectBatchFile[0]
		if self.filesUrl != []:
			self.textBrowser_3.append('选中文件:')
			self.textBrowser_3.append('\n'.join(self.filesUrl))
			self.textBrowser_3.append('----------------------------------')
		else:
			self.textBrowser_3.append('无选中文件')
			self.textBrowser_3.append('----------------------------------')
		app.processEvents()
		return self.filesUrl


	def pdfOperate(self):
		fileUrls = self.filesUrl
		flag ='Y'
		if fileUrls == []:
			reply = QMessageBox.question(self, '信息', '没有选中文件，是否重新选择文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
			if reply == QMessageBox.Yes:
				fileUrls = MyMainWindow.getFiles(self)
				if fileUrls == []:
					flag = 'N'
			else:
				QMessageBox.information(self, "提示信息", "没有选中文件，请重新选择文件", QMessageBox.Yes)
				flag ='N'
		if flag == 'Y':
			guiData = MyMainWindow.getAdminGuiData(self)
			pdfOperate = PDF_Operate
			self.textBrowser_3.append('导出文件夹：%s' % configContent['PDF_Files_Export_URL'])
			self.textBrowser_3.append('导出文件名称：')
			i = 1
			for fileUrl in fileUrls:
				try:
					self.textBrowser_3.append('第%s份文件：' % i)
					msg = {}
					msg['Invoice No'] = ''
					with open(fileUrl, 'rb') as pdfFile:
						fileCon = pdfOperate.readPdf(pdfFile)
						fileNum = 0
						for fileCon[fileNum] in fileCon:
							if re.match('.*P. R. China', fileCon[fileNum]):
								msg['Company Name'] = fileCon[fileNum + 1].replace('Please quote this number on all inquiries and payments.', '').replace('Invoice No.', '')
							elif re.match('%s\d{%s}' % (guiData['invoiceStsrtNum'], int(guiData['invoiceBits'])-1), fileCon[fileNum]):
								msg['Invoice No'] = fileCon[fileNum]
							elif re.search('\d{2}.\d{3}.\d{2}.\d{4,5}', fileCon[fileNum]):
								res = fileCon[fileNum].split(' ')
								for each in res:
									if re.search('\d{2}.\d{3}.\d{2}.\d{4,5}', each):
										msg['Project No'] = each
									elif re.search('%s\d{%s}' % (guiData['invoiceStsrtNum'], int(guiData['invoiceBits'])-1), each) and msg['Invoice No'] == '':
										msg['Invoice No'] = each
							elif re.search('%s\d{%s}' % (guiData['orderStsrtNum'], int(guiData['orderBits'])-1), fileCon[fileNum]):
								res = fileCon[fileNum].split(' ')
								if len(res[1]) == int(guiData['orderBits']):
									msg['Order No'] = res[1]
							fileNum += 1
						pdfNameRule = guiData['pdfName'].split(' + ')
						outputFlieName = ''
						for eachName in pdfNameRule:
							if outputFlieName != '':
								outputFlieName += '-'
							if eachName == 'Invoice No':
								outputFlieName += msg['Invoice No']
							elif eachName == 'Company Name':
								outputFlieName += msg['Company Name']
							elif eachName == 'Order No':
								outputFlieName += msg['Order No']
							elif eachName == 'Project No':
								outputFlieName += msg['Project No']
						outputFlie = outputFlieName + '.pdf'
						# outputFlie = msg['Invoice No'] + '-' + msg['Company Name'] + '.pdf'
						pdfOperate.saveAs(fileUrl, '%s\\%s' % (configContent['PDF_Files_Export_URL'], outputFlie))
					self.textBrowser_3.append('%s' % outputFlie)
					app.processEvents()
				except Exception as errorMsg:
					# self.textBrowser_3.append("<font color='red'>第%s份文件：</font>" % i)
					self.textBrowser_3.append("<font color='red'>出错信息：%s </font>" % errorMsg)
					self.textBrowser_3.append("<font color='red'>出错的文件：%s </font>" % fileUrl)
				i += 1
			self.textBrowser_3.append('----------------------------------')





if __name__ == "__main__":
	import sys
	import os
	import re
	import time
	import math
	# import win32com.client
	import pandas as pd
	import numpy as np
	import datetime
	QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
	app = QApplication(sys.argv)
	myWin = MyMainWindow()
	myWin.show()
	myWin.getConfig()
	sys.exit(app.exec_())