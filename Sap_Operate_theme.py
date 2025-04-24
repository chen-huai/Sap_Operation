import sys
import os
import re
import time
import math
import pandas as pd
import csv
import copy
import numpy as np
import win32com.client
import datetime
import chicon  # 引用图标
# from PyQt5 import QtCore, QtGui, QtWidgets
# from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtWidgets import QApplication, QFileDialog, QMainWindow, QMessageBox, QVBoxLayout, QPushButton, QAction
from PyQt5.QtCore import QDate
from PyQt5.QtGui import QIcon
from Get_Data import *
from File_Operate import *
from PDF_Operate import *
from Sap_Function import *
from Sap_Operate_Ui import Ui_MainWindow
from Data_Table import *
from Logger import *
from Excel_Field_Mapper import excel_field_mapper
from theme_manager_theme import ThemeManager
from Revenue_Operate import *





class MyMainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyMainWindow, self).__init__(parent)
        self.setupUi(self)

        self.theme_manager = ThemeManager(QApplication.instance())
        self.init_theme_action()

        self.setGeometry(100, 100, 300, 200)

        self.theme_manager = ThemeManager(QApplication.instance())

        layout = QVBoxLayout()

        toggle_button = QPushButton("Toggle Theme")
        toggle_button.clicked.connect(self.theme_manager.toggle_theme)
        layout.addWidget(toggle_button)
        


        self.actionExport.triggered.connect(self.exportConfig)
        self.actionImport.triggered.connect(self.importConfig)
        self.actionExit.triggered.connect(MyMainWindow.close)
        self.actionHelp.triggered.connect(self.showVersion)
        self.actionAuthor.triggered.connect(self.showAuthorMessage)
        self.theme_manager.set_theme("blue")  # 设置默认主题
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
        self.pushButton_35.clicked.connect(self.invoiceRenameOperate)
        self.pushButton_33.clicked.connect(self.getPdfFiles)
        self.pushButton_48.clicked.connect(self.addMsg)
        self.pushButton_49.clicked.connect(self.viewOdmData)
        self.pushButton_59.clicked.connect(self.viewBillingListData)
        self.pushButton_50.clicked.connect(self.getEleInvoiceFiles)
        self.pushButton_51.clicked.connect(self.electronicInvoice)
        self.pushButton_58.clicked.connect(self.getBillingListFile)
        self.lineEdit_15.textChanged.connect(self.lineEditChange)
        self.doubleSpinBox_2.valueChanged.connect(self.getAmountVat)
        self.checkBox_9.toggled.connect(lambda: self.pdfNameRule('Invoice No', 'invoice'))
        self.checkBox_20.toggled.connect(lambda: self.pdfNameRule('Invoice No', 'Electron'))
        self.checkBox_10.toggled.connect(lambda: self.pdfNameRule('Company Name', 'invoice'))
        self.checkBox_22.toggled.connect(lambda: self.pdfNameRule('Company Name', 'Electron'))
        self.checkBox_12.toggled.connect(lambda: self.pdfNameRule('Order No', 'invoice'))
        self.checkBox_21.toggled.connect(lambda: self.pdfNameRule('Order No', 'Electron'))
        self.checkBox_11.toggled.connect(lambda: self.pdfNameRule('Project No', 'invoice'))
        self.checkBox_23.toggled.connect(lambda: self.pdfNameRule('FaPiao No', 'Electron'))
        self.checkBox_24.toggled.connect(lambda: self.pdfNameRule('Revenue', 'Electron'))
        self.checkBox_25.toggled.connect(lambda: self.pdfNameRule('CS', 'invoice'))
        self.checkBox_25.toggled.connect(lambda: self.pdfNameRule('CS', 'Electron'))
        self.checkBox_26.toggled.connect(lambda: self.pdfNameRule('Client Contact Name', 'invoice'))
        self.pushButton_56.clicked.connect(lambda: self.orderUnlockOrLock('Unlock'))
        self.pushButton_57.clicked.connect(lambda: self.orderUnlockOrLock('Lock'))
        self.pushButton_63.clicked.connect(lambda: self.get_hour_file_url(self.lineEdit_30))
        self.pushButton_71.clicked.connect(lambda: self.get_hour_file_url(self.lineEdit_38))
        self.pushButton_76.clicked.connect(lambda: self.get_hour_file_url(self.lineEdit_37))
        self.pushButton_72.clicked.connect(self.get_hour_combine_file)
        self.pushButton_77.clicked.connect(self.get_department_hour)
        self.pushButton_73.clicked.connect(self.get_person_hour)
        self.filesUrl = []

    def init_theme_action(self):
        theme_action = QAction(QIcon('theme_icon.png'), 'Toggle Theme', self)
        theme_action.setStatusTip('Toggle Theme')
        theme_action.triggered.connect(self.toggle_theme)

        # 将 action 添加到菜单（如果有的话）
        if hasattr(self, 'menuBar'):
            view_menu = self.menuBar().addMenu('Theme')
            view_menu.addAction(theme_action)

        # # 将 action 添加到工具栏
        # toolbar = self.addToolBar('主题')
        # toolbar.addAction(theme_action)

    def toggle_theme(self):
        self.theme_manager.set_random_theme()
        # 可以在这里添加其他需要在主题切换后更新的UI元素

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

    # 获取配置文件内容
    def getConfigContent(self):
        # 配置文件
        csvFile = pd.read_csv('%s/config_sap.csv' % configFileUrl, names=['A', 'B', 'C'])
        global configContent
        global username
        global role
        global staff_dict
        configContent = {}
        staff_dict = {}
        # configContent[configContent.get('Business_Department','CS')] = []
        # configContent[configContent.get('Lab_1','PHY')] = []
        # configContent[configContent.get('Lab_2','CHM')] = []
        username = list(csvFile['A'])
        number = list(csvFile['B'])
        role = list(csvFile['C'])
        for i in range(len(username)):
            configContent['%s' % username[i]] = number[i]
            if role[i] == configContent.get('Business_Department', 'CS'):
                # 使用 setdefault 确保键存在且为列表类型
                staff_dict.setdefault(configContent.get('Business_Department', 'CS'), []).append(username[i])
            if role[i] == configContent.get('Lab_1', 'PHY'):
                # 使用 setdefault 确保键存在且为列表类型
                staff_dict.setdefault(configContent.get('Lab_1', 'PHY'), []).append(username[i])
            if role[i] == configContent.get('Lab_2', 'CHM'):
                # 使用 setdefault 确保键存在且为列表类型
                staff_dict.setdefault(configContent.get('Lab_2', 'CHM'), []).append(username[i])
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

    # 创建配置文件
    def createConfigContent(self):
        global monthAbbrev
        months = "JanFebMarAprMayJunJulAugSepOctNovDec"
        n = time.strftime('%m')
        pos = (int(n) - 1) * 3
        monthAbbrev = months[pos:pos + 3]

        configContent = [
            ['特殊开票', '内容', '备注'],
            ['SAP_Date_URL', 'N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\收样\\3.Sap\\ODM Data - XM',
             '文件数据路径'],
            ['Invoice_File_URL',
             'N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\收样\\3.Sap\\ODM Data - XM\\2.特殊开票',
             '特殊开票文件路径'],
            ['Invoice_File_Name', '特殊开票要求2022.xlsx', '特殊开票文件名称'],
            ['Data数据处理', '内容', '备注'],
            ['Row Data', 'Client Contact Name', '以;分隔，横向添加数据'],
            ['Column Data', 'Project No.;Currency;Amount with VAT;Reference No.', '以;分隔，纵向添加数据'],
            ['Row Check', 0, '是否默认被选中,1选中，0未选中'],
            ['Column Check', 0, '是否默认被选中,1选中，0未选中'],
            ['Combine Key', "CS;Sales;Currency;Material Code;Invoices' name (Chinese);Buyer(GPC);Month;Exchange Rate",
             '以;分隔，数据透视字段'],
            ['SAP登入信息', '内容', '备注'],
            ['Login_msg', 'DR-0486-01->601-240', '订单类型-销售组织-分销渠道-销售办事处-销售组'],
            ['Business_Department', 'CS', '业务部门,名称会用于后续'],
            ['Lab_1', 'PHY', '代表实验室，会用于后续'],
            ['Lab_2', 'CHM', '代表实验室，会用于后续'],
            ['T20', 'PHY', '代表实验室，会用于后续'],
            ['T75', 'CHM', '代表实验室，会用于后续'],
            ['Hourly Rate', '金额', '备注'],
            ['CS_Hourly_Rate', 300, '客服时薪'],
            ['PHY_Hourly_Rate', 300, '物理时薪'],
            ['CHM_Hourly_Rate', 300, '化学时薪'],
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
            # 新增分配规则参数
            ['405_Item_1000', 0.5, '405分配规则'],
            ['405_Item_2000', 0.5, '405分配规则'],
            ['441_Item_1000', 0.8, '441分配规则'],
            ['441_Item_2000', 0.2, '441分配规则'],
            ['430_Item_1000', 0.8, '430分配规则'],
            ['430_Item_2000', 0.2, '430分配规则'],
            # 新增特殊MC规则
            ['T20-430-A2', 'PHY_1000/CHM_2000', '1000/2000对应的lab'],
            ['T20-430-A2_mc', 'T20-430-00/T75-430-00', '1000/2000对应的mc'],
            ['T75-441-A2', 'CHM_1000/PHY_2000', '1000/2000对应的lab'],
            ['T75-441-A2_mc', 'T75-441-00/T20-441-00', '1000/2000对应的mc'],
            ['T75-405-A2', 'CHM_1000/PHY_2000', '1000/2000对应的lab'],
            ['T75-405-A2_mc', 'T75-405-00/T20-405-00', '1000/2000对应的mc'],
            # 新增公共参数
            ['Max_Hour', 8, '最大工作时长'],
            ['Hours_Combine_Key', "Order Number;Material Code",'以;分隔，数据透视字段'],
            ['Hour_Files_Import_URL', "N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\2.财务\\2.SAP\\1.ODM Data - XM\\3.Hours",'Invoice文件导入路径'],
            ['Hour_Files_Export_URL', "N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\2.财务\\2.SAP\\1.ODM Data - XM\\3.Hours",'Invoice文件导入路径'],
            ['DATA A数据填写', '判断依据', '备注'],
            ['Data_A_E1', '5010815347;5010427355;5010913488;5010685589;5010829635;5010817524', 'Data A录E1,新添加用;隔开即可'],
            ['Data_A_Z2', '5010908478;5010823259', 'Data A录Z2,新添加用;隔开即可'],
            ['SAP操作', '内容', '备注'],
            ['Cost_VAT_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['NVA01_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['NVA02_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['NVF01_Selected', 0, '是否默认被选中,1选中，0未选中'],
            ['NVF03_Selected', 0, '是否默认被选中,1选中，0未选中'],
            ['DataB_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['Plan_Cost_Selected', 25, '每月超过几号自动选中（不包含）'],
            ['Save_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['Every_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['Contact_Selected', 0, '是否默认被选中,1选中，0未选中'],
            ['管理操作', '内容', '备注'],
            ['Billing_List_URL',
             'N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\2.财务\\3.Billing存档\\4.XM-billing list\\2023',
             'Billing list文件默认路径'],
            ['Add_CS_Msg_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['Invoice_No_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['Invoice_Start_Num', 4, 'Invoice的起始数字'],
            ['Invoice_Num', 9, 'Invoice的总位数'],
            ['Company_Name_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['Order_No_Selected', 0, '是否默认被选中,1选中，0未选中'],
            ['Invoice_Contact_Selected', 0, '是否默认被选中,1选中，0未选中'],
            ['Order_Start_Num', 7, 'Order的起始数字'],
            ['Order_Num', 9, 'Order的总位数'],
            ['Project_No_Selected', 0, '是否默认被选中,1选中，0未选中'],
            ['Invoice_Name', 'CS + Invoice No + Company Name', 'Invoice文件名称默认规则'],
            ['Invoice_Files_Import_URL', desktopUrl, 'Invoice文件导入路径'],
            ['Invoice_Files_Export_URL', 'N:\\XM Softlines\\1. Project\\3. Finance\\02. WIP', 'Invoice文件导出路径'],
            ['Ele_Invoice_No_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['Ele_Invoice_Start_Num', 486, '电子发票的起始数字'],
            ['Ele_Invoice_Num', 9, 'Invoice的总位数'],
            ['Ele_Order_No_Selected', 0, '是否默认被选中,1选中，0未选中'],
            ['Ele_Order_Start_Num', 7486, '电子发票的起始数字'],
            ['Ele_Order_Num', 9, 'Order的总位数'],
            ['Ele_Company_Name_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['Ele_Revenue_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['Ele_Fapiao_No_Selected', 0, '是否默认被选中,1选中，0未选中'],
            ['Ele_Invoice_Name', 'CS + Company Name + Invoice No + Revenue', '电子发票文件名称默认规则'],
            ['Ele_Invoice_Files_Import_URL', 'N:\\Company Data\\FCO\\11.全电发票', '全电发票路径'],
            ['Ele_Invoice_Files_Export_URL', 'N:\\XM Softlines\\1. Project\\3. Finance\\02. WIP\\全电发票 2023\\10',
             '全电发票导出路径'],
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

    # 导出配置文件
    def exportConfig(self):
        # 重新导出默认配置文件
        reply = QMessageBox.question(self, '信息', '确认是否要创建默认配置文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if reply == QMessageBox.Yes:
            MyMainWindow.createConfigContent(self)
        else:
            QMessageBox.information(self, "提示信息", "没有创建默认配置文件，保留原有的配置文件", QMessageBox.Yes)

    # 导入配置文件
    def importConfig(self):
        # 重新导入配置文件
        reply = QMessageBox.question(self, '信息', '确认是否要导入配置文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if reply == QMessageBox.Yes:
            MyMainWindow.getConfigContent(self)
        else:
            QMessageBox.information(self, "提示信息", "没有重新导入配置文件，将按照原有的配置文件操作", QMessageBox.Yes)

    # 界面设置默认配置文件信息
    def getDefaultInformation(self):
        # 默认登录界面信息
        try:
            # data处理
            self.checkBox_17.setChecked(int(configContent['Row Check']))
            self.checkBox_18.setChecked(int(configContent['Column Check']))
            self.lineEdit_23.setText(configContent['Row Data'])
            self.lineEdit_24.setText(configContent['Column Data'])
            self.lineEdit_15.setText(configContent['Combine Key'])
            self.lineEdit_16.setText(configContent['Combine Key'])
            # login信息
            loginMsgList = configContent['Login_msg'].split('-')
            self.lineEdit_10.setText(loginMsgList[0])
            self.lineEdit_11.setText(loginMsgList[1])
            self.lineEdit_12.setText(loginMsgList[2])
            self.lineEdit_13.setText(loginMsgList[3])
            self.lineEdit_14.setText(loginMsgList[4])
            # 每小时成本
            self.doubleSpinBox_5.setValue(float(format(float(configContent['CS_Hourly_Rate']), '.2f')))
            self.doubleSpinBox_6.setValue(float(format(float(configContent['CHM_Hourly_Rate']), '.2f')))
            self.doubleSpinBox_8.setValue(float(format(float(configContent['PHY_Hourly_Rate']), '.2f')))
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
            # COST是否含税
            self.checkBox_27.setChecked(int(configContent['Cost_VAT_Selected']))
            # SAP操作
            self.checkBox.setChecked(int(configContent['NVA01_Selected']))
            self.checkBox_2.setChecked(int(configContent['NVA02_Selected']))
            self.checkBox_3.setChecked(int(configContent['NVF01_Selected']))
            self.checkBox_4.setChecked(int(configContent['NVF03_Selected']))
            self.checkBox_7.setChecked(int(configContent['DataB_Selected']))
            self.checkBox_6.setChecked(int(configContent['Save_Selected']))
            self.checkBox_16.setChecked(int(configContent['Every_Selected']))
            self.checkBox_19.setChecked(int(configContent['Contact_Selected']))
            if int(configContent['Plan_Cost_Selected']) < int(today.split('.')[-1]):
                self.checkBox_8.setChecked(True)
            # admin操作
            self.checkBox_25.setChecked(int(configContent['Add_CS_Msg_Selected']))
            self.checkBox_9.setChecked(int(configContent['Invoice_No_Selected']))
            self.spinBox.setValue(int(configContent['Invoice_Start_Num']))
            self.spinBox_2.setValue(int(configContent['Invoice_Num']))
            self.checkBox_10.setChecked(int(configContent['Company_Name_Selected']))
            self.checkBox_12.setChecked(int(configContent['Order_No_Selected']))
            self.checkBox_26.setChecked(int(configContent['Invoice_Contact_Selected']))
            self.spinBox_3.setValue(int(configContent['Order_Start_Num']))
            self.spinBox_4.setValue(int(configContent['Order_Num']))
            self.checkBox_11.setChecked(int(configContent['Project_No_Selected']))
            self.lineEdit_17.setText(configContent['Invoice_Name'])
            self.checkBox_20.setChecked(int(configContent['Ele_Invoice_No_Selected']))
            self.spinBox_8.setValue(int(configContent['Ele_Invoice_Start_Num']))
            self.spinBox_6.setValue(int(configContent['Ele_Invoice_Num']))
            self.checkBox_21.setChecked(int(configContent['Ele_Order_No_Selected']))
            self.spinBox_9.setValue(int(configContent['Ele_Order_Start_Num']))
            self.spinBox_7.setValue(int(configContent['Ele_Order_Num']))
            self.checkBox_22.setChecked(int(configContent['Ele_Company_Name_Selected']))
            self.checkBox_23.setChecked(int(configContent['Ele_Fapiao_No_Selected']))
            self.checkBox_24.setChecked(int(configContent['Ele_Revenue_Selected']))
            self.lineEdit_27.setText(configContent['Ele_Invoice_Name'])
            # hour界面操作
            self.spinBox_10.setValue(int(configContent['Max_Hour']))
            self.lineEdit_39.setText(configContent['Hours_Combine_Key'])
            today_hours = datetime.date.today()
            first_day = today_hours.replace(day=1)
            self.dateEdit.setDate(QDate(first_day.year, first_day.month, first_day.day))  # 当月第一天
            self.dateEdit_2.setDate(QDate.currentDate())  # 当天日期
            self.doubleSpinBox_12.setValue(float(format(float(configContent['CS_Hourly_Rate']), '.2f')))
            self.doubleSpinBox_16.setValue(float(format(float(configContent['CHM_Hourly_Rate']), '.2f')))
            self.doubleSpinBox_15.setValue(float(format(float(configContent['PHY_Hourly_Rate']), '.2f')))
            self.doubleSpinBox_11.setValue(float(format(float(configContent['Plan_Cost_Parameter']), '.2f')))
            self.doubleSpinBox_13.setValue(float(format(float(configContent['CHM_Cost_Parameter']), '.2f')))
            self.doubleSpinBox_12.setValue(float(format(float(configContent['PHY_Cost_Parameter']), '.2f')))
            self.spinBox_11.setValue(int(configContent['Significant_Digits']))
            self.lineEdit_28.setText(configContent['Lab_1'])
            self.lineEdit_29.setText(configContent['Lab_2'])
        except Exception as msg:
            self.textBrowser_2.append("错误信息：%s" % msg)
            self.textBrowser_2.append('----------------------------------')
            app.processEvents()
            reply = QMessageBox.question(self, '信息', '错误信息：%s。\n是否要重新创建配置文件' % msg, QMessageBox.Yes | QMessageBox.No,
                                         QMessageBox.Yes)
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
        self.doubleSpinBox_4.setValue(amount * 1.06)

    # 获取SAP配置信息
    def getGuiData(self):
        guiData = {}
        guiData['sapNo'] = self.lineEdit.text()
        guiData['projectNo'] = self.lineEdit_2.text()
        guiData['materialCode'] = self.comboBox_4.currentText()
        guiData['currencyType'] = self.comboBox.currentText()
        guiData['exchangeRate'] = float(self.doubleSpinBox.text())
        guiData['globalPartnerCode'] = self.lineEdit_3.text()
        guiData['csName'] = self.comboBox_2.currentText()
        if guiData['csName'] != '':
            guiData['csCode'] = configContent[guiData['csName']]
        guiData['salesName'] = self.comboBox_3.currentText()
        if guiData['salesName'] != '':
            guiData['salesCode'] = configContent[guiData['salesName']]
        guiData['amount'] = float(self.doubleSpinBox_2.text())
        if self.checkBox_27.isChecked():
            guiData['cost'] = float(self.doubleSpinBox_3.text())/1.06
        else:
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
        guiData['orderType'] = self.lineEdit_10.text()
        guiData['salesOrganization'] = self.lineEdit_11.text()
        guiData['distributionChannels'] = self.lineEdit_12.text()
        guiData['salesOffice'] = self.lineEdit_13.text()
        guiData['salesGroup'] = self.lineEdit_14.text()
        guiData['csCostCenter'] = self.lineEdit_18.text()
        guiData['chmCostCenter'] = self.lineEdit_19.text()
        guiData['phyCostCenter'] = self.lineEdit_20.text()
        if self.checkBox.isChecked():
            guiData['va01Check'] = True
        else:
            guiData['va01Check'] = False

        if self.checkBox_2.isChecked():
            guiData['va02Check'] = True
        else:
            guiData['va02Check'] = False

        if self.checkBox_3.isChecked():
            guiData['vf01Check'] = True
        else:
            guiData['vf01Check'] = False

        if self.checkBox_4.isChecked():
            guiData['vf03Check'] = True
        else:
            guiData['vf03Check'] = False

        if self.checkBox_6.isChecked():
            guiData['saveCheck'] = True
        else:
            guiData['saveCheck'] = False

        if self.checkBox_7.isChecked():
            guiData['labCostCheck'] = True
        else:
            guiData['labCostCheck'] = False

        if self.checkBox_8.isChecked():
            guiData['planCostCheck'] = True
        else:
            guiData['planCostCheck'] = False

        if self.checkBox_13.isChecked():
            guiData['csCheck'] = True
        else:
            guiData['csCheck'] = False

        if self.checkBox_14.isChecked():
            guiData['chmCheck'] = True
        else:
            guiData['chmCheck'] = False

        if self.checkBox_15.isChecked():
            guiData['phyCheck'] = True
        else:
            guiData['phyCheck'] = False

        if self.checkBox_16.isChecked():
            guiData['everyCheck'] = True
        else:
            guiData['everyCheck'] = False

        if self.checkBox_17.isChecked():
            guiData['rowCheck'] = True
        else:
            guiData['rowCheck'] = False

        if self.checkBox_18.isChecked():
            guiData['columnCheck'] = True
        else:
            guiData['columnCheck'] = False

        if self.checkBox_19.isChecked():
            guiData['contactCheck'] = True
        else:
            guiData['contactCheck'] = False

        return guiData

    # 获取Invoice配置信息
    def getAdminGuiData(self):
        guiAdminData = {}
        # guiAdminData['BillingList'] = self.lineEdit_25.text()
        guiAdminData['invoiceStsrtNum'] = int(self.spinBox.text())
        guiAdminData['invoiceBits'] = int(self.spinBox_2.text())
        guiAdminData['orderStsrtNum'] = int(self.spinBox_3.text())
        guiAdminData['orderBits'] = int(self.spinBox_4.text())
        guiAdminData['pdfName'] = self.lineEdit_17.text()
        guiAdminData['eleInvoiceStsrtNum'] = int(self.spinBox_8.text())
        guiAdminData['eleOrderStsrtNum'] = int(self.spinBox_9.text())
        guiAdminData['eleInvoiceBits'] = int(self.spinBox_6.text())
        guiAdminData['eleOrderBits'] = int(self.spinBox_7.text())
        guiAdminData['fapiaoName'] = self.lineEdit_27.text()
        return guiAdminData

    def getHourGuiData(self):
        guiHourData = {}
        guiHourData['Hours_Combine_Key'] = self.lineEdit_39.text()
        guiHourData['Max_Hour'] = int(self.spinBox_10.text())
        guiHourData['CS_Hourly_Rate'] = float(self.doubleSpinBox_14.text())
        guiHourData['CHM_Hourly_Rate'] = float(self.doubleSpinBox_16.text())
        guiHourData['PHY_Hourly_Rate'] = float(self.doubleSpinBox_15.text())
        guiHourData['Plan_Cost_Parameter'] = float(self.doubleSpinBox_11.text())
        guiHourData['Significant_Digits'] = float(self.spinBox_11.text())
        guiHourData['CHM_Cost_Parameter'] = float(self.doubleSpinBox_13.text())
        guiHourData['PHY_Cost_Parameter'] = float(self.doubleSpinBox_12.text())
        guiHourData['Lab_1'] = self.lineEdit_28.text()
        guiHourData['Lab_2'] = self.lineEdit_29.text()
        guiHourData['Business_Department'] = self.lineEdit_26.text()
        guiHourData['Start_Date'] = self.dateEdit.date().toString("yyyy.MM.dd")
        guiHourData['End_Date'] = self.dateEdit_2.date().toString("yyyy.MM.dd")
        return guiHourData

    # 计算Order所需数据
    def getRevenueData(self, guiData):
        # 计算金额
        # revenue,planCost,revenueForCny,chmCost,phyCost,chmRe,phyRe,chmCsCostAccounting,chmLabCostAccounting,phyCsCostAccounting
        revenueData = {}
        revenueData['revenue'] = guiData['amountVat'] / 1.06
        # plan cost
        # planCost = revenueData['revenue'] * guiData['exchangeRate'] * 0.9 - guiData['cost']
        revenueData['planCost'] = revenueData['revenue'] * guiData['exchangeRate']
        revenueData['revenueForCny'] = revenueData['revenue'] * guiData['exchangeRate']

        # 405-A2的计算公式
        if ('405' in guiData['materialCode']) and (
                ("A2" in guiData['materialCode']) or ("D2" in guiData['materialCode']) or (
                "D3" in guiData['materialCode'])):
            # DataB-CHM成本
            revenueData['chmCost'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['chmCostRate'] * 0.5, '.2f')
            # DataB-PHY成本
            revenueData['phyCost'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['phyCostRate'] * 0.5, '.2f')
            # Item1000 的revenue
            revenueData['chmRe'] = format(revenueData['revenue'] * 0.5, '.2f')
            # Item2000 的revenue
            revenueData['phyRe'] = format(revenueData['revenue'] * 0.5, '.2f')
            # plan cost总算法
            # revenueData['chmCsCostAccounting'] = format(revenueData['planCost'] * 0.5 * (1 - 0.3  - (1 - guiData['planCostRate'] )) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
            # revenueData['chmLabCostAccounting'] = format(revenueData['planCost'] * 0.5 * 0.3 / guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])
            # revenueData['phyCsCostAccounting'] = format(revenueData['planCost'] * 0.5 * (1 - 0.3  - (1 - guiData['planCostRate'] )) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
            # revenueData['phyLabCostAccounting'] = format(revenueData['planCost'] * 0.5 * 0.3 / guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])

            # plan cost，理论上（revenue-total cost）*0.9*0.5，实际上SFL省略了0.9的计算（金额不大）

            # CS的Item1000-Cost
            revenueData['chmCsCostAccounting'] = format(
                (revenueData['revenueForCny'] * guiData['planCostRate'] - guiData['cost']) * 0.5 * (
                        1 - guiData['chmCostRate']) / guiData['csHourlyRate'],
                '.%sf' % guiData['significantDigits'])
            # CHM的Item1000-Cost
            revenueData['chmLabCostAccounting'] = format(
                (revenueData['revenueForCny'] * guiData['planCostRate'] - guiData['cost']) * 0.5 * guiData[
                    'chmCostRate'] /
                guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])
            # CS的Item2000-Cost
            revenueData['phyCsCostAccounting'] = format(
                (revenueData['revenueForCny'] * guiData['planCostRate'] - guiData['cost']) * 0.5 * (
                        1 - guiData['phyCostRate']) / guiData['csHourlyRate'],
                '.%sf' % guiData['significantDigits'])
            # PHY的Item2000-Cost
            revenueData['phyLabCostAccounting'] = format(
                (revenueData['revenueForCny'] * guiData['planCostRate'] - guiData['cost']) * 0.5 * guiData[
                    'phyCostRate'] /
                guiData['phyHourlyRate'], '.%sf' % guiData['significantDigits'])

        # 441-A2计算公式
        elif ('441' in guiData['materialCode']) and ((
                "A2" in guiData['materialCode'] or ("D2" in guiData['materialCode']) or (
                "D3" in guiData['materialCode']))):
            # DataB-CHM成本
            revenueData['chmCost'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['chmCostRate'] * 0.8, '.2f')
            # DataB-PHY成本
            revenueData['phyCost'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['phyCostRate'] * 0.2, '.2f')
            # Item1000 的revenue
            revenueData['chmRe'] = format(revenueData['revenue'] * 0.8, '.2f')
            # Item2000 的revenue
            revenueData['phyRe'] = format(revenueData['revenue'] * 0.2, '.2f')
            # plan cost总算法
            # revenueData['chmCsCostAccounting'] = format(revenueData['planCost'] * 0.8 * (1 - 0.3  - (1 - guiData['planCostRate'] )) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
            # revenueData['chmLabCostAccounting'] = format(revenueData['planCost'] * 0.8 * 0.3 / guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])
            # revenueData['phyCsCostAccounting'] = format(revenueData['planCost'] * 0.2 * (1 - 0.3  - (1 - guiData['planCostRate'] )) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
            # revenueData['phyLabCostAccounting'] = format(revenueData['planCost'] * 0.2 * 0.3 / guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])

            # CS的Item1000-Cost
            revenueData['chmCsCostAccounting'] = format(
                (revenueData['revenueForCny'] * guiData['planCostRate'] - guiData['cost']) * 0.8 * (
                        1 - guiData['chmCostRate']) / guiData['csHourlyRate'],
                '.%sf' % guiData['significantDigits'])
            # CHM的Item1000-Cost
            revenueData['chmLabCostAccounting'] = format(
                (revenueData['revenueForCny'] * guiData['planCostRate'] - guiData['cost']) * 0.8 * guiData[
                    'chmCostRate'] /
                guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])
            # CS的Item2000-Cost
            revenueData['phyCsCostAccounting'] = format(
                (revenueData['revenueForCny'] * guiData['planCostRate'] - guiData['cost']) * 0.2 * (
                        1 - guiData['phyCostRate']) / guiData['csHourlyRate'],
                '.%sf' % guiData['significantDigits'])
            # PHY的Item2000-Cost
            revenueData['phyLabCostAccounting'] = format(
                (revenueData['revenueForCny'] * guiData['planCostRate'] - guiData['cost']) * 0.2 * guiData[
                    'phyCostRate'] /
                guiData['phyHourlyRate'], '.%sf' % guiData['significantDigits'])

        # 430-A2计算公式
        elif ('430' in guiData['materialCode']) and (
                "A2" in guiData['materialCode']):
            # DataB-CHM成本
            revenueData['chmCost'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['chmCostRate'] * 0.2, '.2f')
            # DataB-PHY成本
            revenueData['phyCost'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['phyCostRate'] * 0.8, '.2f')
            # Item1000 的revenue
            revenueData['chmRe'] = format(revenueData['revenue'] * 0.2, '.2f')
            # Item2000 的revenue
            revenueData['phyRe'] = format(revenueData['revenue'] * 0.8, '.2f')
            # plan cost总算法
            # revenueData['chmCsCostAccounting'] = format(revenueData['planCost'] * 0.8 * (1 - 0.3  - (1 - guiData['planCostRate'] )) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
            # revenueData['chmLabCostAccounting'] = format(revenueData['planCost'] * 0.8 * 0.3 / guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])
            # revenueData['phyCsCostAccounting'] = format(revenueData['planCost'] * 0.2 * (1 - 0.3  - (1 - guiData['planCostRate'] )) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
            # revenueData['phyLabCostAccounting'] = format(revenueData['planCost'] * 0.2 * 0.3 / guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])

            # CS的Item1000-Cost
            revenueData['chmCsCostAccounting'] = format(
                (revenueData['revenueForCny'] * guiData['planCostRate'] - guiData['cost']) * 0.2 * (
                        1 - guiData['chmCostRate']) / guiData['csHourlyRate'],
                '.%sf' % guiData['significantDigits'])
            # CHM的Item1000-Cost
            revenueData['chmLabCostAccounting'] = format(
                (revenueData['revenueForCny'] * guiData['planCostRate'] - guiData['cost']) * 0.2 * guiData[
                    'chmCostRate'] /
                guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])
            # CS的Item2000-Cost
            revenueData['phyCsCostAccounting'] = format(
                (revenueData['revenueForCny'] * guiData['planCostRate'] - guiData['cost']) * 0.8 * (
                        1 - guiData['phyCostRate']) / guiData['csHourlyRate'],
                '.%sf' % guiData['significantDigits'])
            # PHY的Item2000-Cost
            revenueData['phyLabCostAccounting'] = format(
                (revenueData['revenueForCny'] * guiData['planCostRate'] - guiData['cost']) * 0.8 * guiData[
                    'phyCostRate'] /
                guiData['phyHourlyRate'], '.%sf' % guiData['significantDigits'])
        else:
            revenueData['chmCost'] = format((revenueData['revenueForCny'] - guiData['cost']) * guiData['chmCostRate'],
                                            '.2f')
            revenueData['phyCost'] = format((revenueData['revenueForCny'] - guiData['cost']) * guiData['phyCostRate'],
                                            '.2f')
            revenueData['chmRe'] = format(revenueData['revenue'], '.2f')
            revenueData['phyRe'] = format(revenueData['revenue'], '.2f')
            # plan cost总算法
            # csCostAccounting = format(planCost * (1 - 0.3  - (1 - guiData['planCostRate'] )) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
            # labCostAccounting = format(planCost * 0.3 / guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])
            if 'T75' in guiData['materialCode']:
                revenueData['labCostRate'] = guiData['chmCostRate']
                revenueData['labHourlyRate'] = guiData['chmHourlyRate']
            else:
                revenueData['labCostRate'] = guiData['phyCostRate']
                revenueData['labHourlyRate'] = guiData['phyHourlyRate']

            revenueData['csCostAccounting'] = format(
                (revenueData['revenueForCny'] * guiData['planCostRate'] - guiData['cost']) * (
                        1 - revenueData['labCostRate']) / guiData[
                    'csHourlyRate'], '.%sf' % guiData['significantDigits'])
            revenueData['labCostAccounting'] = format(
                (revenueData['revenueForCny'] * guiData['planCostRate'] - guiData['cost']) * revenueData[
                    'labCostRate'] / revenueData['labHourlyRate'],
                '.%sf' % guiData['significantDigits'])
        return revenueData

    # SAP开Order操作
    def sapOperate(self, sap_obj):
        logMsg = {}
        logMsg['Remark'] = ''
        logMsg['orderNo'] = ''
        logMsg['Proforma No.'] = ''
        logMsg['sapAmountVat'] = ''
        try:
            flag = 1
            # 获取数据
            guiData = MyMainWindow.getGuiData(self)
            orderNo = ''
            proformaNo = ''
            if guiData['everyCheck']:
                sap_obj = Sap()
            if guiData['sapNo'] == '' or guiData['projectNo'] == '' or guiData['materialCode'] == '' or guiData[
                'currencyType'] == '' or guiData['exchangeRate'] == '' or guiData['globalPartnerCode'] == '' or guiData[
                'csName'] == '' or guiData['amount'] == 0.00 or guiData['amountVat'] == 0.00:
                self.textBrowser.append("<font color='red'>有关键信息未填</font>")
                logMsg['Remark'] = '有关键信息未填'
                self.textBrowser.append(
                    "'Project No.', 'CS', 'Sales', 'Currency', 'GPC Glo. Par. Code', 'Material Code','SAP No.', 'Amount', 'Amount with VAT', 'Exchange Rate'都是必须填写的")
                self.textBrowser.append('----------------------------------')
                app.processEvents()
                if guiData['everyCheck']:
                    QMessageBox.information(self, "提示信息", "有关键信息未填", QMessageBox.Yes)
            else:
                revenueData = MyMainWindow.getRevenueData(self, guiData)
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
                    self.textBrowser.append("CHM Cost:%s" % revenueData['chmCost'])
                    self.textBrowser.append("PHY Cost:%s" % revenueData['phyCost'])
                    self.textBrowser.append("CHM Amount:%s" % revenueData['chmRe'])
                    self.textBrowser.append("PHY Amount:%s" % revenueData['phyRe'])
                    app.processEvents()

                    flag = 1
                    # VA01
                    if guiData['va01Check']:
                        va01_res = sap_obj.va01_operate(guiData, revenueData)
                        if va01_res['flag'] == 1:
                            # 是否要添加lab cost
                            if guiData['labCostCheck'] and va01_res['flag'] == 1:
                                data_b_res = sap_obj.lab_cost(guiData, revenueData)
                                if data_b_res['flag'] == 0:
                                    logMsg['Remark'] += data_b_res['msg']
                                    self.textBrowser.append("<font color='red'>出错信息：%s </font>" % data_b_res['msg'])
                                    app.processEvents()
                                    if guiData['everyCheck']:
                                        QMessageBox.information(self, "错误提示", "出错信息：%s" % data_b_res['msg'],
                                                                QMessageBox.Yes)
                            if guiData['va02Check'] or guiData['saveCheck']:
                                save_res = sap_obj.save_sap('VA01')
                                if save_res['flag'] == 0:
                                    flag = 0
                                    logMsg['Remark'] += ';' + save_res['msg']
                                    self.textBrowser.append("<font color='red'>出错信息：%s </font>" % save_res['msg'])
                                    app.processEvents()
                                    if guiData['everyCheck']:
                                        QMessageBox.information(self, "错误提示", "出错信息：%s" % save_res['msg'],
                                                                QMessageBox.Yes)
                        else:
                            flag = 0
                            logMsg['Remark'] += va01_res['msg']
                            self.textBrowser.append("<font color='red'>出错信息：VA01出错；%s </font>" % va01_res['msg'])
                            app.processEvents()
                            if guiData['everyCheck']:
                                QMessageBox.information(self, "错误提示", "出错信息：%s" % va01_res['msg'],
                                                        QMessageBox.Yes)
                    # VA02
                    if guiData['va02Check'] and flag == 1:
                        va02_res = sap_obj.va02_operate(guiData, revenueData)
                        logMsg['orderNo'] = va02_res['orderNo']
                        self.textBrowser.append("Order No.:%s" % logMsg['orderNo'])
                        app.processEvents()
                        if va02_res['flag'] == 1:
                            amountVatStr = re.sub(r"(\d)(?=(\d\d\d)+(?!\d))", r"\1,",
                                                  format(guiData['amountVat'], '.2f'))
                            sapAmountVat = va02_res['sapAmountVat']
                            self.textBrowser.append("Sap Amount Vat:%s" % sapAmountVat)
                            self.textBrowser.append("Amount Vat:%s" % amountVatStr)
                            app.processEvents()
                            # sapAmountVat在A2是数字，其它为字符串
                            if sapAmountVat.strip() != amountVatStr:
                                self.textBrowser.append("<font color='blue'>提示信息：SAP数据与ODM不一致，请确认并修改后再继续！！！ </font>")
                                app.processEvents()
                                logMsg['Remark'] += ';' + 'SAP数据与ODM不一致，请确认并修改后再继续！！！'
                                if guiData['everyCheck']:
                                    reply = QMessageBox.question(self, '信息', 'SAP数据与ODM不一致，请确认并修改后再继续！！！',
                                                                 QMessageBox.Yes | QMessageBox.No,
                                                                 QMessageBox.Yes)
                                    if reply == QMessageBox.No:
                                        flag = 0
                            if (guiData['vf01Check'] or guiData['saveCheck']) and flag == 1:
                                sava_res = sap_obj.save_sap('VA02')
                                if sava_res['flag'] == 0:
                                    flag = 0
                                    logMsg['Remark'] += ';' + sava_res['msg']
                                    self.textBrowser.append("<font color='red'>出错信息：%s </font>" % sava_res['msg'])
                                    app.processEvents()
                                    if guiData['everyCheck']:
                                        QMessageBox.information(self, "错误提示", "出错信息：%s" % sava_res['msg'],
                                                                QMessageBox.Yes)
                        else:
                            flag = 0
                            logMsg['Remark'] += ';' + va02_res['msg']
                            self.textBrowser.append("<font color='red'>出错信息：VA02出错；%s </font>" % va02_res['msg'])
                            app.processEvents()
                            if guiData['everyCheck']:
                                QMessageBox.information(self, "错误提示", "出错信息：%s" % va02_res['msg'],
                                                        QMessageBox.Yes)

                    # VF01
                    if guiData['vf01Check'] and flag == 1:

                        # save_res = sap_obj.save_sap('VF01准备前')
                        # if save_res['flag'] == 0:
                        #     logMsg['Remark'] += ';' + save_res['msg']
                        #     self.textBrowser.append("<font color='red'>出错信息：%s </font>" % save_res['msg'])
                        #     app.processEvents()
                        #     if guiData['everyCheck']:
                        #         QMessageBox.information(self, "错误提示", "出错信息：%s" % save_res['msg'],
                        #                                 QMessageBox.Yes)
                        vf01_res = sap_obj.vf01_operate()
                        if vf01_res['flag'] == 0:
                            flag = 0
                            logMsg['Remark'] += ';' + vf01_res['msg']
                            self.textBrowser.append("<font color='red'>出错信息：VF01出错；%s </font>" % vf01_res['msg'])
                            app.processEvents()
                            if guiData['everyCheck']:
                                QMessageBox.information(self, "错误提示", "出错信息：%s" % vf01_res['msg'],
                                                        QMessageBox.Yes)
                    # VF03
                    if guiData['vf03Check'] and flag == 1:
                        vf03_res = sap_obj.vf03_operate()
                        if vf03_res['flag'] == 0:
                            logMsg['Remark'] += ';' + vf03_res['msg']
                            self.textBrowser.append("<font color='red'>出错信息：VF03出错;%s </font>" % vf03_res['msg'])
                            app.processEvents()
                            if guiData['everyCheck']:
                                QMessageBox.information(self, "错误提示", "出错信息：%s" % vf03_res['msg'],
                                                        QMessageBox.Yes)
                        proformaNo = vf03_res['Proforma No.']
                        logMsg['Proforma No.'] = proformaNo
                        self.textBrowser.append("Proforma No.:%s" % proformaNo)
                        app.processEvents()
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
            QMessageBox.information(self, "提示信息", '这单%s的数据或者SAP有问题' % guiData['projectNo'], QMessageBox.Yes)
            return logMsg

    # 获取文件
    def getFile(self, path):
        selectBatchFile = QFileDialog.getOpenFileName(self, '选择ODM导出文件',
                                                      '%s\\%s' % (path, today),
                                                      'files(*.docx;*.xls*;*.csv)')
        fileUrl = selectBatchFile[0]
        return fileUrl

    # SAP数据路径
    def getFileUrl(self):
        fileUrl = MyMainWindow.getFile(self, configContent['SAP_Date_URL'])
        if fileUrl:
            self.lineEdit_6.setText(fileUrl)
            app.processEvents()
        else:
            self.textBrowser.append("请重新选择ODM文件")
            QMessageBox.information(self, "提示信息", "请重新选择ODM文件", QMessageBox.Yes)

    # ODM数据路径
    def getODMDataFileUrl(self):
        fileUrl = MyMainWindow.getFile(self, configContent['SAP_Date_URL'])
        if fileUrl:
            self.lineEdit_7.setText(fileUrl)
            app.processEvents()
        else:
            self.textBrowser_2.append("请重新选择ODM文件")
            QMessageBox.information(self, "提示信息", "请重新选择ODM文件", QMessageBox.Yes)

    # 获取需要Combine文件的路径
    def getCombineFileUrl(self):
        fileUrl = MyMainWindow.getFile(self, configContent['SAP_Date_URL'])
        if fileUrl:
            self.lineEdit_8.setText(fileUrl)
            app.processEvents()
        else:
            self.textBrowser_2.append("请重新选择ODM文件")
            QMessageBox.information(self, "提示信息", "请重新选择ODM文件", QMessageBox.Yes)

    # 获取Log文件路径
    def getLogFileUrl(self):
        fileUrl = MyMainWindow.getFile(self, configContent['SAP_Date_URL'])
        if fileUrl:
            self.lineEdit_9.setText(fileUrl)
            app.processEvents()
        else:
            self.textBrowser_2.append("请重新选择ODM文件")
            QMessageBox.information(self, "提示信息", "请重新选择ODM文件", QMessageBox.Yes)

    # 查看SAP操作数据详情
    def viewOdmData(self):
        fileUrl = self.lineEdit_6.text()
        odm_data_obj = Get_Data()
        df = odm_data_obj.getFileData(fileUrl)
        myTable.createTable(df)
        myTable.showMaximized()

    # 批量开Order
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
                excelFileType = 'xlsx'
                logFileName = 'log'
                logDataPath = MyMainWindow.getFileName(self, logFileUrl, logFileName, excelFileType)

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
                combineKeyFieldsList = excel_field_mapper.update_field_names(combineKeyFieldsList)
                logFile = newData.fileData[combineKeyFieldsList]
                logFile['Order No.'] = ''
                logFile['Remark'] = ''
                logFile['Proforma No.'] = ''
                logFile['sapAmountVat'] = ''
                logFile['Update Time'] = '未开Order'
                if 'Text' not in headList:
                    logFile['Text'] = ''
                if 'Long Text' not in headList:
                    logFile['Long Text'] = ''

                fileDataList = newData.getFileDataList(getFileDataListKey)
                headerData = newData.getHeaderData()
                n = 0
                # 实例化sap
                sap_obj = Sap()
                if sap_obj.res['flag']:
                    for n in range(len(fileDataList['Amount'])):
                        if fileDataList['Material Code'][n] == '':
                            QMessageBox.information(self, "提示信息", "无Material Code，请检查", QMessageBox.Yes)
                            break
                        else:
                            materialCode = fileDataList['Material Code'][n]
                        self.lineEdit_2.setText(str(fileDataList['Project No.'][n]))
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
                                self.lineEdit_4.clear()
                            else:
                                pass
                        else:
                            self.lineEdit_4.clear()
                        app.processEvents()
                        logMsg = MyMainWindow.sapOperate(self, sap_obj)
                        # 写log
                        logIndex = logFile[(logFile['Project No.'] == fileDataList['Project No.'][n])].index.tolist()[0]
                        logFile.loc[logIndex, 'Order No.'] = logMsg['orderNo']
                        logFile.loc[logIndex, 'Remark'] = logMsg['Remark']
                        logFile.loc[logIndex, 'Proforma No.'] = logMsg['Proforma No.']
                        nowDate = datetime.datetime.today()
                        logFile.loc[logIndex, 'Update Time'] = nowDate
                        logDataFile = logFile.to_excel('%s' % logDataPath, merge_cells=False)
                        self.lineEdit_9.setText(logDataPath)
                        if n < len(fileDataList['Amount']) - 1:
                            if self.checkBox_5.isChecked():
                                reply = QMessageBox.question(self, '信息', '是否继续填写下一个Order',
                                                             QMessageBox.Yes | QMessageBox.No,
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
                    sap_obj.end_sap()
                else:
                    # sap实例化失败
                    self.textBrowser.append("SAP系统未启动")
            else:
                self.textBrowser.append("请重新选择ODM文件")
                QMessageBox.information(self, "提示信息", "请重新选择ODM文件", QMessageBox.Yes)
        except Exception as msg:
            fileData = self.lineEdit_6.text()
            self.textBrowser.append('这份%s的ODM获取数据有问题' % fileData)
            self.textBrowser.append('错误信息：%s' % msg)
            self.textBrowser.append('----------------------------------')
            QMessageBox.information(self, "提示信息", '这份%s的ODM获取数据有问题' % fileData, QMessageBox.Yes)

    # 获取文件名称
    def getFileName(self, fileUrl, fileName, fileType):
        nowTime = time.strftime('%Y-%m-%d %H.%M.%S')
        fileName = fileUrl + '/' + nowTime + ' - ' + fileName + '.' + fileType
        return fileName

    # 创建文件夹
    def createFolder(self, url):
        isExists = os.path.exists(url)
        if not isExists:
            os.makedirs(url)

    def lineEditChange(self, url):
        combineKey = self.lineEdit_15.text()
        self.lineEdit_16.setText(combineKey)

    # 获取特殊开票内容
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
            invoiceFile = pd.read_excel(
                '%s/%s' % (configContent['Invoice_File_URL'], configContent['Invoice_File_Name']))
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
            self.textBrowser_2.append('特殊开票文件名称：%s/%s' % (configContent['Invoice_File_URL'], configContent['Invoice_File_Name']))
            self.textBrowser_2.append('----------------------------------')
        except Exception as msg:
            self.textBrowser_2.append('错误信息：%s' % msg)
            self.textBrowser_2.append('----------------------------------')

    # 分开ODM数据
    def splitOdmData(self):
        try:
            fileUrl = self.lineEdit_7.text()
            self.textBrowser_2.append('区分特殊开票的原始文件:%s' % fileUrl)
            (filepath, filename) = os.path.split(fileUrl)
            if fileUrl:
                newData = Get_Data()
                newData.getFileData(fileUrl)
                generalData = newData.fileData[~newData.fileData["Invoices' name (Chinese)"].isin(invoiceName)]
                # 新文件地址
                newFolderUrl = '%s/%s' % (filepath, today)
                newFolder = File_Opetate()
                newFolder.createFolder(newFolderUrl)
                ExcelFileType = 'xlsx'
                if generalData.empty:
                    pass
                else:
                    invoiceFileName = '1.正常合并'
                    invoiceFilePath = newFolder.getFileName(newFolderUrl, invoiceFileName, ExcelFileType)
                    self.textBrowser_2.append('1:%s' % invoiceFilePath)
                    generalFile = generalData.to_excel('%s' % invoiceFilePath, merge_cells=False)
                specialData = newData.fileData[newData.fileData["Invoices' name (Chinese)"].isin(invoiceName)]
                fileNum = 2
                for each in specialInvoiceMsg:
                    eachSpecialData = specialData[
                        specialData["Invoices' name (Chinese)"].isin(specialInvoiceMsg[each]['Invoice name'])]
                    if eachSpecialData.empty:
                        pass
                    else:
                        invoiceFileName = str(fileNum) + '.' + each
                        invoiceFilePath = newFolder.getFileName(newFolderUrl, invoiceFileName, ExcelFileType)
                        self.textBrowser_2.append('%s:%s' % (fileNum, invoiceFilePath))
                        specialFile = eachSpecialData.to_excel('%s' % invoiceFilePath, merge_cells=False)
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

    # 数据添加列信息
    def addColumnMsg(self, data):
        key = self.lineEdit_24.text().split(';')
        key = excel_field_mapper.update_field_names(key)
        df = data
        df['column_msg'] = df[key].apply(lambda row: '\t'.join(map(str, row)), axis=1)
        return df

    # 数据添加行信息
    def addRowMsg(self, data):
        key = self.lineEdit_23.text().split(';')
        key = excel_field_mapper.update_field_names(key)
        df = data
        df['row_msg'] = df[key].apply(lambda row: '\n'.join(f"{col}:{val}" for col, val in zip(df[key], row)),
                                      axis=1)
        return df

    # 合并行数据
    def combineMsg(self, key, data):
        newData = Get_Data()
        combineProject = data.groupby(key).apply(newData.concat_func).reset_index()
        df = pd.merge(data, combineProject, on=key, how='right')
        return df

    # 添加信息
    def addMsg(self):
        fileUrl = self.lineEdit_7.text()
        (filepath, filename) = os.path.split(fileUrl)
        if fileUrl:
            rowChecked = self.checkBox_17.isChecked()
            columnChecked = self.checkBox_18.isChecked()
            if rowChecked or columnChecked:
                # try:
                    column_name = self.comboBox_5.currentText()
                    # 转关键词
                    column_key = self.lineEdit_24.text().split(';')
                    # # 老的数据名称
                    # column_key_list = excel_field_mapper.update_field_names(column_key)
                    # column_key_str = ';'.join(column_key_list)
                    # 新的数据名称
                    column_key_str = ';'.join(column_key)
                    key = column_key_str.replace(';', '\t')
                    newData = Get_Data()
                    newData.getFileData(fileUrl)
                    if rowChecked:
                        newData.fileData = MyMainWindow.addRowMsg(self, newData.fileData)
                    if columnChecked:
                        newData.fileData = MyMainWindow.addColumnMsg(self, newData.fileData)
                        newData.fileData['column_msg'] = key + '\n' + newData.fileData['column_msg']
                    if rowChecked and columnChecked:
                        newData.fileData[column_name] = newData.fileData['row_msg'] + '\n' + newData.fileData[
                            'column_msg']
                    elif rowChecked:
                        newData.fileData[column_name] = newData.fileData['row_msg']
                    else:
                        newData.fileData[column_name] = newData.fileData['column_msg']
                    if self.checkBox_17.isChecked() or self.checkBox_18.isChecked():
                        newData.fileData[column_name] = newData.fileData[self.comboBox_5.currentText()].str.replace('nan', '')
                        newData.fileData[column_name] = newData.fileData[self.comboBox_5.currentText()].str.replace('XXXXXX', '')
                    ExcelFileType = 'xlsx'
                    odmFileName = '添加信息后的数据'
                    odmDataPath = MyMainWindow.getFileName(self, filepath, odmFileName, ExcelFileType)
                    # newData.fileData = newData.fileData.applymap(lambda x: x.strip('"') if isinstance(x, str) else x)
                    odmDataFile = newData.fileData.to_excel('%s' % (odmDataPath), merge_cells=False)
                    self.textBrowser_2.append('已完成该文件的信息添加')
                    self.textBrowser_2.append('文件保存在：%s' % odmDataPath)
                    self.textBrowser_2.append('----------------------------------')
                    app.processEvents()
        #         except Exception as msg:
        #             self.textBrowser_2.append('这份%s的ODM获取数据有问题' % fileUrl)
        #             self.textBrowser_2.append('错误信息：%s' % msg)
        #             self.textBrowser_2.append('----------------------------------')
        #             app.processEvents()
        #     else:
        #         self.textBrowser_2.append('没有选中要添加的信息，请选择')
        #         self.textBrowser_2.append('----------------------------------')
        #         app.processEvents()
        # else:
        #     self.textBrowser_2.append('没有文件，请重新选择')
        #     self.textBrowser_2.append('----------------------------------')
        #     app.processEvents()


    # 数据透视并合并
    def odmCombineData(self):
        try:
            fileUrl = self.lineEdit_7.text()
            self.textBrowser_2.append('合并前的原始数据：%s' % fileUrl)
            (filepath, filename) = os.path.split(fileUrl)
            if fileUrl:
                newData = Get_Data()
                newData.getFileData(fileUrl)
                # 删除Amount为0的数据
                deleteRowList = {'Amount': 0}
                newData.deleteTheRows(deleteRowList)
                newData.fileData.sort_values(
                    by=["Invoices' name (Chinese)", 'CS', 'Sales', 'Currency', 'Material Code', 'Buyer(GPC)', 'Month'],
                    axis=0, ascending=[True, True, True, True, True, True, True], inplace=True)
                # 只保留Order No为空的数据
                newData.fileData = newData.fileData[newData.fileData[['SAP Order No.']].isnull().T.any()]
                # material code将空值填上
                headList = newData.getHeaderData()
                if ("PHY Material Code" in headList) and ("CHM Material Code" in headList):
                    fillNanColumnKey = {'Material Code': ["PHY Material Code", "CHM Material Code"]}
                    newData.fillNanColumn(fillNanColumnKey)
                # 将联系人空值填上
                newData.fileData['Client Contact Name'].fillna("XXXXXX", inplace=True)
                # 单个数据保留原始数据
                if self.checkBox_17.isChecked():
                    newData.fileData = myWin.addRowMsg(newData.fileData)
                if self.checkBox_18.isChecked():
                    newData.fileData = myWin.addColumnMsg(newData.fileData)
                # 保存原始数据
                fileUrl = '%s/%s' % (filepath, today)
                MyMainWindow.createFolder(self, fileUrl)
                ExcelFileType = 'xlsx'
                odmFileName = '1.ODM Raw Data'
                odmDataPath = MyMainWindow.getFileName(self, fileUrl, odmFileName, ExcelFileType)
                odmDataFile = newData.fileData.to_excel('%s' % (odmDataPath), merge_cells=False)
                # 数据透视并保存
                combinekeyFields = self.lineEdit_15.text()
                combineKeyFieldsList = combinekeyFields.split(';')
                # 老的数据名称
                combineKeyFieldsList = excel_field_mapper.update_field_names(combineKeyFieldsList)
                pivotTableKey = combineKeyFieldsList
                # pivotTableKey = ['CS', 'Sales', 'Currency', 'Material Code', "Invoices' name (Chinese)", 'Buyer(GPC)', 'Month', 'Exchange Rate']
                valusKey = ['Amount', 'Amount with VAT', 'Total Cost', 'Revenue\n(RMB)']
                pivotTable = newData.pivotTable(pivotTableKey, valusKey)
                combineFileName = '2.Combine'
                combineFileNamePath = MyMainWindow.getFileName(self, fileUrl, combineFileName, ExcelFileType)
                combineFile = pivotTable.to_excel('%s' % (combineFileNamePath), merge_cells=False)
                # 读取数据透视数据
                combineData = Get_Data()
                combineData = combineData.getFileData(combineFileNamePath)
                # 删除列
                deleteColumnList = ['Amount', 'Amount with VAT', 'Total Cost', 'Revenue\n(RMB)']
                newData = newData.deleteTheColumn(deleteColumnList)
                # 合并多行数据
                if self.checkBox_17.isChecked():
                    combineRrData = newData.groupby(combineKeyFieldsList).apply(Get_Data().row_concat_func).reset_index()
                    newData = pd.merge(newData, combineRrData, on=combineKeyFieldsList, how='right')
                if self.checkBox_18.isChecked():
                    combineRcData = newData.groupby(combineKeyFieldsList).apply(Get_Data().column_concat_func).reset_index()
                    newData = pd.merge(newData, combineRcData, on=combineKeyFieldsList, how='right')
                    # 转关键词
                    column_key = self.lineEdit_24.text().split(';')
                    # # 老的数据名称
                    # column_key_list = excel_field_mapper.update_field_names(column_key)
                    # column_key_str = ';'.join(column_key_list)
                    # 新字段
                    column_key_str = ';'.join(column_key)
                    key = column_key_str.replace(';', '\t')
                    # key = self.lineEdit_24.text().replace(';', '\t')
                    newData['combine_column_msg'] = key + '\n' + newData['combine_column_msg']
                # 合并两列数据
                if self.checkBox_17.isChecked() and self.checkBox_18.isChecked():
                    newData[self.comboBox_5.currentText()] = newData['combine_row_msg'] + '\n' + newData['combine_column_msg']
                elif self.checkBox_17.isChecked():
                    newData[self.comboBox_5.currentText()] = newData['combine_row_msg']
                elif self.checkBox_18.isChecked():
                    newData[self.comboBox_5.currentText()] = newData['combine_column_msg']
                #  替换无用字符
                if self.checkBox_17.isChecked() or self.checkBox_18.isChecked():
                    newData[self.comboBox_5.currentText()] = newData[self.comboBox_5.currentText()].str.replace('nan', '')
                    newData[self.comboBox_5.currentText()] = newData[self.comboBox_5.currentText()].str.replace('XXXXXX', '')
                # merge数据，combine和原始数据
                onData = combineKeyFieldsList
                # onData = ['CS', 'Sales', 'Currency', 'Material Code', "Invoices' name (Chinese)", 'Buyer(GPC)', 'Month', 'Exchange Rate']
                mergeData = pd.merge(combineData, newData, on=onData, how='right')
                mergeDataName = '3.Merge to Project'
                mergeFileNamePath = MyMainWindow.getFileName(self, fileUrl, mergeDataName, ExcelFileType)
                mergeFile = mergeData.to_excel('%s' % (mergeFileNamePath), merge_cells=False)
                self.lineEdit_8.setText(mergeFileNamePath)
                # merge数据去重得到最终数据
                mergeData.drop_duplicates(subset=pivotTableKey, keep='first', inplace=True)
                # mergeData['Project No.'] = mergeData['Project No.'].astype(str)
                # mergeData['Project No.'] = mergeData['Project No.'].apply(lambda x: '{:.8f}'.format(float(x)))
                finalDataName = '4.final'
                finalFileNamePath = MyMainWindow.getFileName(self, fileUrl, finalDataName, ExcelFileType)
                ascendingList = [True] * len(combineKeyFieldsList)
                mergeData.sort_values(by=combineKeyFieldsList, axis=0, ascending=ascendingList, inplace=True)
                # mergeData.sort_values(by=["Invoices' name (Chinese)", 'CS', 'Sales', 'Currency', 'Material Code', 'Buyer(GPC)', 'Month', 'Exchange Rate'], axis=0, ascending=[True, True, True, True, True, True, True, True], inplace=True)
                finalFile = mergeData.to_excel('%s' % (finalFileNamePath), merge_cells=False)
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
                ExcelFileType = 'xlsx'
                fileUrl = combineFilepath
                combineFile = Get_Data()
                combineFile.getFileData(combineFileUrl)
                logFile = Get_Data()
                logFile.getFileData(logFileUrl)
                # merge数据，combine和原始数据
                mergekeyFields = self.lineEdit_16.text()
                mergekeyFieldsList = mergekeyFields.split(';')
                mergekeyFieldsList = excel_field_mapper.update_field_names(mergekeyFieldsList)
                # 原来根据多个字段meger
                # combineFile.fileData['SAP No.'] = combineFile.fileData['SAP No.'].apply(int)
                # logFile['SAP No.'] = logFile['SAP No.'].apply(int)
                delColums = ['Text', 'Long Text', 'Proforma No.']
                for col in delColums:
                    if col in combineFile.fileData.columns:
                        combineFile.fileData = combineFile.fileData.drop(columns=col, axis=1)
                onData = mergekeyFieldsList
                mergeData = pd.merge(combineFile.fileData, logFile.fileData, on=onData, how='outer', indicator=True)
                mergeData.sort_values(by=['Order No.'], axis=0, ascending=[True], inplace=True)
                # 保留数据
                leaveDataList = ["_merge", 'Proforma No.', 'Project No._x', 'Order No.', 'Text', 'Long Text', 'Total Cost_x',
                                 'Revenue\n(RMB)', 'SAP No._x', 'Project No._y', 'Remark', 'Update Time']
                leaveDataList += mergekeyFieldsList
                mergeData = mergeData[leaveDataList]
                ascendingList = [True] * len(leaveDataList)
                mergeData.sort_values(by=leaveDataList, axis=0, ascending=ascendingList, inplace=True)

                mergeDataName = '5.Order Merge Project'
                mergeFileNamePath = MyMainWindow.getFileName(self, fileUrl, mergeDataName, ExcelFileType)
                mergeFile = mergeData.to_excel('%s' % (mergeFileNamePath), merge_cells=False)
                self.textBrowser_2.append('Order NO 与 Project No合并的数据：%s' % mergeFileNamePath)
                self.textBrowser_2.append(
                    'Order Merge Project 数据,根据Order No数据透视算Amount with VAT的平均数值与ODM导出数据算Amount with VAT总值比较大小，有差说明错误。')
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

    # PDF命名规则
    def pdfNameRule(self, msg, flag):
        guiData = MyMainWindow.getAdminGuiData(self)
        if flag == 'invoice':
            pdfName = guiData['pdfName']
        else:
            pdfName = guiData['fapiaoName']
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
        if flag == 'invoice':
            self.lineEdit_17.setText(changedPdfName)
        else:
            self.lineEdit_27.setText(changedPdfName)
        app.processEvents()
        return changedPdfName

    # 获取PDF文件
    def getPdfFiles(self):
        # 获取invoice文件
        selectBatchFile = QFileDialog.getOpenFileNames(self, '选择文件', '%s' % configContent['Invoice_Files_Import_URL'],
                                                       'files(*.pdf)')
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

    # Invoice PDF文件重命名并保存至指定文件夹
    def invoiceRenameOperate(self):
        # invoice中的pdf重新命名
        fileUrls = self.filesUrl
        flag = 'Y'
        if fileUrls == []:
            reply = QMessageBox.question(self, '信息', '没有选中文件，是否重新选择文件', QMessageBox.Yes | QMessageBox.No,
                                         QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                fileUrls = MyMainWindow.getPdfFiles(self)
                if fileUrls == []:
                    flag = 'N'
            else:
                QMessageBox.information(self, "提示信息", "没有选中文件，请重新选择文件", QMessageBox.Yes)
                flag = 'N'
        if flag == 'Y':
            try:
                if self.checkBox_25.isChecked() and self.lineEdit_25.text() == '':
                    self.textBrowser_3.append('未选择Billing List文件')
                    self.textBrowser_3.append('----------------------------------')
                    app.processEvents()
                else:
                    guiData = MyMainWindow.getAdminGuiData(self)
                    pdfOperate = PDF_Operate
                    self.textBrowser_3.append('导出文件夹：%s' % configContent['Invoice_Files_Export_URL'])
                    self.textBrowser_3.append('导出文件名称：')
                    if self.checkBox_25.isChecked():
                        billing_df = myWin.getBillingListData()
                    i = 1
                    # 新增加log
                    log_file_name = 'Invoice %s.xlsx' % time.strftime('%Y-%m-%d %H.%M.%S')
                    Log_file = '%s\\%s' % (configContent['Invoice_Files_Export_URL'], log_file_name)
                    log_obj = Logger(Log_file, ['Update', 'Invoice No', 'File Name', 'Company Name', 'CS', 'Project No', 'Customer Name', 'Client Contact Name', 'Remark'])
                    for fileUrl in fileUrls:
                        try:
                            log_list = {}

                            self.textBrowser_3.append('第%s份文件：' % i)
                            msg = {}
                            msg['Invoice No'] = ''
                            with open(fileUrl, 'rb') as pdfFile:
                                fileCon = pdfOperate.readPdf(pdfFile)
                                fileNum = 0
                                for fileCon[fileNum] in fileCon:
                                    if re.match('.*P. R. China', fileCon[fileNum]) or re.match('.*P.R. China', fileCon[fileNum]) or re.match('Pleasequotethisnumberonallinquiriesandpayments', fileCon[fileNum]) or re.match('请在项目咨询或付款时提示此帐单号', fileCon[fileNum]) or re.match(
                                'Please quote this number on all inquiries and payments.', fileCon[fileNum]):
                                        # DG还是+1，包含invoice no；NB是invoice no后是公司名称
                                        if '487' in str(guiData['invoiceStsrtNum']):
                                            msg['Company Name'] = fileCon[fileNum + 2].replace(
                                                'Please quote this number on all inquiries and payments.', '').replace(
                                                'Invoice No.', '')
                                        else:
                                            # XM+1，不包含invoice no；DG还是+1，包含invoice no；
                                            msg['Company Name'] = fileCon[fileNum + 1].replace(
                                                'Please quote this number on all inquiries and payments.', '').replace(
                                                'Invoice No.', '')
                                    elif re.match('请在项目咨询或付款时提示此帐单号', fileCon[fileNum]):
                                        msg['Company Name'] = fileCon[fileNum + 2].replace(
                                            'Please quote this number on all inquiries and payments.', '').replace(
                                            'Invoice No.', '')
                                    elif re.match(r'%s\d{%s}' % (guiData['invoiceStsrtNum'], int(guiData['invoiceBits']) - len(
                                            str(guiData['invoiceStsrtNum']))),
                                                  fileCon[fileNum]):
                                        msg['Invoice No'] = fileCon[fileNum]
                                    elif re.search(r'\d{2}.\d{3}.\d{2}.\d{4,5}', fileCon[fileNum]):
                                        res = fileCon[fileNum].split(' ')
                                        for each in res:
                                            if re.search(r'\d{2}.\d{3}.\d{2}.\d{4,5}', each):
                                                msg['Project No'] = each
                                            elif re.search(
                                                    r'%s\d{%s}' % (guiData['invoiceStsrtNum'], int(guiData['invoiceBits']) - len(str(guiData['invoiceStsrtNum']))),
                                                    each) and msg['Invoice No'] == '':
                                                msg['Invoice No'] = each
                                    elif re.search(r'%s\d{%s}' % (
                                            guiData['orderStsrtNum'],
                                            int(guiData['orderBits']) - len(str(guiData['orderStsrtNum']))),
                                                   fileCon[fileNum]):
                                        res = fileCon[fileNum].split(' ')
                                        if len(res[1]) == int(guiData['orderBits']):
                                            msg['Order No'] = res[1]
                                    elif 'Client Contact Name' in fileCon[fileNum] or 'ClientContactName' in fileCon[fileNum]:
                                        msg['Client Contact Name'] = fileCon[fileNum].replace('Client Contact Name:', '').replace('ClientContactName:', '').replace('/', '&')
                                    fileNum += 1

                                if 'Client Contact Name' not in msg:
                                    msg['Client Contact Name'] = ''
                                if self.checkBox_25.isChecked():
                                    cs = list(billing_df[billing_df['Final Invoice No.'].astype('int64') == int(msg['Invoice No'])]['CS'])
                                    # NB的Company Name需要在这边操作
                                    customerName = list(
                                        billing_df[
                                            billing_df['Final Invoice No.'].astype('int64') == int(msg['Invoice No'])][
                                            'Customer Name'])
                                    if cs == []:
                                        msg['CS'] = ''
                                        msg['Customer Name'] = ''
                                    else:
                                        msg['CS'] = cs[0]
                                        msg['Customer Name'] = customerName[0]
                                    if 'Company Name' not in msg:
                                        msg['Company Name'] = msg['Customer Name']
                                if msg['Invoice No'] in msg['Company Name']:
                                    msg['Company Name'] = msg['Company Name'].replace(' %s' % msg['Invoice No'], '')
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
                                    elif eachName == 'CS':
                                        outputFlieName += msg['CS']
                                    elif eachName == 'Client Contact Name':
                                        outputFlieName += msg['Client Contact Name']
                                outputFlie = outputFlieName + '.pdf'
                                # 替换非法字符后的文件名字符串
                                outputFlie = PDF_Operate.sanitize_filename(outputFlie)
                                log_list['Invoice No'] = msg['Invoice No']
                                log_list['Company Name'] = msg['Company Name']
                                log_list['Project No'] = msg['Project No']
                                log_list['Client Contact Name'] = msg['Client Contact Name']
                                log_list['File Name'] = outputFlie
                                if 'CS' in msg:
                                    log_list['CS'] = msg['CS']
                                    log_list['Customer Name'] = msg['Customer Name']
                                else:
                                    log_list['CS'] = ''
                                    log_list['Customer Name'] = ''
                                if 'Remark' in msg:
                                    log_list['Remark'] = msg['Remark']
                                else:
                                    log_list['Remark'] = ''
                                log_obj.log(log_list)
                                # outputFlie = msg['Invoice No'] + '-' + msg['Company Name'] + '.pdf'
                                pdfOperate.saveAs(r'%s' % fileUrl, '%s\\%s' % (configContent['Invoice_Files_Export_URL'], outputFlie))
                            self.textBrowser_3.append('%s' % outputFlie)
                            app.processEvents()
                        except Exception as errorMsg:
                            # self.textBrowser_3.append("<font color='red'>第%s份文件：</font>" % i)
                            self.textBrowser_3.append("<font color='red'>出错信息：%s </font>" % errorMsg)
                            self.textBrowser_3.append("<font color='red'>出错的文件：%s </font>" % fileUrl)
                        i += 1
                    log_obj.save_log_to_excel()
                    os.startfile(Log_file)
                    os.startfile(configContent['Invoice_Files_Export_URL'])
                    self.textBrowser_3.append('生成的数据文件：%s' % Log_file)
                    self.textBrowser_3.append('----------------------------------')
            except Exception as errorMsg:
                # self.textBrowser_3.append("<font color='red'>第%s份文件：</font>" % i)
                self.textBrowser_3.append("<font color='red'>出错信息：%s </font>" % errorMsg)
                app.processEvents()

    # 获取电子发票文件
    def getEleInvoiceFiles(self):
        # 获取电子发票文件
        selectBatchFile = QFileDialog.getOpenFileNames(self, '选择文件',
                                                       '%s' % configContent['Ele_Invoice_Files_Import_URL'],
                                                       'files(*.pdf)')
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

    # 电子发票文件重命名并保存至指定文件夹
    def electronicInvoice(self):
        # 需要名称，金额，发票号，税号
        # Excel或csv数据保留对吧billing list的金额
        fileUrls = self.filesUrl
        flag = 'Y'
        if fileUrls == []:
            reply = QMessageBox.question(self, '信息', '没有选中文件，是否重新选择文件', QMessageBox.Yes | QMessageBox.No,
                                         QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                fileUrls = MyMainWindow.getPdfFiles(self)
                if fileUrls == []:
                    flag = 'N'
            else:
                QMessageBox.information(self, "提示信息", "没有选中文件，请重新选择文件", QMessageBox.Yes)
                flag = 'N'
        if flag == 'Y':
            if self.checkBox_25.isChecked() and self.lineEdit_25.text() == '':
                self.textBrowser_3.append('未选择Billing List文件')
                self.textBrowser_3.append('----------------------------------')
                app.processEvents()
            else:
                log_file_name = 'Electronic Invoice %s.xlsx' % time.strftime('%Y-%m-%d %H.%M.%S')
                log_file = '%s\\%s' % (configContent['Ele_Invoice_Files_Export_URL'], log_file_name)
                log_obj = Logger(log_file, [
                    'Update',
                    'id',
                    'Company Name',
                    'Invoice No',
                    'File Name',
                    'Order No',
                    'Revenue',
                    'fapiao',
                    'update',
                    'path',
                    'CS',
                    'ODM Revenue',
                    'ODM Customer Name',
                    'ODMRe - Re',
                    '判断客户名称是否正确'
                ])
                try:
                    pdfOperate = PDF_Operate
                    adminGuiData = MyMainWindow.getAdminGuiData(self)
                    self.textBrowser_3.append('导出文件夹：%s' % configContent['Ele_Invoice_Files_Export_URL'])
                    self.textBrowser_3.append('导出文件名称：')
                    if self.checkBox_25.isChecked():
                        billing_df = myWin.getBillingListData()
                    i = 1

                    for fileUrl in fileUrls:
                        try:
                            self.textBrowser_3.append('第%s份文件：' % i)
                            msg = {}
                            with open(fileUrl, 'rb') as pdfFile:
                                fileCon = pdfOperate.readPdf(pdfFile)
                                fileNum = 0

                                # 获取文件内容
                                for fileCon[fileNum] in fileCon:
                                    if '名称：' in fileCon[fileNum]:
                                        msg['Company Name'] = fileCon[fileNum].split('：')[1].split(' ')[0]
                                    elif '（小写）' in fileCon[fileNum]:
                                        msg['Revenue'] = fileCon[fileNum].split('（小写）')[1]
                                    elif '发票号码：' in fileCon[fileNum]:
                                        msg['fapiao'] = str(fileCon[fileNum].split('：')[1])
                                    elif re.search(r'%s\d{%s}' % (adminGuiData['eleInvoiceStsrtNum'],
                                                                 int(adminGuiData['eleInvoiceBits']) - len(
                                                                     str(adminGuiData['eleInvoiceStsrtNum']))),
                                                   fileCon[fileNum]):
                                        text = fileCon[fileNum].split('/')
                                        try:
                                            msg['Invoice No'] = int(text[-1])
                                        except :
                                            msg['Invoice No'] = re.sub(r'[^0-9]', '', text[-1])
                                    elif re.search(r'%s\d{%s}' % (adminGuiData['eleOrderStsrtNum'],
                                                                 int(adminGuiData['eleOrderBits']) - len(
                                                                     str(adminGuiData['eleOrderStsrtNum']))),
                                                   fileCon[fileNum]):
                                        text = fileCon[fileNum].split('/')
                                        try:
                                            msg['Order No'] = text[1]
                                        except:
                                            msg['Order No'] = ''
                                    fileNum += 1
                                if 'Order No' not in msg:
                                    msg['Order No'] = ''
                                if 'Invoice No' not in msg:
                                    msg['Invoice No'] = re.findall(r'%s\d{%s}' % (adminGuiData['eleInvoiceStsrtNum'],
                                                                                 int(adminGuiData['eleInvoiceBits']) - len(
                                                                                     str(adminGuiData['eleInvoiceStsrtNum']))),
                                                                   fileUrl)[0]
                                if self.checkBox_25.isChecked():
                                    # pandas中有int32和int64的格式区别
                                    cs = list(
                                        billing_df[billing_df['Final Invoice No.'].astype('int64') == int(msg['Invoice No'])]['CS'])
                                    odmRe = list(
                                        billing_df[billing_df['Final Invoice No.'].astype('int64') == int(msg['Invoice No'])]['求和项:Amount with VAT'])
                                    customerName = list(
                                        billing_df[billing_df['Final Invoice No.'].astype('int64') == int(msg['Invoice No'])]['Customer Name'])
                                    if cs == []:
                                        msg['CS'] = ''
                                        msg['odmRe'] = 0.00
                                        msg['Customer Name'] = ''
                                    else:
                                        msg['CS'] = cs[0]
                                        msg['odmRe'] = odmRe[0]
                                        msg['Customer Name'] = customerName[0]
                                else:
                                    msg['CS'] = ''
                                    msg['odmRe'] = 0.00
                                    msg['Customer Name'] = ''
                                # 文件命名规则
                                pdfNameRule = adminGuiData['fapiaoName'].split(' + ')
                                outputFlieName = ''
                                for eachName in pdfNameRule:
                                    if outputFlieName != '':
                                        outputFlieName += '-'
                                    if eachName == 'Invoice No':
                                        outputFlieName += str(msg['Invoice No'])
                                    elif eachName == 'Company Name':
                                        outputFlieName += msg['Company Name']
                                    elif eachName == 'Order No':
                                        outputFlieName += str(msg['Order No'])
                                    elif eachName == 'FaPiao No':
                                        outputFlieName += msg['fapiao']
                                    elif eachName == 'Revenue':
                                        outputFlieName += msg['Revenue']
                                    elif eachName == 'CS':
                                        outputFlieName += msg['CS']

                                # outputFlieName = msg['Company Name'] + '-' + msg['Revenue'] + '-' + msg['fapiao']
                                outputFlie = outputFlieName + '.pdf'
                                exportUrl = configContent['Ele_Invoice_Files_Export_URL']
                                newFolder = File_Opetate()
                                newFolder.createFolder(exportUrl)
                                pdfOperate.saveAs(fileUrl, '%s\\%s' % (exportUrl, outputFlie))
                                log_list = {
                                    'id': i,
                                    'Company Name': msg.get('Company Name', ''),
                                    'Invoice No': msg.get('Invoice No', ''),
                                    'File Name': outputFlie,
                                    'Order No': msg.get('Order No', ''),
                                    'Revenue': float(msg.get('Revenue', '0.0').replace('¥', '')),
                                    'fapiao': msg.get('fapiao', ''),
                                    'update': time.strftime('%Y-%m-%d %H.%M.%S'),
                                    'path': '%s\\%s' % (exportUrl, outputFlie),
                                    'CS': msg.get('CS', ''),
                                    'ODM Revenue': msg.get('odmRe', 0.00),
                                    'ODM Customer Name': msg.get('Customer Name', ''),
                                    'ODMRe - Re': msg.get('odmRe', 0.00) - float(
                                        msg.get('Revenue', '0.0').replace('¥', '')),
                                    '判断客户名称是否正确': msg.get('Company Name', '') == msg.get('Customer Name', '')
                                }
                                log_obj.log(log_list)
                            log_obj.save_log_to_excel()
                            self.textBrowser_3.append('%s' % outputFlie)
                            app.processEvents()

                        except Exception as errorMsg:

                            log_obj.save_log_to_excel()
                            # self.textBrowser_3.append("<font color='red'>第%s份文件：</font>" % i)
                            self.textBrowser_3.append("<font color='red'>出错信息：%s </font>" % errorMsg)
                            self.textBrowser_3.append("<font color='red'>出错的文件：%s </font>" % fileUrl)
                            app.processEvents()
                        i += 1

                    os.startfile(log_file)
                    self.textBrowser_3.append('已生成数据文件：%s' % log_file)
                    self.textBrowser_3.append('----------------------------------')
                    os.startfile(configContent['Ele_Invoice_Files_Export_URL'])
                except Exception as errorMsg:
                    log_obj.save_log_to_excel()
                    os.startfile(log_file)
                    self.textBrowser_3.append("<font color='red'>出错信息：%s </font>" % errorMsg)
                    app.processEvents()

    # Order解锁或关闭操作
    def orderUnlockOrLock(self, flag):
        fileUrl = self.lineEdit_6.text()
        (filepath, filename) = os.path.split(fileUrl)
        if fileUrl:
            log_file_name = 'log %s.xlsx' % time.strftime('%Y-%m-%d %H.%M.%S')
            Log_file = '%s\\%s' % (filepath, log_file_name)
            log_obj = Logger(Log_file, ['Update', 'Order No', 'Type', 'Remark'])
            newData = Get_Data()
            file_data = newData.getFileData(fileUrl)
            order_list = list(file_data['Order No'])
            if not self.checkBox_16.isChecked():
                sap_obj = Sap()
            i = 1
            for orderNo in order_list:
                try:
                    log_list = {}
                    log_list['Order No'] = orderNo
                    log_list['Type'] = flag

                    if self.checkBox_16.isChecked():
                        sap_obj = Sap()
                    sap_obj.open_va02(orderNo)
                    lock_res = sap_obj.unlock_or_lock_order(flag)
                    self.textBrowser.append('%s.Order No: %s' % (i, orderNo))
                    self.textBrowser.append('%s' % lock_res['msg'])
                    app.processEvents()
                    if not sap_obj.res['flag']:
                        log_list['Remark'] = lock_res['msg']
                    else:
                        log_list['Remark'] = ''
                    log_obj.log(log_list)
                    i += 1
                except:
                    self.textBrowser.append("<font color='red'>该Order: %s 有问题</font>" % orderNo)
                    app.processEvents()
            log_obj.save_log_to_excel()
            self.textBrowser.append('%s' % Log_file)
            app.processEvents()
            os.startfile(Log_file)
        else:
            self.textBrowser.append('没有文件请添加')
            app.processEvents()

    # 获取billing list文件
    def getBillingListFile(self):
        try:
            selectBatchFile = QFileDialog.getOpenFileName(self, '选择文件',
                                                          '%s' % configContent['Billing_List_URL'],
                                                          'files(*.xlsx)')
            fileUrl = selectBatchFile[0]
            if fileUrl:
                self.lineEdit_25.setText(fileUrl)
                app.processEvents()
                self.textBrowser_3.append('选中Billing List文件：%s' % fileUrl)
                self.textBrowser_3.append('----------------------------------')
            else:
                self.textBrowser_3.append('无选中文件')
                self.textBrowser_3.append('----------------------------------')
            app.processEvents()
            return fileUrl
        except Exception as errorMsg:
            self.textBrowser_3.append("<font color='red'>出错信息：%s </font>" % errorMsg)
            app.processEvents()
            return

    # 获取billing数据
    def getBillingListData(self, sheet_name=[]):
        try:
            billing_list_url = self.lineEdit_25.text()
            if billing_list_url == '':
                self.textBrowser_3.append('无选中文件')
                self.textBrowser_3.append('----------------------------------')
                app.processEvents()
                return None
            else:
                billing_list_obj = Get_Data()
                billing_list_data = billing_list_obj.getFileMoreSheetData(billing_list_url, sheet_name)
                pivotTableKey = ['Final Invoice No.', 'Customer Name', 'CS', 'Cur.']
                valusKey = ['求和项:Amount with VAT']
                billing_list_data = billing_list_obj.pivotTable(pivotTableKey, valusKey)
                billing_list_data = billing_list_data.reset_index()
                return billing_list_data
        except Exception as errorMsg:
            self.textBrowser_3.append("<font color='red'>出错信息：%s </font>" % errorMsg)
            app.processEvents()
            return


    # 查看Billing list数据
    def viewBillingListData(self):
        fileUrl = self.lineEdit_25.text()
        if fileUrl:
            sheet_name = []
            df = myWin.getBillingListData(sheet_name)
            try:
                myTable.createTable(df)
                myTable.showMaximized()
            except Exception as errorMsg:
                self.textBrowser_3.append('数据有问题%s' % errorMsg)
                self.textBrowser_3.append('----------------------------------')
                app.processEvents()
        else:
            self.textBrowser_3.append('无选中文件')
            self.textBrowser_3.append('----------------------------------')
            app.processEvents()



    def get_hour_file_url(self, position):
        fileUrl = myWin.getFile(configContent['Hour_Files_Import_URL'])
        if fileUrl:
            position.setText(fileUrl)
            app.processEvents()
        else:
            self.textBrowser_2.append("请重新选择ODM文件")
            QMessageBox.information(self, "提示信息", "请重新选择ODM文件", QMessageBox.Yes)

    def get_hour_combine_file(self):
        fileUrl = self.lineEdit_30.text()
        pivot_table_key = self.lineEdit_39.text().split(';')
        if fileUrl and pivot_table_key:
            try:
                self.textBrowser_4.append("数据开始合并")
                app.processEvents()
                newData = Get_Data()
                file_data = newData.getFileTableData(fileUrl)
                # 删除
                deleteRowList = {'Order Number': ''}
                newData.deleteTheRows(deleteRowList)
                # 合并
                valus_key = ['Revenue', 'Total Subcon Cost']
                pivot_table_data = newData.pivotTable(pivot_table_key, valus_key)
                current_time = datetime.datetime.now().strftime('%Y-%m-%d %H.%M.%S')
                pivot_table_data_path = '%s\\%s' % (configContent['Hour_Files_Export_URL'], '1.order data %s.xlsx' % current_time)
                pivot_table_data_file = pivot_table_data.to_excel(pivot_table_data_path, merge_cells=False)
                self.lineEdit_37.setText(pivot_table_data_path)
                self.textBrowser_4.append("合并完成")
                self.textBrowser_4.append("文件路径：%s" % pivot_table_data_path)
            except Exception as errorMsg:
                self.textBrowser_4.append("<font color='red'>出错信息：%s </font>" % errorMsg)
                app.processEvents()
        elif pivot_table_key == []:
            self.textBrowser_4.append("请输入合并的key")
        else:
            self.textBrowser_4.append("请重新选择ODM文件")

    def update_config_content(self, update_data):
        # 创建配置字典的深拷贝以避免污染原始配置
        config_content = copy.deepcopy(configContent)
        config_content.update(update_data)
        return config_content  # 返回修改后的副本
    
    def get_department_hour(self):
        """
        计算部门工时并保存结果
        """
        order_data_path = self.lineEdit_37.text()
        hour_gui_data = myWin.getHourGuiData()
        if order_data_path:
            self.textBrowser_4.append("部门开始计算")
            app.processEvents()
            
            # 更新配置内容
            config_content = self.update_config_content(hour_gui_data)
            
            # 获取订单数据
            order_data_obj = Get_Data()
            order_datas = order_data_obj.getFileTableData(order_data_path)
            
            # 初始化结果DataFrame
            all_results = []
            
            # 调用hour方法处理每个订单
            revenue_allocator_obj = RevenueAllocator()
            for _, order_data in order_datas.iterrows():
                # 将Series转换为字典
                order_dict = order_data.to_dict()
                # 计算部门工时
                order_revenue_data = revenue_allocator_obj.allocate_department_hours(order_dict, config_content)
                all_results.extend(order_revenue_data)
            
            # 创建结果DataFrame
            result_df = pd.DataFrame(all_results)
            
            # 生成输出文件名
            current_time = datetime.datetime.now().strftime('%Y-%m-%d %H.%M.%S')
            dept_hour_path = f"{configContent['Hour_Files_Export_URL']}\\2.dept hour {current_time}.xlsx"
            
            # 保存结果
            result_df.to_excel(dept_hour_path, index=False)
            
            # 更新UI
            self.lineEdit_38.setText(dept_hour_path)
            self.textBrowser_4.append("部门计算完成")
            self.textBrowser_4.append(f"文件路径：{dept_hour_path}")
            app.processEvents()
        else:
            self.textBrowser_4.append("请重新选择合并后的文件")
            app.processEvents()

    def get_person_hour(self):
        """
        分配人员工时并保存结果
        """
        dept_hour_path = self.lineEdit_38.text()
        if dept_hour_path:
            self.textBrowser_4.append("开始分配人员")
            app.processEvents()
            
            # 获取配置数据
            hour_gui_data = myWin.getHourGuiData()
            config_content = self.update_config_content(hour_gui_data)
            
            # 获取参数
            max_hours_per_day = int(config_content['Max_Hour'])
            start_date = datetime.datetime.strptime(config_content['Start_Date'], '%Y.%m.%d').date()
            end_date = datetime.datetime.strptime(config_content['End_Date'], '%Y.%m.%d').date()
            
            # 获取部门工时数据
            dept_hour_obj = Get_Data()
            dept_hour_datas = dept_hour_obj.getFileTableData(dept_hour_path)
            
            # 初始化结果列表
            all_results = []
            
            # 处理每个部门工时记录
            revenue_allocator_obj = RevenueAllocator()
            for _, dept_hour in dept_hour_datas.iterrows():
                # 将Series转换为字典
                dept_hour_dict = dept_hour.to_dict()
                # 将单个记录转换为列表形式
                dept_hour_list = [dept_hour_dict]
                self.textBrowser_4.append(f"处理Order Number：{dept_hour_dict['order_no']}")
                app.processEvents()
                # 分配人员工时
                person_hour_data = revenue_allocator_obj.allocate_person_hours(
                    dept_hour_list, 
                    max_hours_per_day, 
                    start_date, 
                    end_date, 
                    staff_dict, 
                    config_content
                )
                all_results.extend(person_hour_data)
            
            # 创建结果DataFrame
            result_df = pd.DataFrame(all_results)
            
            # 生成输出文件名
            current_time = datetime.datetime.now().strftime('%Y-%m-%d %H.%M.%S')
            person_hour_path = f"{configContent['Hour_Files_Export_URL']}\\3.person hour {current_time}.xlsx"
            
            # 保存结果
            result_df.to_excel(person_hour_path, index=False)
            
            # 打开结果文件
            os.startfile(person_hour_path)
            
            # 更新UI
            self.lineEdit_31.setText(person_hour_path)
            self.textBrowser_4.append("分配人员完成")
            self.textBrowser_4.append(f"文件路径：{person_hour_path}")
            app.processEvents()
        else:
            self.textBrowser_4.append("请先完成部门工时计算")
            app.processEvents()




if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QApplication(sys.argv)
    myWin = MyMainWindow()
    myTable = MyTableWindow()
    myWin.show()
    myWin.getConfig()
    sys.exit(app.exec_())
