import sys, win32com.client, time, datetime, re
from PyQt5.QtWidgets import QApplication, QMainWindow

from Sap_Operate import *


# from PyQt5 import QtCore, QtGui, QtWidgets
# from PyQt5.QtWidgets import Qself.application, QMainWindow
# from PyQt5.QtWidgets import *
# from PyQt5.QtCore import *

class Sap():
    def __init__(self):
        self.res = {}
        self.res['flag'] = 1
        self.res['msg'] = '可以操作SAP'
        self.logMsg = {}
        self.logMsg['Remark'] = ''
        self.logMsg['orderNo'] = ''
        self.logMsg['Proforma No.'] = ''
        self.logMsg['sapAmountVat'] = ''
        self.today = time.strftime('%Y.%m.%d')
        self.oneWeekday = (datetime.datetime.now() + datetime.timedelta(days=7)).strftime('%Y.%m.%d')
        try:
            self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
            if not type(self.SapGuiAuto) == win32com.client.CDispatch:
                return

            self.application = self.SapGuiAuto.GetScriptingEngine
            if not type(self.application) == win32com.client.CDispatch:
                self.SapGuiAuto = None
                return

            self.connection = self.application.Children(0)
            if not type(self.connection) == win32com.client.CDispatch:
                self.application = None
                self.SapGuiAuto = None
                return

            self.session = self.connection.Children(0)
            if not type(self.session) == win32com.client.CDispatch:
                self.connection = None
                self.application = None
                self.SapGuiAuto = None
                return
        except:
            self.res['flag'] = 0
            self.res['msg'] = '可以操作SAP'
            print('SAP未启动')

    # 初始化数据
    def initializationLogMsg(self):
        self.logMsg = {}
        self.logMsg['Remark'] = ''
        self.logMsg['orderNo'] = ''
        self.logMsg['Proforma No.'] = ''
        self.logMsg['sapAmountVat'] = ''

    def initializationMsg(self):
        self.res = {}
        self.res['flag'] = 1
        self.res['msg'] = '可以操作SAP'

    # TODO csCode和salesCode需要添加进guiData中
    # 创建order
    def va01_operate(self, guiData, revenueData):
        try:
            # 初始化数据
            Sap.initializationLogMsg()
            # 相当于VA01操作
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nva01"
            # 回车键功能
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/ctxtVBAK-AUART").text = guiData['orderType']
            self.session.findById("wnd[0]/usr/ctxtVBAK-VKORG").text = guiData['salesOrganization']
            self.session.findById("wnd[0]/usr/ctxtVBAK-VTWEG").text = guiData['distributionChannels']
            self.session.findById("wnd[0]/usr/ctxtVBAK-VKBUR").text = guiData['salesOffice']
            self.session.findById("wnd[0]/usr/ctxtVBAK-VKGRP").text = guiData['salesGroup']
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById(
                "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").text = \
                guiData['sapNo']
            self.session.findById(
                "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").caretPosition = 6
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").text = guiData[
                'projectNo']
            self.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBKD-BSTDK").text = self.today
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/ctxtVBKD-FBUDA").text = self.today
            self.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").setFocus()
            self.session.findById(
                "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").caretPosition = 17
            self.session.findById("wnd[0]").sendVKey(0)
            # 售达方按钮
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press()
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-WAERK").text = \
                guiData['currencyType']

            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-WAERK").setFocus()
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-WAERK").caretPosition = 3
            self.session.findById("wnd[0]").sendVKey(0)
            try:
                self.session.findById("wnd[1]").sendVKey(0)
            except:
                pass
            else:
                pass
            if guiData['currencyType'] != "CNY":
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBKD-KURSK").text = \
                    guiData['exchangeRate']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBKD-KURSK").setFocus()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBKD-KURSK").caretPosition = 8
                self.session.findById("wnd[0]").sendVKey(0)
            # 会计
            self.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\06").select()
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\06/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").text = "*"
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\06/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").setFocus()
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\06/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").caretPosition = 1
            self.session.findById("wnd[0]").sendVKey(0)
            # 合作伙伴
            self.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09").select()

            # 获取文本名称
            fourName = self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,4]").text
            fiveName = self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,5]").text

            # # eNum负责雇员位置，gNum送达方位置
            if fourName == '负责雇员' or fourName == 'Employee respons.':
                eNum = 4
                gNum = 5
            else:
                eNum = 5
                gNum = 4
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,%s]" % gNum).key = "ZG"
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,%s]" % gNum).text = \
                guiData['globalPartnerCode']
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,%s]" % gNum).setFocus()
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,%s]" % gNum).caretPosition = 8
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,%s]" % eNum).text = \
                guiData['csCode']
            self.session.findById("wnd[0]").sendVKey(0)

            # 联系人
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,6]").key = "AP"
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,6]").setFocus()
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,6]").caretPosition = 0
            self.session.findById("wnd[0]").sendVKey(4)
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[0]").sendVKey(0)
            if guiData['salesName'] != '':
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,7]").key = "VE"
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,7]").text = \
                    guiData['salesCode']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,7]").setFocus()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,7]").caretPosition = 4
                self.session.findById("wnd[0]").sendVKey(0)

            # 文本
            self.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10").select()
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = \
                guiData['shortText']
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes(
                11, 11)
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cmbLV70T-SPRAS").key = "EN"
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cmbLV70T-SPRAS").setFocus()
            self.session.findById("wnd[0]").sendVKey(0)

            # DATA A
            self.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13").select()
            if 'D2' in guiData['materialCode'] or 'D3' in guiData['materialCode']:
                if guiData['sapNo'] in guiData['dataAE1']:
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR1").key = "E1"
                else:
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR1").key = "Z0"
            elif guiData['sapNo'] in guiData['dataAZ2']:
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR1").key = "Z2"
            else:
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR1").key = "00"

            # DATA B
            self.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14").select()
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/ctxtVBAK-ZZAUART").text = "WO"
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/ctxtVBAK-ZZUNLIMITLIAB").text = "N"
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/ctxtZAUFTD-VORAUS_AUFENDE").text = self.oneWeekday
            if revenueData['revenueForCny'] >= 35000:
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/txtZAUFTD-AUFTRAGSWERT").text = format(
                    revenueData['revenueForCny'], '.2f')
        except:
            self.res['flag'] = 0
            self.res['msg'] = 'Order No未创建成功'
            # myWin.textBrowser.append("Order No未创建成功")

    # 填写Data B
    def lab_cost(self, guiData, revenueData):
        try:
            # 初始化数据
            Sap.initializationMsg()
            # Data B是否要添加cost
            # revenuedata包含revenue,planCost,revenueForCny,chmCost,phyCost,chmRe,phyRe,chmCsCostAccounting,chmLabCostAccounting,phyCsCostAccounting
            if 'A2' in guiData['materialCode'] or 'D2' in guiData['materialCode'] or 'D3' in guiData['materialCode']:
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,0]").text = \
                    guiData['chmCostCenter']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,1]").text = \
                    guiData['phyCostCenter']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/ctxtTABD-KOSTL[0,0]").text = \
                    guiData['chmCostCenter']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/ctxtTABD-KOSTL[0,1]").text = \
                    guiData['phyCostCenter']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,0]").text = \
                    revenueData['chmCost']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,1]").text = \
                    revenueData['phyCost']
            elif 'T20' in guiData['materialCode']:
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,0]").text = \
                    guiData['phyCostCenter']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/ctxtTABD-KOSTL[0,0]").text = \
                    guiData['phyCostCenter']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,0]").text = \
                    revenueData['phyCost']
            else:
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,0]").text = \
                    guiData['chmCostCenter']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/ctxtTABD-KOSTL[0,0]").text = \
                    guiData['chmCostCenter']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,0]").text = \
                    revenueData['chmCost']
        except:
            self.res['flag'] = 0
            self.res['msg'] = 'Data B未填写'
            # myWin.textBrowser.append("Data B未填写")

    # 保存
    def save_sap(self):
        # 保存操作
        try:
            # 初始化数据
            Sap.initializationMsg()
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        except:
            try:
                self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
            except:
                self.res['flag'] = 0
                self.res['msg'] = 'Order添加Item失败'
                # myWin.textBrowser.append("Order添加Item失败")
            else:
                pass
        else:
            pass

    # TODO 添加guiData['planCostCheck']，guiData['saveCheck']，guiData['vf01Check']，guiData['vf02Check']
    # 添加item
    def va02_operate(self, guiData, revenueData):
        try:
            # 初始化数据
            Sap.initializationMsg()
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NVA02"
            self.session.findById("wnd[0]").sendVKey(0)
            orderNo = self.session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text
            # myWin.textBrowser.append("Order No.:%s" % orderNo)
            # app.processEvents()
            self.logMsg['orderNo'] = orderNo
            self.session.findById("wnd[0]").sendVKey(0)
            if 'A2' in guiData['materialCode']:
                if '405' in guiData['materialCode']:
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").text = "T75-405-00"
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,1]").text = "T20-405-00"
                else:
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").text = "T75-441-00"
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,1]").text = "T20-441-00"
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/txtVBAP-ZMENG[2,0]").text = "1"
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/txtVBAP-ZMENG[2,1]").text = "1"
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtVBAP-ZIEME[3,0]").text = "pu"
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtVBAP-ZIEME[3,1]").text = "pu"
                self.session.findById("wnd[0]").sendVKey(0)

                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/txtVBAP-ZMENG[2,1]").setFocus()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/txtVBAP-ZMENG[2,1]").caretPosition = 16
                self.session.findById("wnd[0]").sendVKey(2)
                self.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06").select()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text = \
                    revenueData['phyRe']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").setFocus()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").caretPosition = 16
                self.session.findById("wnd[0]").sendVKey(0)
                sapAmountVatStr = self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text
                sapAmountVat = float(sapAmountVatStr.replace(',', ''))

                self.session.findById("wnd[0]/tbar[0]/btn[3]").press()

                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").setFocus()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").caretPosition = 10
                self.session.findById("wnd[0]").sendVKey(2)
                self.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06").select()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text = \
                    revenueData['chmRe']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").setFocus()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").caretPosition = 16
                self.session.findById("wnd[0]").sendVKey(0)
                sapAmountVatStr = self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text

                sapAmountVat += float(sapAmountVatStr.replace(',', ''))
                sapAmountVat = format(sapAmountVat, '.2f')
                sapAmountVat = re.sub(r"(\d)(?=(\d\d\d)+(?!\d))", r"\1,", sapAmountVat)

                # 是否需要填写plan cost
                Sap.plan_cost(guiData, revenueData)
            else:
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").text = \
                    guiData['materialCode']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/txtVBAP-ZMENG[2,0]").text = "1"
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtVBAP-ZIEME[3,0]").text = "pu"
                self.session.findById("wnd[0]").sendVKey(0)
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").setFocus()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").caretPosition = 10
                self.session.findById("wnd[0]").sendVKey(2)
                self.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06").select()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text = format(
                    revenue, '.2f')
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").setFocus()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").caretPosition = 16
                self.session.findById("wnd[0]").sendVKey(0)
                sapAmountVat = self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text

                # 是否需要填写plan cost
                Sap.plan_cost(guiData, revenueData)

            if guiData['longText'] != '':
                # if myWin.checkBox_8.isChecked() or revenueData['revenueForCny'] >= 35000:
                if guiData['planCostCheck'] or revenueData['revenueForCny'] >= 35000:
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").setFocus()
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").caretPosition = 10
                    self.session.findById("wnd[0]").sendVKey(2)

                self.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09").select()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = \
                    guiData['longText']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes(
                    4, 4)
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cmbLV70T-SPRAS").key = "EN"
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cmbLV70T-SPRAS").setFocus()
                self.session.findById("wnd[0]").sendVKey(0)

            if guiData['planCostCheck'] or revenueData['revenueForCny'] >= 35000:
                pass
            else:
                self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self.logMsg['sapAmountVat'] = sapAmountVat
            #    外面整
            # amountVatStr = re.sub(r"(\d)(?=(\d\d\d)+(?!\d))", r"\1,", format(guiData['amountVat'], '.2f'))
            # myWin.textBrowser.append("Sap Amount Vat:%s" % sapAmountVat)
            # myWin.textBrowser.append("Amount Vat:%s" % amountVatStr)
            # app.processEvents()
            # # sapAmountVat在A2是数字，其它为字符串，外面判断
            # if sapAmountVat.strip() != amountVatStr:
            #     flag = 2
            #     reply = QMessageBox.question(self, '信息', 'SAP数据与ODM不一致，请确认并修改后再继续！！！',
            #                                  QMessageBox.Yes | QMessageBox.No,
            #                                  QMessageBox.Yes)
            #     self.logMsg['Remark'] = 'SAP数据与ODM不一致，请确认并修改后再继续！！！'
            #     if reply == QMessageBox.Yes:
            #         flag = 1
            # if (guiData['saveCheck'] or myWin.checkBox_6.isChecked()) and flag == 1:
            #     Sap.save_sap()
        except:
            self.res['flag'] = 0
            self.res['msg'] = 'Order添加Item失败'
            # myWin.textBrowser.append("编辑order失败")

    # 填写plan cost
    def plan_cost(self, guiData, revenueData):
        try:
            if guiData['planCostCheck'] or revenueData['revenueForCny'] >= 35000:
                if 'A2' in guiData['materialCode']:
                    self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                    if revenueData['revenueForCny'] >= 1000:
                        # 这个是Item2000的
                        self.session.findById(
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,1]").setFocus()
                        self.session.findById(
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,1]").caretPosition = 10
                        self.session.findById("wnd[0]/mbar/menu[3]/menu[7]").select()
                        self.session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        # cs
                        if guiData['csCheck']:
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,0]").text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,0]").text = guiData['csCostCenter']
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,0]").text = "T01AST"
                            # 录金额
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").text = round(
                                float(revenueData['phyCsCostAccounting']), 0)
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").caretPosition = 20
                            self.session.findById("wnd[0]").sendVKey(0)
                        # phy
                        if guiData['phyCheck']:
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,1]").text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,1]").text = guiData['phyCostCenter']
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,1]").text = "T01AST"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").text = round(
                                float(revenueData['phyLabCostAccounting']), 0)
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").caretPosition = 20
                        self.session.findById("wnd[0]").sendVKey(0)

                        # self.session.findById("wnd[0]").sendVKey(0)
                        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()

                        # self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
                        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                        self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()

                        # Items1000的plan cost
                        self.session.findById(
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").setFocus()
                        self.session.findById(
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").caretPosition = 10
                        self.session.findById("wnd[0]/mbar/menu[3]/menu[7]").select()
                        self.session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        # cs
                        if guiData['csCheck']:
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,0]").text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,0]").text = guiData['csCostCenter']
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,0]").text = "T01AST"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").text = round(
                                float(revenueData['chmCsCostAccounting']), 0)
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").caretPosition = 19
                        # 	chm
                        if guiData['chmCheck']:
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,1]").text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,1]").text = guiData['chmCostCenter']
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,1]").text = "T01AST"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").text = round(
                                float(revenueData['chmLabCostAccounting']), 0)
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").caretPosition = 20
                        self.session.findById("wnd[0]").sendVKey(0)
                        if guiData['cost'] > 0:
                            if guiData['chmCheck']:
                                n = 2
                            else:
                                n = 1
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,%s]" % n).text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,%s]" % n).text = guiData[
                                'csCostCenter']
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,%s]" % n).text = "FREMDL"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).text = format(
                                guiData['cost'] / 1.06, '.2f')
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).caretPosition = 20
                            self.session.findById("wnd[0]").sendVKey(0)

                        # self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
                        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                        self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                    # self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                else:
                    self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                    if revenueData['revenueForCny'] >= 1000:
                        self.session.findById(
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").setFocus()
                        self.session.findById(
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").caretPosition = 10
                        self.session.findById("wnd[0]/mbar/menu[3]/menu[7]").select()
                        self.session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        if guiData['csCheck']:
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,0]").text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,0]").text = guiData['csCostCenter']
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,0]").text = "T01AST"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").text = round(
                                float(revenueData['csCostAccounting']), 0)
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").caretPosition = 19

                        if guiData['chmCheck'] or guiData['phyCheck']:
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,1]").text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,1]").text = "T01AST"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").text = round(
                                float(revenueData['labCostAccounting']), 0)
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").caretPosition = 20

                        if 'T75' in guiData['materialCode']:
                            if guiData['chmCheck']:
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,1]").text = guiData[
                                    'chmCostCenter']
                        else:
                            if guiData['phyCheck']:
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,1]").text = guiData[
                                    'phyCostCenter']

                        if guiData['cost'] > 0:
                            if guiData['chmCheck'] or guiData['phyCheck']:
                                n = 2
                            else:
                                n = 1
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,%s]" % n).text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,%s]" % n).text = guiData[
                                'csCostCenter']
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,%s]" % n).text = "FREMDL"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).text = format(
                                guiData['cost'] / 1.06, '.2f')
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).caretPosition = 20
                            self.session.findById("wnd[0]").sendVKey(0)
                        # 直接保存
                        # self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
                        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                        self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        except:
            self.res['flag'] = 0
            self.res['msg'] += 'plan cost未添加成功'

    def vf01_operate(self):
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nvf01"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/tbar[0]/btn[11]").press()

    def vf03_operate(self):
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nvf03"
        self.session.findById("wnd[0]").sendVKey(0)
        proformaNo = self.session.findById("wnd[0]/usr/ctxtVBRK-VBELN").text
        self.logMsg['Proforma No.'] = proformaNo
        self.session.findById("wnd[0]/mbar/menu[0]/menu[11]").select()
        self.session.findById("wnd[1]/tbar[0]/btn[37]").press()

    # 打开order
    def open_va02(self, guiData, revenueData, orderNo):
        try:
            # 初始化数据
            Sap.initializationMsg()
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NVA02"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = orderNo
            self.session.findById("wnd[0]").sendVKey(0)
        except:
            self.res['flag'] = 0
            self.res['msg'] = "该Order No %s 未开启" % orderNo
            # myWin.textBrowser.append("该Order No %s 未开启" % orderNo)

    # 结束sap
    def end_sap(self):
        self.session = None
        self.connection = None
        self.application = None
        self.SapGuiAuto = None


# if __name__ == "__main__":
#     revenue = 230
#     guiData = {}
#     guiData['sapNo'] = 5010920197
#     guiData['projectNo'] = '66.405.23.7556.02A'
#     guiData['materialCode'] = 'T75-405-A2'
#     guiData['currencyType'] = 'CNY'
#     guiData['exchangeRate'] = float(1)
#     guiData['globalPartnerCode'] = 1500155
#     guiData['csName'] = 'cai, barry'
#     guiData['salesName'] = ''
#     guiData['amount'] = float(200)
#     guiData['cost'] = float(0)
#     guiData['amountVat'] = float(212)
#     guiData['csHourlyRate'] = float(300)
#     guiData['chmHourlyRate'] = float(250)
#     guiData['phyHourlyRate'] = float(280)
#     guiData['longText'] = ''
#     guiData['shortText'] = 'TEST'
#     guiData['planCostRate'] = float(0.9)
#     guiData['significantDigits'] = int(0)
#     guiData['chmCostRate'] = float(0.3)
#     guiData['phyCostRate'] = float(0.3)
#     guiData['dataAE1'] = ''
#     guiData['dataAZ2'] = ''
#     guiData['orderType'] = 'DR'
#     guiData['salesOrganization'] = '0486'
#     guiData['distributionChannels'] = '01'
#     guiData['salesOffice'] = '>601'
#     guiData['salesGroup'] = '240'
#     guiData['csCostCenter'] = '48601240'
#     guiData['chmCostCenter'] = '48601293'
#     guiData['phyCostCenter'] = '48601294'
#     guiData['csCode'] = '6375108'
#     guiData['salesCode'] = ''
#     my_w = MyMainWindow()
#     revenueData = my_w.getRevenueData(guiData)
#     sap_obj = Sap()
#     if sap_obj.flag != 0:
#         sap_obj.va01_operate(guiData, revenueData)
