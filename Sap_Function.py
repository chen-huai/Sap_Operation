import sys, win32com.client, time, datetime
from PyQt5.QtWidgets import QApplication, QMainWindow
from Sap_Operate import *


# from PyQt5 import QtCore, QtGui, QtWidgets
# from PyQt5.QtWidgets import Qself.application, QMainWindow
# from PyQt5.QtWidgets import *
# from PyQt5.QtCore import *

# TODO self.logMsg每次都应该初始化
class Sap(MyMainWindow):
    def __init__(self):
        self.res = {}
        self.res['flag'] = 1
        self.res['msg'] = '可以操作SAP'
        self.logMsg = {}
        self.logMsg['Remark'] = ''
        self.logMsg['orderNo'] = ''
        self.logMsg['Proforma No.'] = ''
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


    # TODO csCode和salesCode需要添加进guiData中
    def va01_operate(self, guiData, revenueData):
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

    def lab_cost(self, guiData, revenueData):
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

    def save_sap(self):
        # 保存操作
        try:
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
                myWin.textBrowser.append("跳过保存")
            else:
                pass
        else:
            pass

    def va02_operate(self, guiData, revenueData):

        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NVA02"
        self.session.findById("wnd[0]").sendVKey(0)
        orderNo = self.session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text
        myWin.textBrowser.append("Order No.:%s" % orderNo)
        app.processEvents()
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
            revenueData['sapAmountVat'] = sapAmountVat

            # 是否需要填写plan cost
            Sap.plan_cost()
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
            Sap.plan_cost()

        if guiData['longText'] != '':
            if myWin.checkBox_8.isChecked() or revenueData['revenueForCny'] >= 35000:
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

        if myWin.checkBox_8.isChecked() or revenueData['revenueForCny'] >= 35000:
            pass
        else:
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
        amountVatStr = re.sub(r"(\d)(?=(\d\d\d)+(?!\d))", r"\1,", format(guiData['amountVat'], '.2f'))
        myWin.textBrowser.append("Sap Amount Vat:%s" % sapAmountVat)
        myWin.textBrowser.append("Amount Vat:%s" % amountVatStr)
        app.processEvents()
        # sapAmountVat在A2是数字，其它为字符串
        if sapAmountVat.strip() != amountVatStr:
            flag = 2
            reply = QMessageBox.question(self, '信息', 'SAP数据与ODM不一致，请确认并修改后再继续！！！',
                                         QMessageBox.Yes | QMessageBox.No,
                                         QMessageBox.Yes)
            self.logMsg['Remark'] = 'SAP数据与ODM不一致，请确认并修改后再继续！！！'
            if reply == QMessageBox.Yes:
                flag = 1
        if (myWin.checkBox_3.isChecked() or myWin.checkBox_6.isChecked()) and flag == 1:
            Sap.save_sap()

    def plan_cost(self, guiData, revenueData):
        if myWin.checkBox_8.isChecked() or revenueData['revenueForCny'] >= 35000:
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
                    if myWin.checkBox_13.isChecked():
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
                    if myWin.checkBox_15.isChecked():
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
                    if myWin.checkBox_13.isChecked():
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
                    if myWin.checkBox_14.isChecked():
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
                        if myWin.checkBox_14.isChecked():
                            n = 2
                        else:
                            n = 1
                        self.session.findById(
                            "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,%s]" % n).text = "E"
                        self.session.findById(
                            "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,%s]" % n).text = guiData['csCostCenter']
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
                    if myWin.checkBox_13.isChecked():
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

                    if myWin.checkBox_14.isChecked() or myWin.checkBox_15.isChecked():
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
                        if myWin.checkBox_14.isChecked():
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,1]").text = guiData['chmCostCenter']
                    else:
                        if myWin.checkBox_15.isChecked():
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,1]").text = guiData['phyCostCenter']

                    if guiData['cost'] > 0:
                        if myWin.checkBox_14.isChecked() or myWin.checkBox_15.isChecked():
                            n = 2
                        else:
                            n = 1
                        self.session.findById(
                            "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,%s]" % n).text = "E"
                        self.session.findById(
                            "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,%s]" % n).text = guiData['csCostCenter']
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



if __name__ == "__main__":
    revenue = 230
    sap_obj = Sap()
    if sap_obj.flag != 0:
        sap_obj.Sap_Operate(1)
