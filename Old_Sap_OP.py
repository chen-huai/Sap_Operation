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
        guiData = MyMainWindow.getGuiData(self)
        if guiData['csName'] != '':
            csCode = configContent[guiData['csName']]
        # guiData['salesName'] = self.comboBox_3.currentText()
        if guiData['salesName'] != '':
            salesCode = configContent[guiData['salesName']]
        orderNo = ''
        proformaNo = ''
        if guiData['sapNo'] == '' or guiData['projectNo'] == '' or guiData['materialCode'] == '' or guiData[
            'currencyType'] == '' or guiData['exchangeRate'] == '' or guiData['globalPartnerCode'] == '' or guiData[
            'csName'] == '' or guiData['amount'] == 0.00 or guiData['amountVat'] == 0.00:
            self.textBrowser.append("有关键信息未填")
            logMsg['Remark'] = '有关键信息未填'
            self.textBrowser.append(
                "'Project No.', 'CS', 'Sales', 'Currency', 'GPC Glo. Par. Code', 'Material Code','SAP No.', 'Amount', 'Amount with VAT', 'Exchange Rate'都是必须填写的")
            self.textBrowser.append('----------------------------------')
            app.processEvents()
            QMessageBox.information(self, "提示信息", "有关键信息未填", QMessageBox.Yes)

        else:
            revenue = guiData['amountVat'] / 1.06
            # plan cost
            # planCost = revenue * guiData['exchangeRate'] * 0.9 - guiData['cost']
            planCost = revenue * guiData['exchangeRate']
            revenueForCny = revenue * guiData['exchangeRate']
            if ('405' in guiData['materialCode']) and (
                    ("A2" in guiData['materialCode']) or ("D2" in guiData['materialCode']) or (
                    "D3" in guiData['materialCode'])):
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
                chmCsCostAccounting = format((revenueForCny * guiData['planCostRate'] - guiData['cost']) * 0.5 * (
                        1 - guiData['chmCostRate']) / guiData['csHourlyRate'],
                                             '.%sf' % guiData['significantDigits'])
                # CHM的Item1000-Cost
                chmLabCostAccounting = format(
                    (revenueForCny * guiData['planCostRate'] - guiData['cost']) * 0.5 * guiData['chmCostRate'] /
                    guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])
                # CS的Item2000-Cost
                phyCsCostAccounting = format((revenueForCny * guiData['planCostRate'] - guiData['cost']) * 0.5 * (
                        1 - guiData['phyCostRate']) / guiData['csHourlyRate'],
                                             '.%sf' % guiData['significantDigits'])
                # PHY的Item2000-Cost
                phyLabCostAccounting = format(
                    (revenueForCny * guiData['planCostRate'] - guiData['cost']) * 0.5 * guiData['phyCostRate'] /
                    guiData['phyHourlyRate'], '.%sf' % guiData['significantDigits'])
            elif ('441' in guiData['materialCode']) and ((
                    "A2" in guiData['materialCode'] or ("D2" in guiData['materialCode']) or (
                    "D3" in guiData['materialCode']))):
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
                chmCsCostAccounting = format((revenueForCny * guiData['planCostRate'] - guiData['cost']) * 0.8 * (
                        1 - guiData['chmCostRate']) / guiData['csHourlyRate'],
                                             '.%sf' % guiData['significantDigits'])
                # CHM的Item1000-Cost
                chmLabCostAccounting = format(
                    (revenueForCny * guiData['planCostRate'] - guiData['cost']) * 0.8 * guiData['chmCostRate'] /
                    guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])
                # CS的Item2000-Cost
                phyCsCostAccounting = format((revenueForCny * guiData['planCostRate'] - guiData['cost']) * 0.2 * (
                        1 - guiData['phyCostRate']) / guiData['csHourlyRate'],
                                             '.%sf' % guiData['significantDigits'])
                # PHY的Item2000-Cost
                phyLabCostAccounting = format(
                    (revenueForCny * guiData['planCostRate'] - guiData['cost']) * 0.2 * guiData['phyCostRate'] /
                    guiData['phyHourlyRate'], '.%sf' % guiData['significantDigits'])
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

                csCostAccounting = format(
                    (revenueForCny * guiData['planCostRate'] - guiData['cost']) * (1 - labCostRate) / guiData[
                        'csHourlyRate'], '.%sf' % guiData['significantDigits'])
                labCostAccounting = format(
                    (revenueForCny * guiData['planCostRate'] - guiData['cost']) * labCostRate / labHourlyRate,
                    '.%sf' % guiData['significantDigits'])

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
                        "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").text = \
                        guiData['sapNo']
                    session.findById(
                        "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").caretPosition = 6
                    session.findById("wnd[0]").sendVKey(0)
                    session.findById("wnd[0]").sendVKey(0)
                    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").text = guiData[
                        'projectNo']
                    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBKD-BSTDK").text = today
                    session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/ctxtVBKD-FBUDA").text = today
                    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").setFocus()
                    session.findById(
                        "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").caretPosition = 17
                    session.findById("wnd[0]").sendVKey(0)
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press()
                    session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-WAERK").text = \
                        guiData['currencyType']

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
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBKD-KURSK").text = \
                            guiData['exchangeRate']
                        session.findById(
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBKD-KURSK").setFocus()
                        session.findById(
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBKD-KURSK").caretPosition = 8
                        session.findById("wnd[0]").sendVKey(0)
                    # 会计
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\06").select()
                    session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\06/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").text = "*"
                    session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\06/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").setFocus()
                    session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\06/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").caretPosition = 1
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
                    session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,%s]" % gNum).text = \
                        guiData['globalPartnerCode']
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
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = \
                        guiData['shortText']
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
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/txtZAUFTD-AUFTRAGSWERT").text = format(
                            revenueForCny, '.2f')
                    # 是否要添加cost
                    if self.checkBox_7.isChecked():
                        if 'A2' in guiData['materialCode'] or 'D2' in guiData['materialCode'] or 'D3' in guiData[
                            'materialCode']:
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
                        sapAmountVat = format(sapAmountVat, '.2f')
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
                                    session.findById(
                                        "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,1]").text = "E"
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
                                    session.findById(
                                        "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,0]").text = "E"
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
                                    session.findById(
                                        "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,1]").text = "E"
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
                                    session.findById(
                                        "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).text = format(
                                        guiData['cost'] / 1.06, '.2f')
                                    session.findById(
                                        "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).setFocus()
                                    session.findById(
                                        "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).caretPosition = 20
                                    session.findById("wnd[0]").sendVKey(0)

                                # session.findById("wnd[0]/tbar[0]/btn[11]").press()
                                session.findById("wnd[0]/tbar[0]/btn[3]").press()
                                session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                            # session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                    else:
                        session.findById(
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").text = \
                            guiData['materialCode']
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
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text = format(
                            revenue, '.2f')
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
                                    session.findById(
                                        "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,0]").text = "E"
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
                                    session.findById(
                                        "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,1]").text = "E"
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
                                    session.findById(
                                        "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).text = format(
                                        guiData['cost'] / 1.06, '.2f')
                                    session.findById(
                                        "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).setFocus()
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
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = \
                            guiData['longText']
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
                    amountVatStr = re.sub(r"(\d)(?=(\d\d\d)+(?!\d))", r"\1,", format(guiData['amountVat'], '.2f'))
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