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
        pass
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
        pass