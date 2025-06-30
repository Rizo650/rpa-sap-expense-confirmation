If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/radP_AI").setFocus
session.findById("wnd[0]/usr/radP_AI").select
session.findById("wnd[0]/usr/ctxtS_PDATE-LOW").text = "01.05.2025"
session.findById("wnd[0]/usr/ctxtS_PDATE-HIGH").text = "31.05.2025"
session.findById("wnd[0]/usr/ctxtS_PDATE-HIGH").setFocus
session.findById("wnd[0]/usr/ctxtS_PDATE-HIGH").caretPosition = 10
session.findById("wnd[0]/usr/btnITEM_SEL").press
session.findById("wnd[1]").sendVKey 71
session.findById("wnd[2]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").text = "profit center"
session.findById("wnd[2]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 13
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/tblSAPLSE16NMULTI_OR_TC/ctxtGS_MULTI_OR-LOW[2,0]").text = "101030"
session.findById("wnd[1]/usr/tblSAPLSE16NMULTI_OR_TC/ctxtGS_MULTI_OR-LOW[2,0]").setFocus
session.findById("wnd[1]/usr/tblSAPLSE16NMULTI_OR_TC/ctxtGS_MULTI_OR-LOW[2,0]").caretPosition = 6
session.findById("wnd[1]").sendVKey 71
session.findById("wnd[2]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").text = "entry date"
session.findById("wnd[2]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 10
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/tblSAPLSE16NMULTI_OR_TC/ctxtGS_MULTI_OR-LOW[2,0]").text = "01.05.2025"
session.findById("wnd[1]/usr/tblSAPLSE16NMULTI_OR_TC/ctxtGS_MULTI_OR-HIGH[3,0]").text = "26.05.2025"
session.findById("wnd[1]/usr/tblSAPLSE16NMULTI_OR_TC/ctxtGS_MULTI_OR-HIGH[3,0]").setFocus
session.findById("wnd[1]/usr/tblSAPLSE16NMULTI_OR_TC/ctxtGS_MULTI_OR-HIGH[3,0]").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtP_LAYOUT").text = "/SCAN-EXP"
session.findById("wnd[0]/usr/ctxtP_LAYOUT").setFocus
session.findById("wnd[0]/usr/ctxtP_LAYOUT").caretPosition = 9
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/shellcont/shell").setCurrentCell 1,"GA_TXT50"
session.findById("wnd[0]/shellcont/shell").contextMenu
session.findById("wnd[0]/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "E:\RPA Flow\IT.DEV.SAP.ExpenseConfirmation\Data\Source Data"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Raw Import Expense W1 - BP_26052025.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 7
session.findById("wnd[1]/tbar[0]/btn[0]").press
