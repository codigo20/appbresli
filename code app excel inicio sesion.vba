call Shell("C:/Program Files (x86)/SAP/FrontEnd/FrontEnd/SAPgui/saplogon.exe", vbMinimizedFocus)
Sleep 5 ' Funcion para esperar 5 segundos
Set SapGui = GetObject("SAPGUI")

Sleep 4
Set Appl = SapGui.GetScriptingEngine
Set Connection = Appl.Openconnection("RP3 Productivo", True) 'Nombre de la sesion SAP
Set session = Connection.Children(0)

'Loguearse facilmente sacar script de sap logeandose y pegarlo aqquí
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "3402251"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = "Anconsm20400*"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = "Anconsm20400*"
