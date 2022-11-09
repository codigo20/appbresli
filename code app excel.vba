

session.findById("wnd[0]").maximize

rem :::::::::::::::: Pegar despues de Maximizar code SAP

Dim objExcel
Dim objSheet, intRow, i
Set objExcel = GetObject(,"Excel.Application" or "calc.Application") rem irar bien este a ver si funciona con libre office
Set objSheet = objExcel.ActiveWorkbook.ActiveSheet
For i = 2 to objSheet.UsedRange.Rows.Count

c0L1 = Trim(CStr(objSheet.Cells(i, 1).Value))
c0L2 = Trim(CStr(objSheet.Cells(i, 2).Value))
c0L3 = Trim(CStr(objSheet.Cells(i, 3).Value))

rem ::::::::::::::::::::::::::::::::::::::::::::::::::::::

objExcel.Cells(i, 6).Value = session.findById("wnd[0]/sbar").Text

aux=col1 & " " & col2 & " " & col3
CreateObject("WScript.Shell").run("cmd /c @echo %date% %time% " & aux & " >> C:\SCRIPT\Pl0rCreationLog.txt")
next
msgbox "Accion Terminada con exito!"