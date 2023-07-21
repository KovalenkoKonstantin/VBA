Attribute VB_Name = "Engagement"
Sub SaveToEXL()

 Dim ThisWorkbook, wb As Workbook
 Dim SaveName, Folder, Path, myPathName As String
 Dim sName As String
 Dim fso As Object
 Set ThisWorkbook = ActiveWorkbook
 Set fso = CreateObject("Scripting.FileSystemObject")
 
 On Error GoTo ExitHandler
 
' ThisWorkbook.Sheets("Задействование2").Activate
 SaveName = ActiveSheet.Range("I1").Text
 ThisWorkbook.Activate
 Path = ThisWorkbook.Path
 
 Application.ScreenUpdating = False
 Application.EnableEvents = False
 ActiveSheet.DisplayPageBreaks = False
 Application.DisplayStatusBar = False
 Application.DisplayAlerts = False

'Sheets(Array("2_2", "2_2_23", "2_2_24", "7_2", "9_2э" _
'        , "10", "12", "20_2", "21ф", "22ф", "23ф", "П8")).Select

 'Type:=xlTypeXLS, _ '
'ActiveSheet.ExportAsFixedFormat xlTypeXLS, _
'Filename:=Path & "\" & _
'SaveName & ".xls", Type:=xlTypeXLS
', Quality:=xlQualityStandard _
', IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
'False

'
'ActiveWorkbook.SaveAs Filename:="D:\РКМ\ТФЦ\022-7\111.xlsx", FileFormat:= _
'        xlOpenXMLWorkbook, CreateBackup:=False

On Error Resume Next
fso.CreateFolder (Path & "\" & "Задействование")
myPathName = Path & "\" & "Задействование" & "\" & SaveName & ".xlsx"
If Dir(myPathName) <> "" Then Kill myPathName

ActiveSheet.Copy
ActiveWorkbook.SaveAs Path & "\" & "Задействование" & "\" & SaveName & ".xlsx"
ActiveWorkbook.Close

MsgBoxEx "Данные сохранены в папке Задействование", 0, "Выполнено", 1

ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
'    ThisWorkbook.Sheets("Preferences").Activate
Exit Sub

End Sub


'Sub SaveToEXL2()
'
' Dim ThisWorkbook, wb As Workbook
' Dim SaveName, Folder, Path, myPathName As String
' Dim sName As String
' Dim fso As Object
' Set ThisWorkbook = ActiveWorkbook
' Set fso = CreateObject("Scripting.FileSystemObject")
'
' On Error GoTo ExitHandler
'
'' ThisWorkbook.Sheets("Задействование2").Activate
' SaveName = ActiveSheet.Range("I1").Text
' ThisWorkbook.Activate
' Path = ThisWorkbook.Path
'
' Application.ScreenUpdating = False
' Application.EnableEvents = False
' ActiveSheet.DisplayPageBreaks = False
' Application.DisplayStatusBar = False
' Application.DisplayAlerts = False
'
''Sheets(Array("2_2", "2_2_23", "2_2_24", "7_2", "9_2э" _
''        , "10", "12", "20_2", "21ф", "22ф", "23ф", "П8")).Select
'
' 'Type:=xlTypeXLS, _ '
''ActiveSheet.ExportAsFixedFormat xlTypeXLS, _
''Filename:=Path & "\" & _
''SaveName & ".xls", Type:=xlTypeXLS
'', Quality:=xlQualityStandard _
'', IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
''False
'
''
''ActiveWorkbook.SaveAs Filename:="D:\РКМ\ТФЦ\022-7\111.xlsx", FileFormat:= _
''        xlOpenXMLWorkbook, CreateBackup:=False
'
'On Error Resume Next
'fso.CreateFolder (Path & "\" & "Подстановка")
'myPathName = Path & "\" & "Подстановка" & "\" & SaveName & ".xlsx"
'If Dir(myPathName) <> "" Then Kill myPathName
'
'ActiveSheet.Copy
'ActiveWorkbook.SaveAs Path & "\" & "Подстановка" & "\" & SaveName & ".xlsx"
'ActiveWorkbook.Close
'
'ExitHandler:
'    Application.ScreenUpdating = True
'    Application.EnableEvents = True
'    ActiveSheet.DisplayPageBreaks = True
'    Application.DisplayStatusBar = True
'    Application.DisplayAlerts = True
''    ThisWorkbook.Sheets("Preferences").Activate
'Exit Sub
'
'End Sub

