Attribute VB_Name = "Layout"
Sub LayoutOn()

 Dim ThisWorkbook As Workbook
 Dim ws As Worksheet
 HideSys
 tottal = Application.Sheets.Count
 
' Dim xSht As Variant
'    Dim I As Long
'    For Each xSht In ActiveWorkbook.Sheets
'        If xSht.Visible Then I = I + 1
'    Next
'tottal = I

Set ThisWorkbook = ActiveWorkbook
CompanyName = ThisWorkbook.Sheets("Preferences").Range("C7").Value2 'имя компании
 On Error GoTo ExitHandler
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False

For Each ws In ThisWorkbook.Worksheets
    If ws.Visible = True Then
        ws.Activate
        aws = Application.ActiveSheet.Index
    
    'статус бар
    Application.StatusBar = "Обрабатывается " + Str(aws) _
    + " лист из " + Str(tottal) + " листов. Выполнено: " _
    + Str(Int(aws / tottal * 100)) + " %. " + "Расчётное время до конца выполнения программы: " _
    + Str(Int((Str(tottal) - Str(aws)) * 3)) + " секунд(ы)."

'        ActiveSheet.PageSetup.CenterHeaderPicture.Filename = _
'        "D:\Аннотация.png"

            With ActiveSheet.PageSetup
'                .CenterHeader = "&G"
                .CenterHeader = _
                "&""Times New Roman,обычный""&KFF0000Данный документ не согласован."
'                .RightHeader = _
'                "&""Times New Roman,обычный""&KFF0000Любое несанкционированное раскрытие, распространение или копирование ЗАПРЕЩЕНО."
                .RightFooter = _
                "&""Times New Roman,обычный""&KFF0000Настоящий документ и любые приложения к нему содержат информацию относящуюся к коммерческой тайне " & CompanyName
            End With
    End If
Next ws

ExitHandler:
Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
ThisWorkbook.Sheets("Preferences").Activate
ActiveWindow.View = xlNormalView
 
Exit Sub
 
ErrHandler:
MsgBox Err.Description
Resume ExitHandler

End Sub


Sub LayoutOff()

 Dim ThisWorkbook As Workbook
 Dim ws As Worksheet
 HideSys
 tottal = Application.Sheets.Count

Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False

For Each ws In ThisWorkbook.Worksheets
    If ws.Visible = True Then
        ws.Activate
        aws = Application.ActiveSheet.Index
        
        'статус бар
        Application.StatusBar = "Обрабатывается " + Str(aws) _
        + " лист из " + Str(tottal) + " листов. Выполнено: " _
        + Str(Int(aws / tottal * 100)) + " %. " + "Расчётное время до конца выполнения программы: " _
        + Str(Int((Str(tottal) - Str(aws)) * 3)) + " секунд."
        
        ActiveSheet.PageSetup.CenterHeaderPicture.Filename = ""
        
            With ActiveSheet.PageSetup
                .CenterHeader = ""
                .RightHeader = ""
                .RightFooter = ""
            End With
    End If
Next ws
ExitHandler:
Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
ThisWorkbook.Sheets("Preferences").Activate
ActiveWindow.View = xlNormalView
 
Exit Sub
 
ErrHandler:
MsgBox Err.Description
Resume ExitHandler

End Sub



