Attribute VB_Name = "Merge"
Sub Merge9()

 Dim ThisWorkbook As Workbook
 Dim ws As Worksheet
 a = "9" 'лист
 b = 13 'первая строка
 C = 200 'последняя строка цикла

Set ThisWorkbook = ActiveWorkbook
On Error GoTo ExitHandler
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False
    

ThisWorkbook.Sheets(a).Activate

Range("E12:E500").Select
For i = b To C

'статус бар
    Application.StatusBar = "Выполнено: " _
    + str(Int(i / C * 100)) + " %. "
    
    If Range("E" & i).Value2 = 0 And Range("E" & i).Value2 <> "" Then
        Range("E" & i & ":E" & (i - 1)).Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Selection.Merge
    End If

    If Range("F" & i).Value2 = 0 And Range("F" & i).Value2 <> "" Then
        Range("F" & i & ":F" & (i - 1)).Select
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
    Selection.Merge
    End If
Next i

ExitHandler:
Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
ActiveWindow.View = xlNormalView
 
Exit Sub
 
ErrHandler:
MsgBox Err.Description
Resume ExitHandler

End Sub

