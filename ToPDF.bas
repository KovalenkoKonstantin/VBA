Attribute VB_Name = "ToPDF"
Sub SaveToPDF()

 Start = Now()
 Dim array_distinct() 'задекларируем динамический массив
 Dim ThisWorkbook As Workbook
' Dim st, CheckBoxName, SaveName, Folder, Path As String
' Dim CheckBoxObject As Variant
 Dim sName As String
 Set ThisWorkbook = ActiveWorkbook
 
 On Error GoTo ExitHandler
 
 ThisWorkbook.Sheets("Preferences").Activate
 SaveName = ActiveSheet.Range("R30").Text
 ThisWorkbook.Activate
 Path = ActiveWorkbook.Path
 
 Application.ScreenUpdating = False
 Application.EnableEvents = False
 ActiveSheet.DisplayPageBreaks = False
 Application.DisplayStatusBar = False
 Application.DisplayAlerts = False


ThisWorkbook.Sheets(Array("1", "2", "2_23" _
        , "2_1", "2_1_23", "2_2", "2_2_23", _
        "9", "9_23", "9_1", "9_1_23", _
        "9_2", "9_2_23", _
        "10", "12", "20", "20_1", "20_2", _
        "21ф", "22ф", "23ф", _
        "П8")).Select

ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
Filename:=Path & "\" & _
SaveName & ".pdf", Quality:=xlQualityStandard _
, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
False

Finish = (Now() - Start) * 24 * 60 * 60

ExitHandler:
If Finish < 0.3 Then
    MsgBox ("Неправильные диапазоны")
Else
    MsgBox ("Файл сохранён в формате PDF в корневой папке")
End If

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
Exit Sub

End Sub


