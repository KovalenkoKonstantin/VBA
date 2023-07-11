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

If ThisWorkbook.Sheets("Preferences").Range("C13").Value2 = "Поиск-ПМ" Then
    ThisWorkbook.Sheets(Array("ПМ.Прот", "ПМ.ПЗ", "ПМ.П2", "ПМ.ОК" _
        , "ПМ.Ф2", "ПМ.Вед", "ПМ.ОжЗп", "ПМ.КУЗ", "ПМ.НР")).Select
Else

ThisWorkbook.Sheets(Array("2", "2_21", "2_22", "2_23" _
        , "2_1", "2_1_21", "2_1_22", "2_1_23", "2_2", "2_2_23", _
        "3", "3_21", "3_22", "3_23", _
        "9", "9_21", "9_22", "9_23", "9_1", "9_1_21", "9_1_22", "9_1_23", _
        "9_2", "9_2_23", _
        "10", "12", "20", "20_1", "20_2", _
        "21ф", "22ф", "23ф", _
        "Приказ", "КУЗ_1", "КУЗ_2", "П5", "П6", "П7", "П8", "НЧ", _
        "ПЗ", "Табель", "Сопр", "Опись")).Select
End If
        
        

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
    MsgBoxEx "Файл сохранён в формате PDF в корневой папке", 0, "Done", 1
End If

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
Exit Sub

End Sub


