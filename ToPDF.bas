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
 SaveName = ActiveSheet.Range("S30").Text
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


'ThisWorkbook.Sheets(Array("1. Акт", "2. Прот", "3. ЭП", "4. СЦ" _
'        , "5. ПЗ", "6. Мат", "7. ЗП", "8. Проч", "9. КУЗ", "10. КЗП", _
'        "11. НР", "12. АктЭО", "13. АктСПО", "14. ПротССП", _
'        "15. ПО", "16. Рез Инв", "17. Сохр", "18. Акт матц", "19. Приказ", _
'        "20.Спр.%накл.", "21.Спр.%стр.")).Select
'ThisWorkbook.Sheets(Array("1", _
'        "2", "2_21", "2_22", "2_23", "2_1", "2_1_21", "2_1_22", "2_1_23", "2_2", "2_2_23", _
'        "3", "3_21", "3_22", "3_23", _
'        "9", "9_21", "9_22", "9_23", "9_1", "9_1_21", "9_1_22", "9_1_23", "9_2", "9_2_23", _
'        "10", "12", _
'        "20", "20_1", "20_2", _
'        "21ф", "22ф", "23Ф", _
'        "П5", "п6", "П7", "П8", "НЧ", "Приказ", "КУЗ_1", "КУЗ_2")).Select
'End If
ElseIf ThisWorkbook.Sheets("Preferences").Range("C13").Value2 = _
"Программно-аппаратный комплекс ViPNet Coordinator HW2000" Then
ThisWorkbook.Sheets(Array("1", _
        "2", "4д", "6", "9", _
        "10", "12", _
        "20", "21ф", "22ф", "23Ф", _
        "Труд", "Прайс")).Select
ElseIf ThisWorkbook.Sheets("Preferences").Range("C13").Value2 = _
"Дружба" Then
ThisWorkbook.Sheets(Array("1", _
        "2", "4д", "6", "9", _
        "10", "12", _
        "20", "21ф", "22ф", "23Ф", _
        "Труд", "Прайс")).Select
Else
ThisWorkbook.Sheets(Array("1", _
        "2", _
        "4д", _
        "9", _
        "10", "12", _
        "20", _
        "21ф", "22ф", "23Ф", _
        "П5", "П6", "П7", "П8", _
        "НЧ", "КУЗ_1", "Табель")).Select
End If
        
        

ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
Filename:=Path & "\" & _
SaveName & ".pdf", Quality:=xlQualityStandard _
, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
False

Finish = (Now() - Start) * 24 * 60 * 60

ExitHandler:
If Finish < 0.1 Then
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


