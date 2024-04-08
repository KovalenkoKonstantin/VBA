Attribute VB_Name = "CopyWB"
Sub Copy_W()

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

Dim WbLinks
Dim SaveName As String
Dim DistinctList As Variant
Dim FullNameColumn As Range

Path = ActiveWorkbook.Path
SaveName = ActiveSheet.Range("H30").Text
ThisWorkbook.Sheets("Preferences").Activate
Set FullNameColumn = ActiveSheet.UsedRange.Range("I2:I20") ' Диапазон значений. С пустыми ячейками
DistinctList = GetDistinctItems(FullNameColumn) ' Передаем диапазон в функцию.


'удаляю значение массива содержащее пустую ячейку
n = LBound(DistinctList) ' удаляемый элемент он первый т.к. массив отсортирован функцией
For i = n To UBound(DistinctList) - 1
    DistinctList(i) = DistinctList(i + 1)
Next
ReDim Preserve DistinctList(LBound(DistinctList) To i - 1)


'дебаг
Debug.Print Join(DistinctList, vbCrLf) ' Выводим результат.
Debug.Print ("____________________________")
'добавялю новый элемент массива
ReDim Preserve DistinctList(UBound(DistinctList) + 1)
'задаю имя нового элемента
DistinctList(UBound(DistinctList)) = "Ninth"
'дебаг
Debug.Print Join(DistinctList, vbCrLf) ' Выводим результат.

ActiveWorkbook.Sheets(DistinctList).Copy
'значения как на экране
ActiveWorkbook.PrecisionAsDisplayed = True
'удаляю последний элемент массива
ReDim Preserve DistinctList(UBound(DistinctList) - 1)

Sheets(DistinctList).Select

'создаю копию книги
Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
'удаляю лишние листы
Sheets("Ninth").delete
Sheets("Табель").delete

'разрываю связи
WbLinks = ActiveWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks)
If Not IsEmpty(WbLinks) Then
    For i = LBound(WbLinks) To UBound(WbLinks)
        ActiveWorkbook.BreakLink Name:=WbLinks(i), Type:=xlLinkTypeExcelLinks
    Next
Else
End If

'по названию первого элемента массива активирую лист книги
ActiveWorkbook.Sheets(DistinctList(LBound(DistinctList))).Select

'удаление файла если уже существует
FilePath = Path & "\" & SaveName & ".xls"
If Dir(FilePath) <> "" Then
    Kill FilePath
    ActiveWorkbook.SaveAs Filename:=Path & "\" & _
    SaveName & ".xls"
Else
    ActiveWorkbook.SaveAs Filename:=Path & "\" & _
    SaveName & ".xls"
End If

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True

'NewWb.Close
    
End Sub
