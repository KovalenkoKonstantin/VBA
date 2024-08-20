Attribute VB_Name = "CopyWB"
Sub Copy_W()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False

    Dim WbLinks As Variant
    Dim SaveName As String
    Dim DistinctList As Variant
    Dim FullNameColumn As Range
    Dim i As Long
    Dim Path As String
    Dim FilePath As String
    
    ' Получаем путь к файлу и имя для сохранения
    Path = ActiveWorkbook.Path
    SaveName = ActiveSheet.Range("H30").Text
    Set FullNameColumn = ThisWorkbook.Sheets("Preferences").Range("I2:I9").SpecialCells(xlCellTypeVisible) ' Получаем диапазон значений без пустых ячеек
    
    ' Получаем уникальные значения из указанного диапазона
    DistinctList = GetDistinctItems(FullNameColumn)
    If IsEmpty(DistinctList) Then Exit Sub ' Убедимся, что массив не пустой

    ' Удаляем пустые значения (переписывание и изменение размера массива более оптимально)
    Dim NewList() As Variant
    Dim count As Long
    ReDim NewList(0)
    
    For i = LBound(DistinctList) To UBound(DistinctList)
        If Not IsEmpty(DistinctList(i)) Then
            NewList(count) = DistinctList(i)
            count = count + 1
            ReDim Preserve NewList(count) ' Пересоздаем массив
        End If
    Next i
    If count > 0 Then ReDim Preserve NewList(count - 1) ' Убираем последний ненужный элемент

    ' Добавляем новый элемент массива
    ReDim Preserve NewList(UBound(NewList) + 1)
    NewList(UBound(NewList)) = "Ninth"

    ' Копируем указанные листы
    ActiveWorkbook.Sheets(NewList).Copy
    ActiveWorkbook.PrecisionAsDisplayed = True

    
    ' Сохраняем значения с листа "Ф2 (1)" перед удалением
    ActiveWorkbook.Sheets("Ф2 (1)").Activate
    Cells.Select
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    ' Сохраняем значения с листа "ЗП (1)" перед удалением
    ActiveWorkbook.Sheets("ЗП (1)").Activate
    Cells.Select
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    
    ' Удаляем лишние листы
    On Error Resume Next ' Игнорируем ошибки при удалении
    Sheets("Ninth").delete
    Sheets("ПЗ").delete
    On Error GoTo 0

    ' Разрываем связи
    WbLinks = ActiveWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks)
    If Not IsEmpty(WbLinks) Then
        For i = LBound(WbLinks) To UBound(WbLinks)
            ActiveWorkbook.BreakLink Name:=WbLinks(i), Type:=xlLinkTypeExcelLinks
        Next i
    End If

    ' Для сохранения файла
    FilePath = Path & "\" & SaveName & ".xls"
    If Dir(FilePath) <> "" Then Kill FilePath
    ActiveWorkbook.SaveAs Filename:=Path & "\" & SaveName & ".xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

    ' Удаление "Табель" из массива
    Dim val As String: val = "Табель"
    Dim FindIndex As Long
    FindIndex = -1

    For i = LBound(NewList) To UBound(NewList)
        If NewList(i) = val Then
            FindIndex = i
            Exit For
        End If
    Next i

    ' Удаляем элемент "Табель", если он был найден
    If FindIndex <> -1 Then
        For i = FindIndex To UBound(NewList) - 1
            NewList(i) = NewList(i + 1)
        Next i
        ReDim Preserve NewList(LBound(NewList) To UBound(NewList) - 1)
    End If

    ' Снова выбираем оставшиеся листы
    ActiveWorkbook.Sheets(1).Activate

    ' Включаем все обратно
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True

End Sub
