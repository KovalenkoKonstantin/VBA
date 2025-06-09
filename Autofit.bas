Attribute VB_Name = "Autofit"
Sub SetRowHeightToContent()
    Dim ws As Worksheet
    Dim mergedCell As Range
    Dim tempCell As Range
    Dim originalHeight As Double
    
    ' Находим лист с именем "3. Отчет"
    Set ws = ThisWorkbook.Worksheets("3. Отчет")
    ws.Activate
    
    ' Устанавливаем ссылку на объединенную ячейку
    Set mergedCell = ws.Range("A31:E31")
    
    ' Включаем перенос текста
    mergedCell.WrapText = True
    
    ' Сохраняем оригинальную высоту
    originalHeight = ws.Rows(31).RowHeight
    
    ' Временно используем очень большую высоту для расчета
    ws.Rows(31).RowHeight = 300
    
    ' Создаем временную ячейку для расчета
    Set tempCell = ws.Range("Z31")
    tempCell.Value = mergedCell.Value
    tempCell.Font.Size = mergedCell.Font.Size
    tempCell.Font.Name = mergedCell.Font.Name
    tempCell.WrapText = True
    tempCell.ColumnWidth = (mergedCell.Width / 7.5) ' Конвертируем в единицы ширины столбца
    
    ' Применяем автоподбор к временной ячейке
    tempCell.EntireRow.Autofit
    
    ' Получаем рассчитанную высоту
    Dim calculatedHeight As Double
    calculatedHeight = ws.Rows(31).RowHeight
    
    ' Очищаем временную ячейку
    tempCell.Clear
    
    ' Применяем рассчитанную высоту к нужной строке
    ws.Rows(31).RowHeight = calculatedHeight
    
End Sub
