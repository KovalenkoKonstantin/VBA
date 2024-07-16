Attribute VB_Name = "TableAndPivotTableRefresh"
'Sub RefreshAllTables()
'    Dim ws As Worksheet
'    Dim lo As ListObject
'    Dim info As String
'    Dim pt As PivotTable
'
'    For Each ws In ThisWorkbook.Worksheets
'        For Each lo In ws.ListObjects
'        On Error Resume Next
'            info = "Имя таблицы: " & lo.Name & vbCrLf
'            info = info & "Лист: " & ws.Name & vbCrLf
'            info = info & "Количество строк: " & lo.ListRows.Count & vbCrLf
'            info = info & "Количество столбцов: " & lo.ListColumns.Count & vbCrLf
'            info = info & vbCrLf
'            Debug.Print info
'            lo.QueryTable.Refresh BackgroundQuery:=False
'
'            lo.TableObject.Refresh
'
'        Next lo
'    Next ws
'    For Each ws In ThisWorkbook
'        pt.Refresh
'    Next ws
'
'End Sub


Sub RefreshAllTables()
    ' Объявляем переменные
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim info As String
    Dim pt As PivotTable
    
    ' Проходим по всем листам в текущей книге
    For Each ws In ThisWorkbook.Worksheets
        ' Проходим по всем таблицам (ListObjects) на текущем листе
        For Each lo In ws.ListObjects
            ' Игнорируем ошибки, чтобы продолжить выполнение кода даже при возникновении ошибки
            On Error Resume Next
            
            ' Формируем строку с информацией о таблице
            info = "Имя таблицы: " & lo.Name & vbCrLf
            info = info & "Лист: " & ws.Name & vbCrLf
            info = info & "Количество строк: " & lo.ListRows.Count & vbCrLf
            info = info & "Количество столбцов: " & lo.ListColumns.Count & vbCrLf
            info = info & vbCrLf
            
            ' Выводим информацию в окно отладки
            Debug.Print info
            
            ' Обновляем таблицу, если она связана с внешним источником данных
            lo.QueryTable.Refresh BackgroundQuery:=False
            
            ' Обновляем объект таблицы
            lo.TableObject.Refresh
        Next lo
    Next ws
    
    ' Проходим по всем листам в текущей книге
    For Each ws In ThisWorkbook.Worksheets
        ' Проходим по всем сводным таблицам на текущем листе
        For Each pt In ws.PivotTables
            ' Обновляем сводную таблицу
            pt.RefreshTable
        Next pt
    Next ws
End Sub
