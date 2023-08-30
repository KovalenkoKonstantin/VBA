Attribute VB_Name = "Cleaning"
Sub DecreaseWeightProcessing21()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "Processing21"
 SearchingString = "База взносов" 'ключ последней удаляемой колонки
 begin = 12 'первый ряд вставки
 
 'определение колонок рабочей книги
On Error Resume Next
For i = 1 To 20
    If Worksheets(SheetName).Cells(i, 1) = "Сотрудник" Then
        DataRow = i
    End If
Next i

 'определение последней удаляемой колоноки рабочей книги
On Error Resume Next
For i = 1 To 200
    If Worksheets(SheetName).Cells(DataRow, i) = SearchingString Then
        awLastCol = i
    End If
Next i

'удаление предыдущих данных
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 [a12] = 1
 'завершение
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub
Sub DecreaseWeightProcessing22()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "Processing22"
 SearchingString = "База взносов" 'ключ последней удаляемой колонки
 begin = 12 'первый ряд вставки
 
 'определение колонок рабочей книги
On Error Resume Next
For i = 1 To 20
    If Worksheets(SheetName).Cells(i, 1) = "Сотрудник" Then
        DataRow = i
    End If
Next i

 'определение последней удаляемой колоноки рабочей книги
On Error Resume Next
For i = 1 To 200
    If Worksheets(SheetName).Cells(DataRow, i) = SearchingString Then
        awLastCol = i
    End If
Next i

'удаление предыдущих данных
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 [a12] = 1
 'завершение
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub
Sub DecreaseWeightProcessing23()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "Processing23"
 SearchingString = "База взносов" 'ключ последней удаляемой колонки
 begin = 12 'первый ряд вставки
 
 'определение колонок рабочей книги
On Error Resume Next
For i = 1 To 20
    If Worksheets(SheetName).Cells(i, 1) = "Сотрудник" Then
        DataRow = i
    End If
Next i

 'определение последней удаляемой колоноки рабочей книги
On Error Resume Next
For i = 1 To 200
    If Worksheets(SheetName).Cells(DataRow, i) = SearchingString Then
        awLastCol = i
    End If
Next i

'удаление предыдущих данных
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 [a12] = 1
 'завершение
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub
Sub DecreaseWeightProcessing24()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "Processing24"
 SearchingString = "База взносов" 'ключ последней удаляемой колонки
 begin = 12 'первый ряд вставки
 
 'определение колонок рабочей книги
On Error Resume Next
For i = 1 To 20
    If Worksheets(SheetName).Cells(i, 1) = "Сотрудник" Then
        DataRow = i
    End If
Next i

 'определение последней удаляемой колоноки рабочей книги
On Error Resume Next
For i = 1 To 200
    If Worksheets(SheetName).Cells(DataRow, i) = SearchingString Then
        awLastCol = i
    End If
Next i

'удаление предыдущих данных
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 [a12] = 1
 'завершение
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub

Sub DecreaseWeightССЧ21()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "ССЧ21"
 SearchingString = "Количество дней простоя" 'ключ последней удаляемой колонки

 begin = 15 'первый ряд вставки
 
 'определение колонок рабочей книги
On Error Resume Next
For i = 1 To 20
    If Worksheets(SheetName).Cells(i, 1) = "Сотрудник" Then
        DataRow = i
    End If
Next i

 'определение последней удаляемой колоноки рабочей книги
On Error Resume Next
For i = 1 To 200
    If Worksheets(SheetName).Cells(DataRow, i) = SearchingString Then
        awLastCol = i
    End If
Next i

'удаление предыдущих данных
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 'завершение
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub

Sub DecreaseWeightССЧ22()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "ССЧ22"
 SearchingString = "Количество дней простоя" 'ключ последней удаляемой колонки

 begin = 15 'первый ряд вставки
 
 'определение колонок рабочей книги
On Error Resume Next
For i = 1 To 20
    If Worksheets(SheetName).Cells(i, 1) = "Сотрудник" Then
        DataRow = i
    End If
Next i

 'определение последней удаляемой колоноки рабочей книги
On Error Resume Next
For i = 1 To 200
    If Worksheets(SheetName).Cells(DataRow, i) = SearchingString Then
        awLastCol = i
    End If
Next i

'удаление предыдущих данных
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 'завершение
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub
Sub DecreaseWeightССЧ23()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "ССЧ23"
 SearchingString = "Количество дней простоя" 'ключ последней удаляемой колонки

 begin = 15 'первый ряд вставки
 
 'определение колонок рабочей книги
On Error Resume Next
For i = 1 To 20
    If Worksheets(SheetName).Cells(i, 1) = "Сотрудник" Then
        DataRow = i
    End If
Next i

 'определение последней удаляемой колоноки рабочей книги
On Error Resume Next
For i = 1 To 200
    If Worksheets(SheetName).Cells(DataRow, i) = SearchingString Then
        awLastCol = i
    End If
Next i

'удаление предыдущих данных
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 'завершение
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub
Sub DecreaseWeightССЧ24()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "ССЧ24"
 SearchingString = "Количество дней простоя" 'ключ последней удаляемой колонки

 begin = 15 'первый ряд вставки
 
 'определение колонок рабочей книги
On Error Resume Next
For i = 1 To 20
    If Worksheets(SheetName).Cells(i, 1) = "Сотрудник" Then
        DataRow = i
    End If
Next i

 'определение последней удаляемой колоноки рабочей книги
On Error Resume Next
For i = 1 To 200
    If Worksheets(SheetName).Cells(DataRow, i) = SearchingString Then
        awLastCol = i
    End If
Next i

'удаление предыдущих данных
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 ThisWorkbook.Sheets(SheetName).Range("C5").Select
 With Selection
        .Clear
 End With
 
 
 'завершение
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub

Sub DecreaseWeightРВ_Проекта()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "РВ_Проекта"
 SearchingString = "База взносов на проекте" 'ключ последней удаляемой колонки
 begin = 12 'первый ряд вставки
 
 'определение колонок рабочей книги
On Error Resume Next
For i = 1 To 20
    If Worksheets(SheetName).Cells(i, 1) = "Сотрудник" Then
        DataRow = i
    End If
Next i

 'определение последней удаляемой колоноки рабочей книги
On Error Resume Next
For i = 1 To 200
    If Worksheets(SheetName).Cells(DataRow, i) = SearchingString Then
        awLastCol = i
    End If
Next i

'удаление предыдущих данных
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 [a12] = 1
 'завершение
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub
Sub DecreaseWeightExpenditures()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "Expenditures"
 SearchingString = "База взносов" 'ключ последней удаляемой колонки
 begin = 12 'первый ряд вставки
 
 'определение колонок рабочей книги
On Error Resume Next
For i = 1 To 20
    If Worksheets(SheetName).Cells(i, 1) = "Сотрудник" Then
        DataRow = i
    End If
Next i

 'определение последней удаляемой колоноки рабочей книги
On Error Resume Next
For i = 1 To 200
    If Worksheets(SheetName).Cells(DataRow, i) = SearchingString Then
        awLastCol = i
    End If
Next i

'удаление предыдущих данных
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 [a12] = 1
 'завершение
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub
Sub DecreaseWeightBudget()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "Бюджет"
 SearchingString = "График работы" 'ключ последней удаляемой колонки
 begin = 5 'первый ряд вставки
 
 'определение колонок рабочей книги
On Error Resume Next
For i = 1 To 20
    If Worksheets(SheetName).Cells(i, 1) = "Должность" Then
        DataRow = i
    End If
Next i

 'определение последней удаляемой колоноки рабочей книги
On Error Resume Next
For i = 1 To 200
    If Worksheets(SheetName).Cells(DataRow, i) = SearchingString Then
        awLastCol = i
    End If
Next i

'удаление предыдущих данных
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 [a5] = 1
 'завершение
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 2

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub
Sub DecreaseWeightTabel()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "Табель"
 SearchingString = "График работы" 'ключ последней удаляемой колонки
 begin = 5 'первый ряд вставки
 
 'определение колонок рабочей книги
'On Error Resume Next
'For I = 1 To 20
'    If Worksheets(SheetName).Cells(I, 1) = "Должность" Then
'        DataRow = I
'    End If
'Next I

 'определение последней удаляемой колоноки рабочей книги
'On Error Resume Next
'For I = 1 To 200
'    If Worksheets(SheetName).Cells(DataRow, I) = SearchingString Then
'        awLastCol = I
'    End If
'Next I
awLastCol = 63

'удаление предыдущих данных
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "AU").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
' [a5] = 1
 'завершение
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 2

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub

Sub DecreaseWeightPayrollProject()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "РВ_Проекта"
 SearchingString = "База взносов" 'ключ последней удаляемой колонки
 begin = 12 'первый ряд вставки
 
 'определение колонок рабочей книги
On Error Resume Next
For i = 1 To 20
    If Worksheets(SheetName).Cells(i, 1) = "Сотрудник" Then
        DataRow = i
    End If
Next i

 'определение последней удаляемой колоноки рабочей книги
On Error Resume Next
For i = 1 To 200
    If Worksheets(SheetName).Cells(DataRow, i) = SearchingString Then
        awLastCol = i
    End If
Next i

'удаление предыдущих данных
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 [a12] = 1
 'завершение
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 2

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub

Sub DecreaseWeightSeconds()

  Dim kolvo As Variant
  Dim i As Long
  Dim list As Worksheet
  Dim ThisWorkbook As Workbook
  Dim SheetName As String
  Dim Sht As Worksheet
  
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False
  
  Set ThisWorkbook = ActiveWorkbook
  SheetName = "2"
  kolvo = 14
  
  'удаление предыдущих данных
  
  

  On Error Resume Next
  
  For i = 1 To kolvo
    Sheets("2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Первий диапазон. " & _
    "Выполнено: " & Int(100 * i / kolvo) & "%."
  Next i
  
  For i = 1 To 4
    Sheets("2_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Второй диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = 1 To 2
    Sheets("2_" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Третий диапазон. " & _
    "Выполнено: " & Int(100 * i / 2) & "%."
  Next i
  
  For i = 1 To 4
    Sheets("2_1_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Четвёртый диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = 1 To 4
    Sheets("2_2_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Пятый диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub
Sub DecreaseWeightNinth()

  Dim kolvo As Variant
  Dim i As Long
  Dim list As Worksheet
  Dim ThisWorkbook As Workbook
  Dim SheetName As String
  Dim Sht As Worksheet
  
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False
  
  Set ThisWorkbook = ActiveWorkbook
  SheetName = "9"
  kolvo = 14
  
  'удаление предыдущих данных
  
  

  On Error Resume Next
  
  For i = 1 To kolvo
    Sheets("9" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Первий диапазон. " & _
    "Выполнено: " & Int(100 * i / kolvo) & "%."
  Next i
  
  For i = 1 To 4
    Sheets("9_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Второй диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = 1 To 2
    Sheets("9_" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Третий диапазон. " & _
    "Выполнено: " & Int(100 * i / 2) & "%."
  Next i
  
  For i = 1 To 4
    Sheets("9_1_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Четвёртый диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = 1 To 4
    Sheets("9_2_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Пятый диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub
Sub DecreaseWeightTwentyth()

 Dim kolvo As Variant
  Dim i As Long
  Dim list As Worksheet
  Dim ThisWorkbook As Workbook
  Dim SheetName As String
  Dim Sht As Worksheet
  
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False
  
  Set ThisWorkbook = ActiveWorkbook
  SheetName = "20"
  kolvo = 14
  
  'удаление предыдущих данных
  
  

  On Error Resume Next
  
  For i = 1 To kolvo
    Sheets("20" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Первий диапазон. " & _
    "Выполнено: " & Int(100 * i / kolvo) & "%."
  Next i
  
  For i = 1 To 4
    Sheets("20_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Второй диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = 1 To 2
    Sheets("20_" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Третий диапазон. " & _
    "Выполнено: " & Int(100 * i / 2) & "%."
  Next i
  
  For i = 1 To 4
    Sheets("20_1_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Четвёртый диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = 1 To 4
    Sheets("20_2_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Пятый диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub
