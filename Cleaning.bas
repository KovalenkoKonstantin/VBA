Attribute VB_Name = "Cleaning"
Sub DecreaseWeightProcessing21()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 SheetName = "Processing21"
 ps = "123$"
  
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 
 SearchingString = "База взносов" 'ключ последней удаляемой колонки
 begin = 12 'первый ряд вставки
 
ThisWorkbook.UnProtect Password:=ps
ThisWorkbook.Sheets(SheetName).UnProtect Password:=ps
ThisWorkbook.Sheets(SheetName).Visible = True
ThisWorkbook.Sheets(SheetName).Activate
ActiveSheet.ShowAllData

If ThisWorkbook.Sheets("Preferences").Range("W87").Value2 = False Then
 GoTo exithandler
 End If
 
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
 
exithandler:
 'завершение
ThisWorkbook.Sheets(SheetName).Protect Password:=ps
ThisWorkbook.Sheets(SheetName).Visible = False
ThisWorkbook.Sheets("Preferences").Activate
'MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    ThisWorkbook.Sheets("Preferences").UnProtect Password:=ps
    Rows("81:91").EntireRow.AutoFit
    ThisWorkbook.Sheets("Preferences").Protect Password:=ps
    ThisWorkbook.Protect Password:=ps
    
End Sub
Sub DecreaseWeightProcessing22()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 SheetName = "Processing22"
 ps = "123$"
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 
 SearchingString = "База взносов" 'ключ последней удаляемой колонки
 begin = 12 'первый ряд вставки
 
ThisWorkbook.UnProtect Password:=ps
ThisWorkbook.Sheets(SheetName).UnProtect Password:=ps
ThisWorkbook.Sheets(SheetName).Visible = True
ThisWorkbook.Sheets(SheetName).Activate
ActiveSheet.ShowAllData

If ThisWorkbook.Sheets("Preferences").Range("W88").Value2 = False Then
 GoTo exithandler
 End If
 
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
 
exithandler:
 'завершение
ThisWorkbook.Sheets(SheetName).Protect Password:=ps
ThisWorkbook.Sheets(SheetName).Visible = False
ThisWorkbook.Sheets("Preferences").Activate
'MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    ThisWorkbook.Sheets("Preferences").UnProtect Password:=ps
    Rows("81:91").EntireRow.AutoFit
    ThisWorkbook.Sheets("Preferences").Protect Password:=ps
    ThisWorkbook.Protect Password:=ps
    
End Sub

Sub DecreaseWeightProcessing23()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 SheetName = "Processing23"
 ps = "123$"
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 
 SearchingString = "База взносов" 'ключ последней удаляемой колонки
 begin = 12 'первый ряд вставки
 
ThisWorkbook.UnProtect Password:=ps
ThisWorkbook.Sheets(SheetName).UnProtect Password:=ps
ThisWorkbook.Sheets(SheetName).Visible = True
ThisWorkbook.Sheets(SheetName).Activate
ActiveSheet.ShowAllData

If ThisWorkbook.Sheets("Preferences").Range("W88").Value2 = False Then
 GoTo exithandler
 End If
 
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
 
exithandler:
 'завершение
ThisWorkbook.Sheets(SheetName).Protect Password:=ps
ThisWorkbook.Sheets(SheetName).Visible = False
ThisWorkbook.Sheets("Preferences").Activate
'MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    ThisWorkbook.Sheets("Preferences").UnProtect Password:=ps
    Rows("81:91").EntireRow.AutoFit
    ThisWorkbook.Sheets("Preferences").Protect Password:=ps
    ThisWorkbook.Protect Password:=ps
    
End Sub
Sub DecreaseWeightССЧ21()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 SheetName = "ССЧ21"
 ps = "123$"
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 
 SearchingString = "-" 'ключ последней удаляемой колонки
 begin = 15 'первый ряд вставки
 
ThisWorkbook.UnProtect Password:=ps
ThisWorkbook.Sheets(SheetName).UnProtect Password:=ps
ThisWorkbook.Sheets(SheetName).Visible = True
ThisWorkbook.Sheets(SheetName).Activate

If ThisWorkbook.Sheets("Preferences").Range("W89").Value2 = False Then
 GoTo exithandler
 End If
 
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
 
exithandler:
 'завершение
ThisWorkbook.Sheets(SheetName).Protect Password:=ps
ThisWorkbook.Sheets(SheetName).Visible = False
ThisWorkbook.Sheets("Preferences").Activate
'MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    ThisWorkbook.Sheets("Preferences").UnProtect Password:=ps
    Rows("81:91").EntireRow.AutoFit
    ThisWorkbook.Sheets("Preferences").Protect Password:=ps
    ThisWorkbook.Protect Password:=ps
    
End Sub

Sub DecreaseWeightССЧ22()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 SheetName = "ССЧ22"
 ps = "123$"
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 
 SearchingString = "-" 'ключ последней удаляемой колонки

 begin = 15 'первый ряд вставки
 
ThisWorkbook.UnProtect Password:=ps
ThisWorkbook.Sheets(SheetName).UnProtect Password:=ps
ThisWorkbook.Sheets(SheetName).Visible = True
ThisWorkbook.Sheets(SheetName).Activate

If ThisWorkbook.Sheets("Preferences").Range("W90").Value2 = False Then
 GoTo exithandler
 End If
 
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
 
exithandler:
 'завершение
ThisWorkbook.Sheets(SheetName).Protect Password:=ps
ThisWorkbook.Sheets(SheetName).Visible = False
ThisWorkbook.Sheets("Preferences").Activate
'MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    ThisWorkbook.Sheets("Preferences").UnProtect Password:=ps
    Rows("81:91").EntireRow.AutoFit
    ThisWorkbook.Sheets("Preferences").Protect Password:=ps
    ThisWorkbook.Protect Password:=ps
    
End Sub

Sub DecreaseWeightАнализы()

 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim SheetName As String
 ps = "123$"
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo exithandler
 SheetName = "Ан.сч26"
 awLastCol = 9
  
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

ThisWorkbook.UnProtect Password:=ps

 If ThisWorkbook.Sheets("Preferences").Range("W83").Value2 = False Then
    GoTo exithandler
 End If
 
ThisWorkbook.Sheets(SheetName).UnProtect Password:=ps
ThisWorkbook.Sheets(SheetName).Visible = True
ThisWorkbook.Sheets(SheetName).Activate

On Error Resume Next
ActiveSheet.ShowAllData

'удаление предыдущих данных
 ThisWorkbook.Sheets(SheetName).Activate
awLastRow = Cells(Rows.Count, "D").End(xlUp).row

Range(Cells(1, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With

ThisWorkbook.Sheets(SheetName).Visible = False

exithandler:
SheetName = "Ан.сч44"
 If ThisWorkbook.Sheets("Preferences").Range("W84").Value2 = False Then
    GoTo exithandler1
 End If
 
ThisWorkbook.Sheets(SheetName).UnProtect Password:=ps
ThisWorkbook.Sheets(SheetName).Visible = True
ThisWorkbook.Sheets(SheetName).Activate
ActiveSheet.ShowAllData
ThisWorkbook.Sheets(SheetName).Activate
awLastRow = Cells(Rows.Count, "D").End(xlUp).row
Range(Cells(1, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With

ThisWorkbook.Sheets(SheetName).Visible = False

exithandler1:
SheetName = "Ан.сч90"
 If ThisWorkbook.Sheets("Preferences").Range("W85").Value2 = False Then
    GoTo exithandler2
 End If
 
On Error Resume Next
ThisWorkbook.Sheets(SheetName).UnProtect Password:=ps
ThisWorkbook.Sheets(SheetName).Visible = True
ThisWorkbook.Sheets(SheetName).Activate
ActiveSheet.ShowAllData
 ThisWorkbook.Sheets(SheetName).Activate
awLastRow = Cells(Rows.Count, "D").End(xlUp).row
Range(Cells(1, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 
 ThisWorkbook.Sheets(SheetName).Visible = False
 
exithandler2:
SheetName = "Ан.сч20"
 If ThisWorkbook.Sheets("Preferences").Range("W82").Value2 = False Then
    GoTo exithandler3
 End If

ThisWorkbook.Sheets(SheetName).UnProtect Password:=ps
ThisWorkbook.Sheets(SheetName).Visible = True
ThisWorkbook.Sheets(SheetName).Activate
ActiveSheet.ShowAllData
 ThisWorkbook.Sheets(SheetName).Activate
awLastRow = Cells(Rows.Count, "D").End(xlUp).row
Range(Cells(1, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
    
'завершение
importWB.Close
ThisWorkbook.Sheets(SheetName).Activate
'MsgBox ("Данные успешно удалены")

ThisWorkbook.Sheets(SheetName).Protect Password:=ps
ThisWorkbook.Sheets(SheetName).Visible = False

exithandler3:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    ThisWorkbook.Sheets("Preferences").UnProtect Password:=ps
        Rows("81:91").EntireRow.AutoFit
    ThisWorkbook.Sheets("Preferences").Protect Password:=ps
    ThisWorkbook.Protect Password:=ps
 Exit Sub
    
End Sub

Sub DecreaseWeightOFR()

 Dim ThisWorkbook, importWB As Workbook
 Dim SheetName As String
 ps = "123$"
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo exithandler
 SheetName = "ОФР"
 awLastCol = 20
  
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False


ThisWorkbook.UnProtect Password:=ps
ThisWorkbook.Sheets(SheetName).UnProtect Password:=ps
ThisWorkbook.Sheets(SheetName).Visible = True
ThisWorkbook.Sheets(SheetName).Activate
On Error Resume Next
ActiveSheet.ShowAllData

 If ThisWorkbook.Sheets("Preferences").Range("W86").Value2 = False Then
    GoTo exithandler
 End If

'удаление предыдущих данных
 ThisWorkbook.Sheets(SheetName).Activate
awLastRow = Cells(Rows.Count, "N").End(xlUp).row
Range(Cells(1, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
    
'завершение
ThisWorkbook.Sheets(SheetName).Activate
Company = ThisWorkbook.Sheets(SheetName).Cells(10, 22).Value2
Period = ThisWorkbook.Sheets(SheetName).Cells(10, 23).Value2

ThisWorkbook.Sheets(SheetName).Protect Password:=ps
ThisWorkbook.Sheets(SheetName).Visible = False

exithandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets(SheetName).Visible = False
    ThisWorkbook.Sheets("Preferences").Activate
    ThisWorkbook.Sheets("Preferences").UnProtect Password:=ps
        Rows("81:91").EntireRow.AutoFit
    ThisWorkbook.Sheets("Preferences").Protect Password:=ps
    ThisWorkbook.Protect Password:=ps
 Exit Sub
 
errhandler:
 MsgBox Err.Description
 ThisWorkbook.Sheets(SheetName).Visible = False
 Resume exithandler
End Sub
