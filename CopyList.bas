Attribute VB_Name = "CopyList"
Sub Clone2()
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
Application.Calculation = xlManual
  
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
  
  For i = 3 To 6
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
  
  For i = 3 To 6
    Sheets("2_1_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Четвёртый диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = 3 To 6
    Sheets("2_2_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Пятый диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
 
  'копирвоание листов
  ThisWorkbook.Sheets(SheetName).Activate
  Set list = ActiveSheet
'kolvo = InputBox("Укажите необходимое количество листов")

'If kolvo = "" Then Exit Sub
'If IsNumeric(kolvo) Then
'    kolvo = Fix(kolvo)
    For i = 1 To kolvo
        list.Copy After:=ActiveSheet
        ActiveSheet.Name = list.Name & i
        'статус бар
        Application.StatusBar = "Копирование листов. " & _
        "Выполнено: " & Int(100 * i / kolvo) & "%."
    Next
'Else
'    MsgBox "Неправильное число"
'End If

'переименование листов
On Error Resume Next
For Each Sht In Worksheets
    'статус бар
    Application.StatusBar = "Переименование листов."
    If Sht.Name = "21" Then Sht.Name = "2_23"
    If Sht.Name = "22" Then Sht.Name = "2_24"
    If Sht.Name = "23" Then Sht.Name = "2_25"
    If Sht.Name = "24" Then Sht.Name = "2_26"
    If Sht.Name = "25" Then Sht.Name = "2_1"
    If Sht.Name = "26" Then Sht.Name = "2_1_23"
    If Sht.Name = "27" Then Sht.Name = "2_1_24"
    If Sht.Name = "28" Then Sht.Name = "2_1_25"
    If Sht.Name = "29" Then Sht.Name = "2_1_26"
    If Sht.Name = "210" Then Sht.Name = "2_2"
    If Sht.Name = "211" Then Sht.Name = "2_2_23"
    If Sht.Name = "212" Then Sht.Name = "2_2_24"
    If Sht.Name = "213" Then Sht.Name = "2_2_25"
    If Sht.Name = "214" Then Sht.Name = "2_2_26"
Next

'выставление настроек
  On Error Resume Next
  
  'года
  For i = 3 To 6
    Sheets("2_2" & i).Activate
        [Q4] = "202" & i
        Range("E68:F68").ClearContents
        Range("I68:I68").ClearContents
        Range("X3:AC60").Clear
        [D71] = "20_2" & i
        [G68] = "202" & i
        [H68] = "202" & i
        'статус бар
        Application.StatusBar = "Выставление настроек листов. Первий диапазон. " & _
        "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  'этапы
  For i = 1 To 2
    Sheets("2_" & i).Activate
        [Q3] = "Этап " & i
        Range("X3:AC60").Clear
        Range("E73:I74").ClearContents
        [D71] = "20_" & i
        [E72] = "Этап " & i
        [F72] = "Этап " & i
        [G72] = "Этап " & i
        [H72] = "Этап " & i
        [I72] = "Этап " & i
        'статус бар
        Application.StatusBar = "Выставление настроек листов. Второй диапазон. " & _
        "Выполнено: " & Int(100 * i / 2) & "%."
        
  Next i
  
  'Этап 1
  For i = 3 To 6
    Sheets("2_1_2" & i).Activate
        [Q3] = "Этап 1"
        [Q4] = "202" & i
        Range("X3:AC60").Clear
        [D71] = "20_1_2" & i
        Range("E73:I74").ClearContents
        [E72] = "Этап 1"
        [F72] = "Этап 1"
        [G72] = "Этап 1"
        [H72] = "Этап 1"
        [I72] = "Этап 1"
        'статус бар
        Application.StatusBar = "Выставление настроек листов. Третий диапазон. " & _
        "Выполнено: " & Int(100 * i / 4) & "%."
        
  Next i
  
  'Этап 2
  For i = 3 To 6
    Sheets("2_2_2" & i).Activate
        [Q3] = "Этап 2"
        [Q4] = "202" & i
        Range("X3:AC60").Clear
        [D71] = "20_2_2" & i
        Range("E73:I74").ClearContents
        [E72] = "Этап 2"
        [F72] = "Этап 2"
        [G72] = "Этап 2"
        [H72] = "Этап 2"
        [I72] = "Этап 2"
        'статус бар
        Application.StatusBar = "Выставление настроек листов. Четвёртый диапазон. " & _
        "Выполнено: " & Int(100 * i / 4) & "%."
        
  Next i
  
  'донастройка диапазонов
  Sheets("2_1_23").Activate
  Range("F68:I68").ClearContents
  
  Sheets("2_1_24").Activate
  Range("G68:I68").ClearContents
  Range("E68:E68").ClearContents
  
  Sheets("2_1_25").Activate
  Range("E68:F68").ClearContents
  Range("I68:I68").ClearContents
  
  Sheets("2_1_26").Activate
  Range("E68:H68").ClearContents
  
'____________________________________________________________________
  
  Sheets("2_2_23").Activate
  Range("F68:I68").ClearContents
  
  Sheets("2_2_24").Activate
  Range("G68:I68").ClearContents
  Range("E68:E68").ClearContents
  
  Sheets("2_2_25").Activate
  Range("E68:F68").ClearContents
  Range("I68:I68").ClearContents
  
  Sheets("2_2_26").Activate
  Range("E68:H68").ClearContents
  

ThisWorkbook.Sheets("Preferences").Activate
Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
    
End Sub
Sub Clone6()
  Dim kolvo As Variant
  Dim i As Long
  Dim list As Worksheet
  Dim ThisWorkbook As Workbook
  Dim SheetName, var As String
  Dim Sht As Worksheet
  Dim WorkRng As Range
  
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual
  
  Set ThisWorkbook = ActiveWorkbook
  SheetName = "6"
  kolvo = 15
  
  'удаление предыдущих данных

  On Error Resume Next
  
  For i = 1 To kolvo
    Sheets("6" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Первий диапазон. " & _
    "Выполнено: " & Int(100 * i / kolvo) & "%."
  Next i
  
  For i = 3 To 6
    Sheets("6_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Второй диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = 1 To 2
    Sheets("6_" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Третий диапазон. " & _
    "Выполнено: " & Int(100 * i / 2) & "%."
  Next i
  
  For i = 3 To 6
    Sheets("6_1_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Четвёртый диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = 3 To 6
    Sheets("6_2_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Пятый диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
 
  'копирование листов
  ThisWorkbook.Sheets(SheetName).Activate
  Set list = ActiveSheet
'kolvo = InputBox("Укажите необходимое количество листов")

'If kolvo = "" Then Exit Sub
'If IsNumeric(kolvo) Then
'    kolvo = Fix(kolvo)
    For i = 2 To kolvo
        list.Copy After:=ActiveSheet
        ActiveSheet.Name = list.Name & i
        'статус бар
        Application.StatusBar = "Копирование листов. " & _
        "Выполнено: " & Int(100 * i / kolvo) & "%."
    Next
'Else
'    MsgBox "Неправильное число"
'End If

'переименование листов
On Error Resume Next
For Each Sht In Worksheets
    'статус бар
    Application.StatusBar = "Переименование листов."
    If Sht.Name = "62" Then Sht.Name = "6_23"
    If Sht.Name = "63" Then Sht.Name = "6_24"
    If Sht.Name = "64" Then Sht.Name = "6_25"
    If Sht.Name = "65" Then Sht.Name = "6_26"
    If Sht.Name = "66" Then Sht.Name = "6_1"
    If Sht.Name = "67" Then Sht.Name = "6_1_23"
    If Sht.Name = "68" Then Sht.Name = "6_1_24"
    If Sht.Name = "69" Then Sht.Name = "6_1_25"
    If Sht.Name = "610" Then Sht.Name = "6_1_26"
    If Sht.Name = "611" Then Sht.Name = "6_2"
    If Sht.Name = "612" Then Sht.Name = "6_2_23"
    If Sht.Name = "613" Then Sht.Name = "6_2_24"
    If Sht.Name = "614" Then Sht.Name = "6_2_25"
    If Sht.Name = "615" Then Sht.Name = "6_2_26"
Next

'выставление настроек
  On Error Resume Next
  
  'года
  For i = 3 To 6
    Sheets("6_2" & i).Activate
        [AD2] = "202" & i

    'статус бар
    Application.StatusBar = "Выставление настроек листов. Первий диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  'этапы
  For i = 1 To 2
    Sheets("6_" & i).Activate
        [AD1] = "Этап " & i
        'статус бар
        Application.StatusBar = "Выставление настроек листов. Второй диапазон. " & _
        "Выполнено: " & Int(100 * i / 2) & "%."
  Next i
  
  'Этап 1
  For i = 3 To 6
    Sheets("6_1_2" & i).Activate
        [AD1] = "Этап 1"
        [AD2] = "202" & i
        'статус бар
        Application.StatusBar = "Выставление настроек листов. Третий диапазон. " & _
        "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  'Этап 2
  For i = 3 To 6
    Sheets("6_2_2" & i).Activate
        [AD1] = "Этап 2"
        [AD2] = "202" & i
        'статус бар
        Application.StatusBar = "Выставление настроек листов. Четвёртый диапазон. " & _
        "Выполнено: " & Int(100 * i / 4) & "%."
  Next i

ThisWorkbook.Sheets("Preferences").Activate
Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
    
End Sub
Sub Clone7()
  Dim kolvo As Variant
  Dim i As Long
  Dim list As Worksheet
  Dim ThisWorkbook As Workbook
  Dim SheetName, var As String
  Dim Sht As Worksheet
  Dim WorkRng As Range
  
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual
  
  Set ThisWorkbook = ActiveWorkbook
  SheetName = "7"
  kolvo = 15
  
  'удаление предыдущих данных

  On Error Resume Next
  
  For i = 1 To kolvo
    Sheets("7" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Первий диапазон. " & _
    "Выполнено: " & Int(100 * i / kolvo) & "%."
  Next i
  
  For i = 1 To 5
    Sheets("7_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Второй диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = 1 To 2
    Sheets("7_" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Третий диапазон. " & _
    "Выполнено: " & Int(100 * i / 2) & "%."
  Next i
  
  For i = 1 To 5
    Sheets("7_1_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Четвёртый диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = 1 To 5
    Sheets("7_2_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Пятый диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
 
  'копирование листов
  ThisWorkbook.Sheets(SheetName).Activate
  Set list = ActiveSheet
'kolvo = InputBox("Укажите необходимое количество листов")

'If kolvo = "" Then Exit Sub
'If IsNumeric(kolvo) Then
'    kolvo = Fix(kolvo)
    For i = 2 To kolvo
        list.Copy After:=ActiveSheet
        ActiveSheet.Name = list.Name & i
        'статус бар
        Application.StatusBar = "Копирование листов. " & _
        "Выполнено: " & Int(100 * i / kolvo) & "%."
    Next
'Else
'    MsgBox "Неправильное число"
'End If

'переименование листов
On Error Resume Next
For Each Sht In Worksheets
    'статус бар
    Application.StatusBar = "Переименование листов."
    If Sht.Name = "72" Then Sht.Name = "7_22"
    If Sht.Name = "73" Then Sht.Name = "7_23"
    If Sht.Name = "74" Then Sht.Name = "7_24"
    If Sht.Name = "75" Then Sht.Name = "7_25"
    If Sht.Name = "76" Then Sht.Name = "7_1"
    If Sht.Name = "77" Then Sht.Name = "7_1_22"
    If Sht.Name = "78" Then Sht.Name = "7_1_23"
    If Sht.Name = "79" Then Sht.Name = "7_1_24"
    If Sht.Name = "710" Then Sht.Name = "7_1_25"
    If Sht.Name = "711" Then Sht.Name = "7_2"
    If Sht.Name = "712" Then Sht.Name = "7_2_22"
    If Sht.Name = "713" Then Sht.Name = "7_2_23"
    If Sht.Name = "714" Then Sht.Name = "7_2_24"
    If Sht.Name = "715" Then Sht.Name = "7_2_25"
Next

'выставление настроек
  On Error Resume Next
  
  'года
  For i = 2 To 5
    Sheets("7_2" & i).Activate
        [L3] = "202" & i

    'статус бар
    Application.StatusBar = "Выставление настроек листов. Первий диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  'этапы
  For i = 1 To 2
    Sheets("7_" & i).Activate
        [L2] = "Этап " & i
        'статус бар
        Application.StatusBar = "Выставление настроек листов. Второй диапазон. " & _
        "Выполнено: " & Int(100 * i / 2) & "%."
  Next i
  
  'Этап 1
  For i = 2 To 5
    Sheets("7_1_2" & i).Activate
        [L2] = "Этап 1"
        [L3] = "202" & i
        'статус бар
        Application.StatusBar = "Выставление настроек листов. Третий диапазон. " & _
        "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  'Этап 2
  For i = 2 To 5
    Sheets("7_2_2" & i).Activate
        [L2] = "Этап 2"
        [L3] = "202" & i
        'статус бар
        Application.StatusBar = "Выставление настроек листов. Четвёртый диапазон. " & _
        "Выполнено: " & Int(100 * i / 4) & "%."
  Next i

ThisWorkbook.Sheets("Preferences").Activate
Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
    
End Sub
Sub Clone9()
  Dim kolvo As Variant
  Dim i As Long
  Dim list As Worksheet
  Dim ThisWorkbook As Workbook
  Dim SheetName, var As String
  Dim Sht As Worksheet
  Dim WorkRng As Range
  
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual
  
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
  
  For i = 3 To 6
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
  
  For i = 3 To 6
    Sheets("9_1_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Четвёртый диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = 3 To 6
    Sheets("9_2_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Пятый диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
 
  'копирование листов
  ThisWorkbook.Sheets(SheetName).Activate
  Set list = ActiveSheet
'kolvo = InputBox("Укажите необходимое количество листов")

'If kolvo = "" Then Exit Sub
'If IsNumeric(kolvo) Then
'    kolvo = Fix(kolvo)
    For i = 1 To kolvo
        list.Copy After:=ActiveSheet
        ActiveSheet.Name = list.Name & i
        'статус бар
        Application.StatusBar = "Копирование листов. " & _
        "Выполнено: " & Int(100 * i / kolvo) & "%."
    Next
'Else
'    MsgBox "Неправильное число"
'End If

'переименование листов
On Error Resume Next
For Each Sht In Worksheets
    'статус бар
    Application.StatusBar = "Переименование листов."
    If Sht.Name = "91" Then Sht.Name = "9_23"
    If Sht.Name = "92" Then Sht.Name = "9_24"
    If Sht.Name = "93" Then Sht.Name = "9_25"
    If Sht.Name = "94" Then Sht.Name = "9_26"
    If Sht.Name = "95" Then Sht.Name = "9_1"
    If Sht.Name = "96" Then Sht.Name = "9_1_23"
    If Sht.Name = "97" Then Sht.Name = "9_1_24"
    If Sht.Name = "98" Then Sht.Name = "9_1_25"
    If Sht.Name = "99" Then Sht.Name = "9_1_26"
    If Sht.Name = "910" Then Sht.Name = "9_2"
    If Sht.Name = "911" Then Sht.Name = "9_2_23"
    If Sht.Name = "912" Then Sht.Name = "9_2_24"
    If Sht.Name = "913" Then Sht.Name = "9_2_25"
    If Sht.Name = "914" Then Sht.Name = "9_2_26"
Next

'выставление настроек
  On Error Resume Next
  
  'года
  For i = 3 To 6
    Sheets("9_2" & i).Activate
        [O2] = "202" & i
        Range("Z:AI").Clear
        
'        var = Int("202" & i)
'    For j = 210 To 13 Step -1
'        If Range("S" & j).Value2 <> var Then
'            Range("S" & j).EntireRow.delete
'        End If
'    Next j

    'статус бар
    Application.StatusBar = "Выставление настроек листов. Первий диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  'этапы
  For i = 1 To 2
    Sheets("9_" & i).Activate
        [O1] = "Этап " & i
        Range("Z:AI").Clear
        'статус бар
        Application.StatusBar = "Выставление настроек листов. Второй диапазон. " & _
        "Выполнено: " & Int(100 * i / 2) & "%."
  Next i
  
  'Этап 1
  For i = 3 To 6
    Sheets("9_1_2" & i).Activate
        [O1] = "Этап 1"
        [O2] = "202" & i
        Range("Z:AI").Clear
        'статус бар
        Application.StatusBar = "Выставление настроек листов. Третий диапазон. " & _
        "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  'Этап 2
  For i = 3 To 6
    Sheets("9_2_2" & i).Activate
        [O1] = "Этап 2"
        [O2] = "202" & i
        Range("Z:AI").Clear
        'статус бар
        Application.StatusBar = "Выставление настроек листов. Четвёртый диапазон. " & _
        "Выполнено: " & Int(100 * i / 4) & "%."
  Next i

ThisWorkbook.Sheets("Preferences").Activate
Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
    
End Sub
Sub Clone20()
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
Application.Calculation = xlManual
  
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
  
  For i = 3 To 6
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
  
  For i = 3 To 6
    Sheets("20_1_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Четвёртый диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = 3 To 6
    Sheets("20_2_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Пятый диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
 
  'копирвоание листов
  ThisWorkbook.Sheets(SheetName).Activate
  Set list = ActiveSheet
'kolvo = InputBox("Укажите необходимое количество листов")

'If kolvo = "" Then Exit Sub
'If IsNumeric(kolvo) Then
'    kolvo = Fix(kolvo)
    For i = 1 To kolvo
        list.Copy After:=ActiveSheet
        ActiveSheet.Name = list.Name & i
        'статус бар
        Application.StatusBar = "Копирование листов. " & _
        "Выполнено: " & Int(100 * i / kolvo) & "%."
    Next
'Else
'    MsgBox "Неправильное число"
'End If

'переименование листов
On Error Resume Next
For Each Sht In Worksheets
    'статус бар
    Application.StatusBar = "Переименование листов."
    If Sht.Name = "201" Then Sht.Name = "20_23"
    If Sht.Name = "202" Then Sht.Name = "20_24"
    If Sht.Name = "203" Then Sht.Name = "20_25"
    If Sht.Name = "204" Then Sht.Name = "20_26"
    If Sht.Name = "205" Then Sht.Name = "20_1"
    If Sht.Name = "206" Then Sht.Name = "20_1_23"
    If Sht.Name = "207" Then Sht.Name = "20_1_24"
    If Sht.Name = "208" Then Sht.Name = "20_1_25"
    If Sht.Name = "209" Then Sht.Name = "20_1_26"
    If Sht.Name = "2010" Then Sht.Name = "20_2"
    If Sht.Name = "2011" Then Sht.Name = "20_2_23"
    If Sht.Name = "2012" Then Sht.Name = "20_2_24"
    If Sht.Name = "2013" Then Sht.Name = "20_2_25"
    If Sht.Name = "2014" Then Sht.Name = "20_2_26"
Next

'выставление настроек
  On Error Resume Next
  
  For i = 3 To 6
    Sheets("20_2" & i).Activate
        [H2] = "202" & i
        [C59] = "2_2" & i
        [D59] = "2_2" & i
        Range("K3:N44").Clear
        'статус бар
        Application.StatusBar = "Выставление настроек листов. Первий диапазон. " & _
        "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = 1 To 2
    Sheets("20_" & i).Activate
        [H1] = "Этап " & i
        [C59] = "2_" & i
        [D59] = "2_" & i
        Range("K3:N44").Clear
        'статус бар
        Application.StatusBar = "Выставление настроек листов. Второй диапазон. " & _
        "Выполнено: " & Int(100 * i / 2) & "%."
  Next i
  
  For i = 3 To 6
    Sheets("20_1_2" & i).Activate
        [H1] = "Этап 1"
        [H2] = "202" & i
        [C59] = "2_1_2" & i
        [D59] = "2_1_2" & i
        Range("K3:N44").Clear
        'статус бар
        Application.StatusBar = "Выставление настроек листов. Третий диапазон. " & _
        "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = 3 To 6
    Sheets("20_2_2" & i).Activate
        [H1] = "Этап 2"
        [H2] = "202" & i
        [C59] = "2_2_2" & i
        [D59] = "2_2_2" & i
        Range("K3:N44").Clear
        'статус бар
        Application.StatusBar = "Выставление настроек листов. Четвёртый диапазон. " & _
        "Выполнено: " & Int(100 * i / 4) & "%."
  Next i

ThisWorkbook.Sheets("Preferences").Activate
Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
    
End Sub


Sub Clone4()
    Dim kolvo As Long
    Dim i As Long
    Dim list As Worksheet
    Dim ThisWorkbook As Workbook
    Dim SheetName As String
    Dim Sht As Worksheet

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayAlerts = False
    Application.Calculation = xlManual

    Set ThisWorkbook = ActiveWorkbook
    SheetName = "4"
    kolvo = 15

    ' Удаление существующих листов
    On Error Resume Next
    For i = 23 To 26
        Sheets("4_" & i).delete
        Application.StatusBar = "Удаление листов."
    Next i
    For i = 23 To 26
        Sheets("4_1_" & i).delete
        Application.StatusBar = "Удаление листов."
    Next i
    For i = 23 To 26
        Sheets("4_2_" & i).delete
        Application.StatusBar = "Удаление листов."
    Next i
    Sheets("4_1").delete
    Sheets("4_2").delete

    ' Клонирование листов
    ThisWorkbook.Sheets(SheetName).Activate
    Set list = ActiveSheet

    For i = 1 To kolvo - 1
        list.Copy After:=ActiveSheet
        ActiveSheet.Name = list.Name & i
        Application.StatusBar = "Клонирование листов. Выполнено: " & Int(100 * i / (kolvo - 1)) & "%."
    Next

    ' Переименование листов
    On Error Resume Next
    For Each Sht In Worksheets
        Application.StatusBar = "Переименование листов."
        Select Case Sht.Name
            Case "41": Sht.Name = "4_23"
            Case "42": Sht.Name = "4_24"
            Case "43": Sht.Name = "4_25"
            Case "44": Sht.Name = "4_26"
            Case "45": Sht.Name = "4_1"
            Case "46": Sht.Name = "4_1_23"
            Case "47": Sht.Name = "4_1_24"
            Case "48": Sht.Name = "4_1_25"
            Case "49": Sht.Name = "4_1_26"
            Case "410": Sht.Name = "4_2"
            Case "411": Sht.Name = "4_2_23"
            Case "412": Sht.Name = "4_2_24"
            Case "413": Sht.Name = "4_2_25"
            Case "414": Sht.Name = "4_2_26"
        End Select
    Next

    ' Обновление значений
    On Error Resume Next
    For i = 23 To 26
        Sheets("4_" & i).Activate
        [AH4] = "20" & i
        Application.StatusBar = "Обновление значений листов. Выполнено: " & Int(100 * (i - 23) / (kolvo - 1)) & "%."
    Next i
    For i = 23 To 26
        Sheets("4_1" & i).Activate
        [AH4] = "20" & i
        Application.StatusBar = "Обновление значений листов. Выполнено: " & Int(100 * (i - 23) / (kolvo - 1)) & "%."
    Next i
    For i = 23 To 26
        Sheets("4_2" & i).Activate
        [AH4] = "20" & i
        Application.StatusBar = "Обновление значений листов. Выполнено: " & Int(100 * (i - 23) / (kolvo - 1)) & "%."
    Next i
    

    ThisWorkbook.Sheets("Preferences").Activate
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayAlerts = True
    Application.Calculation = xlAutomatic
End Sub
