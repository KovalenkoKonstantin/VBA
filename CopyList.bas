Attribute VB_Name = "CopyList"
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
 
  'копирование листов
  ThisWorkbook.Sheets(SheetName).Activate
  Set list = ActiveSheet
'kolvo = InputBox("Укажите необходимое количество листов")

'If kolvo = "" Then Exit Sub
'If IsNumeric(kolvo) Then
'    kolvo = Fix(kolvo)
    For i = 1 To kolvo
        list.Copy after:=ActiveSheet
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
    If Sht.Name = "91" Then Sht.Name = "9_21"
    If Sht.Name = "92" Then Sht.Name = "9_22"
    If Sht.Name = "93" Then Sht.Name = "9_23"
    If Sht.Name = "94" Then Sht.Name = "9_24"
    If Sht.Name = "95" Then Sht.Name = "9_1"
    If Sht.Name = "96" Then Sht.Name = "9_1_21"
    If Sht.Name = "97" Then Sht.Name = "9_1_22"
    If Sht.Name = "98" Then Sht.Name = "9_1_23"
    If Sht.Name = "99" Then Sht.Name = "9_1_24"
    If Sht.Name = "910" Then Sht.Name = "9_2"
    If Sht.Name = "911" Then Sht.Name = "9_2_21"
    If Sht.Name = "912" Then Sht.Name = "9_2_22"
    If Sht.Name = "913" Then Sht.Name = "9_2_23"
    If Sht.Name = "914" Then Sht.Name = "9_2_24"
Next

'выставление настроек
  On Error Resume Next
  
  'года
  For i = 1 To 4
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
  For i = 1 To 4
    Sheets("9_1_2" & i).Activate
        [O1] = "Этап 1"
        [O2] = "202" & i
        Range("Z:AI").Clear
        'статус бар
        Application.StatusBar = "Выставление настроек листов. Третий диапазон. " & _
        "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  'Этап 2
  For i = 1 To 4
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
 
  'копирвоание листов
  ThisWorkbook.Sheets(SheetName).Activate
  Set list = ActiveSheet
'kolvo = InputBox("Укажите необходимое количество листов")

'If kolvo = "" Then Exit Sub
'If IsNumeric(kolvo) Then
'    kolvo = Fix(kolvo)
    For i = 1 To kolvo
        list.Copy after:=ActiveSheet
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
    If Sht.Name = "21" Then Sht.Name = "2_21"
    If Sht.Name = "22" Then Sht.Name = "2_22"
    If Sht.Name = "23" Then Sht.Name = "2_23"
    If Sht.Name = "24" Then Sht.Name = "2_24"
    If Sht.Name = "25" Then Sht.Name = "2_1"
    If Sht.Name = "26" Then Sht.Name = "2_1_21"
    If Sht.Name = "27" Then Sht.Name = "2_1_22"
    If Sht.Name = "28" Then Sht.Name = "2_1_23"
    If Sht.Name = "29" Then Sht.Name = "2_1_24"
    If Sht.Name = "210" Then Sht.Name = "2_2"
    If Sht.Name = "211" Then Sht.Name = "2_2_21"
    If Sht.Name = "212" Then Sht.Name = "2_2_22"
    If Sht.Name = "213" Then Sht.Name = "2_2_23"
    If Sht.Name = "214" Then Sht.Name = "2_2_24"
Next

'выставление настроек
  On Error Resume Next
  
  'года
  For i = 1 To 4
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
  For i = 1 To 4
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
  For i = 1 To 4
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
  Sheets("2_1_21").Activate
  Range("F68:I68").ClearContents
  
  Sheets("2_1_22").Activate
  Range("G68:I68").ClearContents
  Range("E68:E68").ClearContents
  
  Sheets("2_1_23").Activate
  Range("E68:F68").ClearContents
  Range("I68:I68").ClearContents
  
  Sheets("2_1_24").Activate
  Range("E68:H68").ClearContents
  
'____________________________________________________________________
  
  Sheets("2_2_21").Activate
  Range("F68:I68").ClearContents
  
  Sheets("2_2_22").Activate
  Range("G68:I68").ClearContents
  Range("E68:E68").ClearContents
  
  Sheets("2_2_23").Activate
  Range("E68:F68").ClearContents
  Range("I68:I68").ClearContents
  
  Sheets("2_2_24").Activate
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
 
  'копирвоание листов
  ThisWorkbook.Sheets(SheetName).Activate
  Set list = ActiveSheet
'kolvo = InputBox("Укажите необходимое количество листов")

'If kolvo = "" Then Exit Sub
'If IsNumeric(kolvo) Then
'    kolvo = Fix(kolvo)
    For i = 1 To kolvo
        list.Copy after:=ActiveSheet
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
    If Sht.Name = "201" Then Sht.Name = "20_21"
    If Sht.Name = "202" Then Sht.Name = "20_22"
    If Sht.Name = "203" Then Sht.Name = "20_23"
    If Sht.Name = "204" Then Sht.Name = "20_24"
    If Sht.Name = "205" Then Sht.Name = "20_1"
    If Sht.Name = "206" Then Sht.Name = "20_1_21"
    If Sht.Name = "207" Then Sht.Name = "20_1_22"
    If Sht.Name = "208" Then Sht.Name = "20_1_23"
    If Sht.Name = "209" Then Sht.Name = "20_1_24"
    If Sht.Name = "2010" Then Sht.Name = "20_2"
    If Sht.Name = "2011" Then Sht.Name = "20_2_21"
    If Sht.Name = "2012" Then Sht.Name = "20_2_22"
    If Sht.Name = "2013" Then Sht.Name = "20_2_23"
    If Sht.Name = "2014" Then Sht.Name = "20_2_24"
Next

'выставление настроек
  On Error Resume Next
  
  For i = 1 To 4
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
  
  For i = 1 To 4
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
  
  For i = 1 To 4
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
