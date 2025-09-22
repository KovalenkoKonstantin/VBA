Attribute VB_Name = "CopyList"
Sub Clone2()
  Dim kolvo As Variant
  Dim i As Long
  Dim list As Worksheet
  Dim ThisWorkbook As Workbook
  Dim SheetName As String
  Dim Sht As Worksheet
  Dim Start As Integer
  Dim Finish As Integer
  
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual
  
  Set ThisWorkbook = ActiveWorkbook
  SheetName = "2"
  kolvo = 14
  Start = 4
  Finish = 7
  
  'удаление предыдущих данных
  On Error Resume Next
  
  For i = 1 To kolvo
    Sheets("2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Первий диапазон. " & _
    "Выполнено: " & Int(100 * i / kolvo) & "%."
  Next i
  
  For i = Start To Finish
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
  
  For i = Start To Finish
    Sheets("2_1_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Четвёртый диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = Start To Finish
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
'статус бар
Application.StatusBar = "Переименование листов."
For Each Sht In ThisWorkbook.Worksheets
    Select Case Sht.Name
        Case "21":  Sht.Name = "2_2" & CStr(Start)
        Case "22":  Sht.Name = "2_2" & CStr(Start + 1)
        Case "23":  Sht.Name = "2_2" & CStr(Start + 2)
        Case "24":  Sht.Name = "2_2" & CStr(Start + 3)

        Case "25":  Sht.Name = "2_1"

        Case "26":  Sht.Name = "2_1_2" & CStr(Start)
        Case "27":  Sht.Name = "2_1_2" & CStr(Start + 1)
        Case "28":  Sht.Name = "2_1_2" & CStr(Start + 2)
        Case "29":  Sht.Name = "2_1_2" & CStr(Start + 3)

        Case "210": Sht.Name = "2_2"
        
        Case "211": Sht.Name = "2_2_2" & CStr(Start)
        Case "212": Sht.Name = "2_2_2" & CStr(Start + 1)
        Case "213": Sht.Name = "2_2_2" & CStr(Start + 2)
        Case "214": Sht.Name = "2_2_2" & CStr(Start + 3)
    End Select
Next Sht

'выставление настроек
  On Error Resume Next
  
  'года
  For i = Start To Finish
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
  For i = Start To Finish
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
  For i = Start To Finish
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
  Dim Start As Integer
  Dim Finish As Integer
  
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual
  
  Set ThisWorkbook = ActiveWorkbook
  SheetName = "6"
  kolvo = 15
  Start = 4
  Finish = 7
  
  'удаление предыдущих данных

  On Error Resume Next
  
  For i = 1 To kolvo
    Sheets("6" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Первий диапазон. " & _
    "Выполнено: " & Int(100 * i / kolvo) & "%."
  Next i
  
  For i = Start To Finish
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
  
  For i = Start To Finish
    Sheets("6_1_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Четвёртый диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = Start To Finish
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
  For i = Start To Finish
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
  For i = Start To Finish
    Sheets("6_1_2" & i).Activate
        [AD1] = "Этап 1"
        [AD2] = "202" & i
        'статус бар
        Application.StatusBar = "Выставление настроек листов. Третий диапазон. " & _
        "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  'Этап 2
  For i = Start To Finish
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
  Dim Start As Integer
  Dim Finish As Integer
  
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual
  
  Set ThisWorkbook = ActiveWorkbook
  SheetName = "7"
  kolvo = 15
  Start = 4
  Finish = 7
  
  'удаление предыдущих данных

  On Error Resume Next
  
  For i = 1 To kolvo
    Sheets("7" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Первий диапазон. " & _
    "Выполнено: " & Int(100 * i / kolvo) & "%."
  Next i
  
  For i = Start To Finish
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
  
  For i = Start To Finish
    Sheets("7_1_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Четвёртый диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = Start To Finish
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
'статус бар
Application.StatusBar = "Переименование листов."

For Each Sht In ThisWorkbook.Worksheets
    Select Case Sht.Name
        Case "71":  Sht.Name = "7_2" & CStr(Start)
        Case "72":  Sht.Name = "7_2" & CStr(Start + 1)
        Case "73":  Sht.Name = "7_2" & CStr(Start + 2)
        Case "74":  Sht.Name = "7_2" & CStr(Start + 3)

        Case "75":  Sht.Name = "7_1"

        Case "76":  Sht.Name = "7_1_2" & CStr(Start)
        Case "77":  Sht.Name = "7_1_2" & CStr(Start + 1)
        Case "78":  Sht.Name = "7_1_2" & CStr(Start + 2)
        Case "79":  Sht.Name = "7_1_2" & CStr(Start + 3)

        Case "710": Sht.Name = "7_2"
        Case "711": Sht.Name = "7_2_2" & CStr(Start)
        Case "712": Sht.Name = "7_2_2" & CStr(Start + 1)
        Case "713": Sht.Name = "7_2_2" & CStr(Start + 2)
        Case "714": Sht.Name = "7_2_2" & CStr(Start + 3)
    End Select
Next Sht

'выставление настроек
  On Error Resume Next
  
  'года
  For i = Start To Finish
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
  For i = Start To Finish
    Sheets("7_1_2" & i).Activate
        [L2] = "Этап 1"
        [L3] = "202" & i
        'статус бар
        Application.StatusBar = "Выставление настроек листов. Третий диапазон. " & _
        "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  'Этап 2
  For i = Start To Finish
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
  Dim Start As Integer
  Dim Finish As Integer
  
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual
  
  Set ThisWorkbook = ActiveWorkbook
  SheetName = "9"
  kolvo = 14
  Start = 4
  Finish = 7
  
  'удаление предыдущих данных
  On Error Resume Next
  
  For i = 1 To kolvo
    Sheets("9" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Первий диапазон. " & _
    "Выполнено: " & Int(100 * i / kolvo) & "%."
  Next i
  
  For i = Start To Finish
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
  
  For i = Start To Finish
    Sheets("9_1_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Четвёртый диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = Start To Finish
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
'статус бар
Application.StatusBar = "Переименование листов."
    
For Each Sht In ThisWorkbook.Worksheets
    Select Case Sht.Name
        Case "91":  Sht.Name = "9_2" & CStr(Start)
        Case "92":  Sht.Name = "9_2" & CStr(Start + 1)
        Case "93":  Sht.Name = "9_2" & CStr(Start + 2)
        Case "94":  Sht.Name = "9_2" & CStr(Start + 3)

        Case "95":  Sht.Name = "9_1"

        Case "96":  Sht.Name = "9_1_2" & CStr(Start)
        Case "97":  Sht.Name = "9_1_2" & CStr(Start + 1)
        Case "98":  Sht.Name = "9_1_2" & CStr(Start + 2)
        Case "99":  Sht.Name = "9_1_2" & CStr(Start + 3)

        Case "910": Sht.Name = "9_2"
        Case "911": Sht.Name = "9_2_2" & CStr(Start)
        Case "912": Sht.Name = "9_2_2" & CStr(Start + 1)
        Case "913": Sht.Name = "9_2_2" & CStr(Start + 2)
        Case "914": Sht.Name = "9_2_2" & CStr(Start + 3)
    End Select
Next Sht

'выставление настроек
  On Error Resume Next
  
  'года
  For i = Start To Finish
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
  For i = Start To Finish
    Sheets("9_1_2" & i).Activate
        [O1] = "Этап 1"
        [O2] = "202" & i
        Range("Z:AI").Clear
        'статус бар
        Application.StatusBar = "Выставление настроек листов. Третий диапазон. " & _
        "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  'Этап 2
  For i = Start To Finish
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
  Dim Start As Integer
  Dim Finish As Integer
  
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual
  
  Set ThisWorkbook = ActiveWorkbook
  SheetName = "20"
  kolvo = 14
  Start = 4
  Finish = 7
  
  'удаление предыдущих данных
  
  

  On Error Resume Next
  
  For i = 1 To kolvo
    Sheets("20" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Первий диапазон. " & _
    "Выполнено: " & Int(100 * i / kolvo) & "%."
  Next i
  
  For i = Start To Finish
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
  
  For i = Start To Finish
    Sheets("20_1_2" & i).delete
    'статус бар
    Application.StatusBar = "Удаление листов. Четвёртый диапазон. " & _
    "Выполнено: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = Start To Finish
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

'статус бар
Application.StatusBar = "Переименование листов."


For Each Sht In ThisWorkbook.Worksheets
    Select Case Sht.Name
        Case "201":  Sht.Name = "20_2" & CStr(Start)
        Case "202":  Sht.Name = "20_2" & CStr(Start + 1)
        Case "203":  Sht.Name = "20_2" & CStr(Start + 2)
        Case "204":  Sht.Name = "20_2" & CStr(Start + 3)

        Case "205":  Sht.Name = "20_1"

        Case "206":  Sht.Name = "20_1_2" & CStr(Start)
        Case "207":  Sht.Name = "20_1_2" & CStr(Start + 1)
        Case "208":  Sht.Name = "20_1_2" & CStr(Start + 2)
        Case "209":  Sht.Name = "20_1_2" & CStr(Start + 3)

        Case "2010": Sht.Name = "20_2"
        Case "2011": Sht.Name = "20_2_2" & CStr(Start)
        Case "2012": Sht.Name = "20_2_2" & CStr(Start + 1)
        Case "2013": Sht.Name = "20_2_2" & CStr(Start + 2)
        Case "2014": Sht.Name = "20_2_2" & CStr(Start + 3)
    End Select
Next Sht

'выставление настроек
  On Error Resume Next
  
  For i = Start To Finish
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
  
  For i = Start To Finish
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
  
  For i = Start To Finish
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
    Dim Start As Integer
    Dim Finish As Integer

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayAlerts = False
    Application.Calculation = xlManual

    Set ThisWorkbook = ActiveWorkbook
    SheetName = "4"
    kolvo = 15
    Start = 4
    Finish = 7

    ' Удаление существующих листов
    On Error Resume Next
    For i = 2 & CStr(Start) To 2 & CStr(Start + 3)
        Sheets("4_" & i).delete
        Application.StatusBar = "Удаление листов."
    Next i
    For i = 2 & CStr(Start) To 2 & CStr(Start + 3)
        Sheets("4_1_" & i).delete
        Application.StatusBar = "Удаление листов."
    Next i
    For i = 2 & CStr(Start) To 2 & CStr(Start + 3)
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
    For i = 2 & CStr(Start) To 2 & CStr(Start + 3)
        Sheets("4_" & i).Activate
        [AH4] = "20" & i
        Application.StatusBar = "Обновление значений листов. Выполнено: " & Int(100 * (i - 23) / (kolvo - 1)) & "%."
    Next i
    For i = 2 & CStr(Start) To 2 & CStr(Start + 3)
        Sheets("4_1" & i).Activate
        [AH4] = "20" & i
        Application.StatusBar = "Обновление значений листов. Выполнено: " & Int(100 * (i - 23) / (kolvo - 1)) & "%."
    Next i
    For i = 2 & CStr(Start) To 2 & CStr(Start + 3)
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
