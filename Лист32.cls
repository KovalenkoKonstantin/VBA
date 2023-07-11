VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub CheckBox_Click()
 Dim ThisWorkbook As Workbook
 Dim st As String
 Dim CheckBoxName As String
 Dim CheckBoxObject As Variant
 Set ThisWorkbook = ActiveWorkbook
 ThisWorkbook.Sheets("Preferences").Activate
 On Error GoTo ExitHandler
 
 Application.ScreenUpdating = False
 Application.EnableEvents = False
 ActiveSheet.DisplayPageBreaks = False
 Application.DisplayAlerts = False
 
' ищем номер колонки по "Содержание"
For I = 1 To 50
    If Worksheets("Preferences").Cells(1, I) = "Содержание" Then
        column = I
    End If
Next

' всего чекбоксов
'K = 50
K = ActiveSheet.OLEObjects.Count

'номер цвета
a = 13434879 'бледно-желтый
b = 6750105 'насыщенный зелёный
C = 16759277 'бледно-фиолетовый
d = 13434828 'салатовый

'выбор цвета вкладки
Y = a

'цикл проверки
    For J = 2 To K 'проверка начинается со второго чекбокса
'        If TypeOf j Is MSForms.CheckBox And Not TypeOf j Is MSForms.OptionButton Then
'статус бар
Application.StatusBar = "Выполнено: " & Int((J / K) * 100) & " %"

        CheckBoxName = "CheckBox" & CStr(J) 'кастую к стринге номер чекбокса
        Set CheckBoxObject = Me.Evaluate(CheckBoxName) 'Me работает только на листе
        
        st = Cells(J, column).Text
        
            If CheckBoxObject.Object.Value = True Then
    
    '             If CheckBox1.Value = True Then
                Sheets(st).Visible = True
                Sheets(st).Select
                With ActiveWorkbook.Sheets(st).Tab
                    .Color = Y
                    .TintAndShade = 0
                End With
    '            Sheets(st).Select
    '            Sheets(st).Range("AZ1") = "+"
                Sheets("Preferences").Select
            Else
                Sheets(st).Visible = True
                Sheets(st).Select
                With ActiveWorkbook.Sheets(st).Tab
                    .ColorIndex = xlNone
                    .TintAndShade = 0
                End With
    '            Sheets(st).Select
    '            Sheets(st).Range("AZ1") = emplty
                Sheets(st).Visible = False
                Sheets("Preferences").Select
            End If
'        End If
    Next J
    
'    'сообщение
'MsgBoxEx "Всё исправлено", 0, "еее", 2
    
ExitHandler:
       Application.StatusBar = False
       Application.ScreenUpdating = True
       Application.EnableEvents = True
       ActiveSheet.DisplayPageBreaks = True
       Application.DisplayStatusBar = True
       Application.DisplayAlerts = True
    Exit Sub
    
End Sub

Private Sub CheckBox1_Click()
 Dim ThisWorkbook As Workbook
 Set ThisWorkbook = ActiveWorkbook
 ThisWorkbook.Sheets("Preferences").Activate
 
 CheckBoxName = "CheckBox1"
        Set CheckBoxObject = Me.Evaluate(CheckBoxName)
        If CheckBoxObject.Object.Value = True Then
            ThisWorkbook.Sheets("Preferences").Range("W2:W100") = True
        ElseIf CheckBoxObject.Object.Value = False Then
            ThisWorkbook.Sheets("Preferences").Range("W2:W100") = False
        End If
 
End Sub

Private Sub CheckBox34_Click()

End Sub

Private Sub CheckBox66_Click()

End Sub

Private Sub CommandButton1_Click()
    UserForm1.Show
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

Dim row, column, X As Integer
On Error GoTo ErrHandler
 Application.ScreenUpdating = False
 Application.EnableEvents = False
 ActiveSheet.DisplayPageBreaks = False
 Application.DisplayStatusBar = False
 Application.DisplayAlerts = False

' ищем номер строки по "Количество этапов"
For I = 1 To 50
    If Worksheets("Preferences").Cells(I, 1) = "Количество этапов" Then
        row = I
    End If
Next

'колонка в котоую будет добавляться значение
column = 2

'If (Target.column = column) And (Target.row = row) And ActiveSheet.Range(Cells(Target.row, column)).Value Is Not Empty Then
'    GoTo ErrHandler
'End If

If (Target.column = column) And (Target.row = row) And Worksheets("Preferences").Cells(Target.row, column) <> "" Then
        X = Application.Worksheets("Preferences").Cells(Target.row, column).Value 'значение ячейки соответствующее количеству этапов
        Rows(Target.row + 1).Resize(X - 1).Insert 'вставляем столько колонок, сколько указано в ячейке минус 1
        
        'объединение ячеек
        For I = 1 To column
            ActiveSheet.Range(Cells(Target.row, I), Cells(Target.row + X - 1, I)).Select '(x-1) - количество вставленных строк
                With Selection
                    .Merge
                    .Font.Name = "Times New Roman"
'                    .WrapText = False
                    .MergeCells = True
                    .Borders.LineStyle = True
                End With
        Next
        
        'заполнение значениями добавленные строки
        For I = 1 To X
            ActiveSheet.Range(Cells(Target.row + I - 1, 3), Cells(Target.row + I - 1, 3)).Value = "Этап № " & I
            ActiveSheet.Range(Cells(Target.row + I - 1, 3), Cells(Target.row + I - 1, 3)).Select
            With Selection
                .Borders.LineStyle = True
            End With
        Next
End If
ErrHandler:
'Application.Worksheets("Preferences").Cells(row, column).Activate
 Application.ScreenUpdating = True
 Application.EnableEvents = True
 ActiveSheet.DisplayPageBreaks = True
 Application.DisplayStatusBar = True
 Application.DisplayAlerts = True
End Sub

'Sub SaveToPDF()
'
' Dim array_distinct() 'задекларируем динамический массив
' Dim ThisWorkbook As Workbook
' Dim st, CheckBoxName, SaveName, Folder, Path As String
' Dim CheckBoxObject As Variant
' Dim sName As String
' Set ThisWorkbook = ActiveWorkbook
'
' On Error GoTo ExitHandler
'
' ThisWorkbook.Sheets("Preferences").Activate
' SaveName = ActiveSheet.Range("R29").Text
' ThisWorkbook.Activate
' Path = ActiveWorkbook.Path
'
' Application.ScreenUpdating = False
' Application.EnableEvents = False
' ActiveSheet.DisplayPageBreaks = False
' Application.DisplayStatusBar = False
' Application.DisplayAlerts = False
'
'' ищем номер колонки по "Содержание"
''For i = 1 To 50
''    If Worksheets("Preferences").Cells(1, i) = "Содержание" Then
''        column = i
''    End If
''Next
'
'' ищем номер последнего заполненного ряда по колонке "Содержание"
'' ThisWorkbook.Sheets("Preferences").Activate
'' lLastRow = Cells(Rows.Count, column).End(xlUp).row
'
''добавление выделенных листов в массив
''On Error Resume Next
''k = -1
''sName = ""
''    For i = 1 To lLastRow - 1 'количество checkbox меньше числа записей на 1
''
''        CheckBoxName = "CheckBox" & CStr(i)
''        Set CheckBoxObject = Me.Evaluate(CheckBoxName)
''        If CheckBoxObject.Object.Value = True Then
''
''            k = k + 1 'динамически устанавливаем вертикальный размер массива
''            ReDim array_distinct(k, 0) 'переопределяем размер массива
''            array_distinct(k, 0) = Cells(i + 1, column).Value2 'добавляем значение в массив
''            sName = Chr(34) & array_distinct(k, 0) & Chr(34) & ", " & sName
''        End If
''    Next i
'
''сохраняю под заданным именем
''sName = Right(sName, Len(sName) - 1)
''sName = Left(sName, Len(sName) - 2)
'
''ThisWorkbook.Sheets(Array(sName)).Select
'Sheets(Array("2_1", "6", "7_1", "9_1э" _
'        , "22ф", "Приказ", "П5", "КУЗ")).Select
''Sheets("1").Activate
'ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
'Filename:=Path & "\" & _
'SaveName & ".pdf", Quality:=xlQualityStandard _
', IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
'False
'
''Sheets(Array("Сопр", "Опись", "ПЗ", "1", "2", "2_23", "2_1_23", "2_2_23", "9", "10", "12" _
''        , "20", "20_1", "20_2", "21ф", "22ф", "23ф", "Приказ", "П6", "П8")).Select
''    Sheets("Сопр").Activate
''    ChDir "D:\РКМ\ОРЦ\Улей-23"
''    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
''        "D:\РКМ\ОРЦ\Улей-23\ОРЦ Улей-23 работа_v1.1.pdf", Quality:=xlQualityStandard _
''        , IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
''        True
'
'ExitHandler:
'    Application.ScreenUpdating = True
'    Application.EnableEvents = True
'    ActiveSheet.DisplayPageBreaks = True
'    Application.DisplayStatusBar = True
'    Application.DisplayAlerts = True
'    ThisWorkbook.Sheets("Preferences").Activate
'Exit Sub
'
'End Sub
