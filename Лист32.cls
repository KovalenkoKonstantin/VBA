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
 On Error Resume Next
 
 Application.ScreenUpdating = False
 Application.EnableEvents = False
 ActiveSheet.DisplayPageBreaks = False
 Application.DisplayAlerts = False
 Application.Calculation = xlManual
 
' ищем номер колонки по "Содержание"
For i = 1 To 150
    If Worksheets("Preferences").Cells(1, i) = "Содержание" Then
        column = i
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
    For j = 2 To K 'проверка начинается со второго чекбокса
'        If TypeOf j Is MSForms.CheckBox And Not TypeOf j Is MSForms.OptionButton Then
'статус бар
Application.StatusBar = "Выполнено: " & Int((j / K) * 100) & " %"

        CheckBoxName = "CheckBox" & CStr(j) 'кастую к стринге номер чекбокса
        Set CheckBoxObject = Me.Evaluate(CheckBoxName) 'Me работает только на листе
        
        st = Cells(j, column).Text
        
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
    Next j
    
'    'сообщение
'MsgBoxEx "Всё исправлено", 0, "еее", 2
    
ExitHandler:
       Application.StatusBar = False
       Application.ScreenUpdating = True
       Application.EnableEvents = True
       ActiveSheet.DisplayPageBreaks = True
       Application.DisplayStatusBar = True
       Application.DisplayAlerts = True
       Application.Calculation = xlAutomatic
    Exit Sub
    
End Sub

Private Sub CheckBox1_Click()
 Dim ThisWorkbook As Workbook
 Set ThisWorkbook = ActiveWorkbook
 ThisWorkbook.Sheets("Preferences").Activate
 
 CheckBoxName = "CheckBox1"
        Set CheckBoxObject = Me.Evaluate(CheckBoxName)
        If CheckBoxObject.Object.Value = True Then
            ThisWorkbook.Sheets("Preferences").Range("X2:X100") = True
        ElseIf CheckBoxObject.Object.Value = False Then
            ThisWorkbook.Sheets("Preferences").Range("X2:X100") = False
        End If
 
End Sub



Private Sub CheckBox24_Click()

End Sub

Private Sub CheckBox25_Click()

End Sub

Private Sub CheckBox62_Click()

End Sub

Private Sub CheckBox77_Click()

End Sub

Private Sub CommandButton1_Click()
    UserForm11.Show
End Sub

