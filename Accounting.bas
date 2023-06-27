Attribute VB_Name = "Accounting"
Sub Data_insertion_26()
 
 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim SheetName As String
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 SheetName = "��.��26"
 awLastCol = 9
 Start = Now()
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

ThisWorkbook.Sheets(SheetName).Activate

 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xls), *.xls", _
 MultiSelect:=True, Title:="�������� ���� � �������� 26 �����")

 If TypeName(FilesToOpen) = "Boolean" Then
 GoTo ExitHandler
 End If
 
ThisWorkbook.Sheets(SheetName).Activate
On Error Resume Next
ActiveSheet.ShowAllData

'������� ������
Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

On Error Resume Next

 importWB.Sheets(1).Activate

'�������� ���������� ������
 ThisWorkbook.Sheets(SheetName).Activate
awLastRow = Cells(Rows.Count, "D").End(xlUp).row
'awLastCol = Cells(begin, Columns.Count).End(xlUp).Column

Range(Cells(1, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With

 '������� ������
 importWB.Sheets(1).Activate
 iwLastRow = Cells(Rows.Count, "D").End(xlUp).row
 importWB.Activate
 Range(Cells(1, 1), Cells(iwLastRow, awLastCol)).Copy
 
 ThisWorkbook.Sheets(SheetName).Activate
 Range(Cells(1, 1), Cells(iwLastRow, awLastCol)).Select
    With Selection
           .PasteSpecial Paste:=xlPasteAll
           .UnMerge
           .Font.Name = "Times New Roman"
           .WrapText = False
           .MergeCells = False
           .Font.Size = 8
    End With
    
'����������
importWB.Close
ThisWorkbook.Sheets(SheetName).Activate
MsgBox ("������ ������� ���������")

ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
 ThisWorkbook.Sheets(SheetName).Activate
' Finish = (Now() - Start) * 24 * 60 * 60
'MsgBox (Finish)
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub

Sub Data_insertion_44()
 
 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim SheetName As String
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 SheetName = "��.��44"
 awLastCol = 9
 Start = Now()
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

ThisWorkbook.Sheets(SheetName).Activate

 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xls), *.xls", _
 MultiSelect:=True, Title:="�������� ���� � �������� 44 �����")

 If TypeName(FilesToOpen) = "Boolean" Then
 GoTo ExitHandler
 End If
 
ThisWorkbook.Sheets(SheetName).Activate
On Error Resume Next
ActiveSheet.ShowAllData

'������� ������
Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

On Error Resume Next

 importWB.Sheets(1).Activate

'�������� ���������� ������
 ThisWorkbook.Sheets(SheetName).Activate
awLastRow = Cells(Rows.Count, "D").End(xlUp).row
'awLastCol = Cells(begin, Columns.Count).End(xlUp).Column

Range(Cells(1, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With

 '������� ������
 importWB.Sheets(1).Activate
 iwLastRow = Cells(Rows.Count, "D").End(xlUp).row
 importWB.Activate
 Range(Cells(1, 1), Cells(iwLastRow, awLastCol)).Copy
 
 ThisWorkbook.Sheets(SheetName).Activate
 Range(Cells(1, 1), Cells(iwLastRow, awLastCol)).Select
    With Selection
           .PasteSpecial Paste:=xlPasteAll
           .UnMerge
           .Font.Name = "Times New Roman"
           .WrapText = False
           .MergeCells = False
           .Font.Size = 8
    End With
    
'����������
importWB.Close
ThisWorkbook.Sheets(SheetName).Activate
MsgBox ("������ ������� ���������")

ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
 ThisWorkbook.Sheets(SheetName).Activate
' Finish = (Now() - Start) * 24 * 60 * 60
'MsgBox (Finish)
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub

Sub Data_insertion_OFR()
 
 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim SheetName As String
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 SheetName = "���"
 awLastCol = 20
 Start = Now()
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

ThisWorkbook.Sheets(SheetName).Activate

 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xls), *.xls", _
 MultiSelect:=True, Title:="�������� ���� � ������ ������ ���")

 If TypeName(FilesToOpen) = "Boolean" Then
 GoTo ExitHandler
 End If
 
ThisWorkbook.Sheets(SheetName).Activate
On Error Resume Next
ActiveSheet.ShowAllData

'������� ������
Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

On Error Resume Next

 importWB.Sheets(1).Activate

'�������� ���������� ������
 ThisWorkbook.Sheets(SheetName).Activate
awLastRow = Cells(Rows.Count, "N").End(xlUp).row
'awLastCol = Cells(begin, Columns.Count).End(xlUp).Column

Range(Cells(1, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With

 '������� ������
 importWB.Sheets(1).Activate
 iwLastRow = Cells(Rows.Count, "N").End(xlUp).row
 importWB.Activate
 Cells.Select
    Selection.NumberFormat = _
        "_-* #,##0.00 _?_-;-* #,##0.00 _?_-;_-* ""-""?? _?_-;_-@_-"
 Range(Cells(1, 1), Cells(iwLastRow, awLastCol)).Copy
 
 ThisWorkbook.Sheets(SheetName).Activate
 Range(Cells(1, 1), Cells(iwLastRow, awLastCol)).Select
    With Selection
           .PasteSpecial Paste:=xlPasteAll
           .UnMerge
           .Font.Name = "Times New Roman"
           .WrapText = False
           .MergeCells = False
           .Font.Size = 8
    End With
    
'���������� ������
If ThisWorkbook.Sheets(SheetName).Range("X1").Value2 = 2021 Then
    '�������
    ThisWorkbook.Sheets(SheetName).Range("AE2:AF6").Select
    Selection.ClearContents
    '2020
    ThisWorkbook.Sheets(SheetName).Range("AE2").Select
    ActiveCell.FormulaR1C1 = "=SUMIFS(C28,C21,RC29)"
    Range("AE2").Select
    Selection.AutoFill Destination:=Range("AE2:AE7"), Type:=xlFillDefault
    Range("AE2:AF7").Select
    Range("AE2:AF7").Copy
    Range("AE2:AF7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    '2021
    ThisWorkbook.Sheets(SheetName).Range("AF2").Select
    ActiveCell.FormulaR1C1 = "=SUMIFS(C25,C21,RC29)"
    Range("AF2").Select
    Selection.AutoFill Destination:=Range("AF2:AF7"), Type:=xlFillDefault
    Range("AF2:AF7").Select
    Range("AF2:AF7").Copy
    Range("AF2:AF7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
ElseIf ThisWorkbook.Sheets(SheetName).Range("X1").Value2 = 2022 Then
    '�������
    ThisWorkbook.Sheets(SheetName).Range("AF2:AG7").Select
    Selection.ClearContents
    '2021
    ThisWorkbook.Sheets(SheetName).Range("AF2").Select
    ActiveCell.FormulaR1C1 = "=SUMIFS(C28,C21,RC29)"
    Range("AF2").Select
    Selection.AutoFill Destination:=Range("AF2:AF7"), Type:=xlFillDefault
    Range("AF2:AF7").Select
    Range("AF2:AF7").Copy
    Range("AF2:AF7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    '2022
    ThisWorkbook.Sheets(SheetName).Range("AG2").Select
    ActiveCell.FormulaR1C1 = "=SUMIFS(C25,C21,RC29)"
    Range("AG2").Select
    Selection.AutoFill Destination:=Range("AG2:AG7"), Type:=xlFillDefault
    Range("AG2:AG7").Select
    Range("AG2:AG7").Copy
    Range("AG2:AG7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End If
    
'����������
importWB.Close
ThisWorkbook.Sheets(SheetName).Activate
Company = ThisWorkbook.Sheets(SheetName).Cells(10, 22).Value2
Period = ThisWorkbook.Sheets(SheetName).Cells(10, 23).Value2
'MsgBox ("������ �� �������� " & Company & Period & " ������� ���������")

ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
' ThisWorkbook.Sheets("Preferences").Activate
' Finish = (Now() - Start) * 24 * 60 * 60
'MsgBox (Finish)
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub

'Sub auto_open()                     '������ ��� ��������������� ������� ������� �������
'    Application.Run ("OpenUserForm")
'End Sub

Sub OpenUserForm()              '������ ������
    UserForm1.Show                '�������� �����
End Sub

Sub Data_insertion_SS4()
beging:
 Start = Now()
 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim SheetName As String
 Dim ws As Worksheet
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 SheetName = "���22"
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

ThisWorkbook.Sheets(SheetName).Activate

 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="�������� ���� � ������������ � ���������� ������")

 If TypeName(FilesToOpen) = "Boolean" Then
 GoTo ExitHandler
 End If

ThisWorkbook.Sheets(SheetName).Activate
On Error Resume Next
ActiveSheet.ShowAllData

'������� ������
Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

On Error Resume Next

 importWB.Sheets(1).Activate

'�������� ���������� ������
 ThisWorkbook.Sheets(SheetName).Activate
awLastRow = Cells(Rows.Count, "A").End(xlUp).row
'awLastCol = Cells(begin, Columns.Count).End(xlUp).Column
awLastCol = 29
Range(Cells(1, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With

 '������� ������
For Each ws In importWB.Sheets
 ws.Activate
 iwLastRow = Cells(Rows.Count, "A").End(xlUp).row

 importWB.Activate
 Range(Cells(1, 1), Cells(iwLastRow, awLastCol)).Copy
 
 ThisWorkbook.Sheets(SheetName).Activate
 awFirstRow = 1
 awFirstRow = Cells(Rows.Count, "A").End(xlUp).row
 awFirstCol = 1
 
 Range(Cells(awFirstRow + 1, awFirstCol), Cells(iwLastRow + awFirstRow, awLastCol)).Select
    With Selection
           .PasteSpecial Paste:=xlPasteAll
           .UnMerge
           .Font.Name = "Times New Roman"
           .WrapText = False
           .MergeCells = False
           .Font.Size = 8
    End With
Next ws
'����������
importWB.Close
ThisWorkbook.Sheets(SheetName).Activate
X = ThisWorkbook.Sheets(SheetName).Range("AG5").Value2
Y = ThisWorkbook.Sheets("Calculation22").Range("E2").Value2
If X <> Y Then
    MsgBox "��������!" _
    & vbCr & "����������� ������ �� ����������� ����������� �� ��������� " _
    & "� ����� ��������� �������� � ��������� ���������!" _
    & vbCr & "���������� ���������� �� ���������!" _
    , vbCritical
    result = MsgBox("��������� ���������� ��������� ���������?", vbYesNo)
    If result = vbYes Then
        Application.Run "Data_insertion"
    Else: MsgBox "�������� ��������!" _
        & vbCr & "�������� ���������� ����� �� ����������� � ��������� " _
        & vbCr & ThisWorkbook.Sheets("Calculation22").Range("E2").Value2
    End If
    GoTo beging
End If
MsgBox "������ �� ����������� ���������"

ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
' ThisWorkbook.Sheets("Preferences").Activate
' Finish = (Now() - Start) * 24 * 60 * 60
'MsgBox (Finish)
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub
