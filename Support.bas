Attribute VB_Name = "Support"
Sub word()
MsgBox "���������"
End Sub
Sub CleanIt()

Dim row, column, X As Integer
On Error GoTo ErrHandler
 Application.ScreenUpdating = False
 Application.EnableEvents = False
 ActiveSheet.DisplayPageBreaks = False
 Application.DisplayStatusBar = False
 Application.DisplayAlerts = False

' ���� ����� ������ �� "���������� ������"
For I = 1 To 50
    If Worksheets("�������������").Cells(I, 1) = "���������� ������" Then
        row = I
    End If
Next

' ����� �������
column = 1


'If Application.Worksheets("�������������").Cells(row, column + 1).Value Is Empty Then
'    GoTo ErrHandler
'End If

X = Application.Worksheets("�������������").Cells(row, column + 1).Value '�������� ������

'������� ������ ������
For I = row + X - 1 To row + 1 Step -1
    Rows(I).EntireRow.delete
'    Range(i, column).EntireRow.Delete
Next

'������� �������� �����
For I = 2 To 3
    Application.Worksheets("�������������").Cells(row, I).Clear
Next

ErrHandler:
 Application.ScreenUpdating = True
 Application.EnableEvents = True
 ActiveSheet.DisplayPageBreaks = True
 Application.DisplayStatusBar = True
 Application.DisplayAlerts = True
End Sub

Sub CleanIt2()
'
'Dim row_b, row_e As Integer
'On Error GoTo ErrHandler
' Application.ScreenUpdating = False
' Application.EnableEvents = False
' ActiveSheet.DisplayPageBreaks = False
' Application.DisplayStatusBar = False
' Application.DisplayAlerts = False
'
'' ���� ����� ���� �� "����� ����������"
'For i = 1 To 50
'    If Worksheets("9").Cells(i, 2) = "����� ����������" Then
'        row_b = i
'    End If
'Next
'
'' ���� ����� ���� �� "���������� �������������� ���������"
'For i = 1 To 50
'    If Worksheets("9").Cells(i, 2) = "���������� �������������� ���������" Then
'        row_e = i
'    End If
'Next
'
''���������� ������������� ��������
'If row_e - row_b = 1 Then
'    GoTo ErrHandler
'Else
'
''������� ������ ������
'    Range("A" & row_b + 1 & ":A" & row_e - 1).EntireRow.Delete
'End If
'
'ErrHandler:
' Application.ScreenUpdating = True
' Application.EnableEvents = True
' ActiveSheet.DisplayPageBreaks = True
' Application.DisplayStatusBar = True
' Application.DisplayAlerts = True
End Sub

Sub Social_contribution()
 
' Dim FilesToOpen
 Dim ThisWorkbook As Workbook
' Dim ws, this As Worksheet
' Dim pt As PivotTable
' Dim �, d As Range
 Dim temp, temp1, temp2 As String
' Dim x As Integer
 X = "7,8"
 Y = "��2"
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
 
 '����� ������ �������
 ThisWorkbook.Sheets(Y).Activate
 s = Cells(Rows.Count, "B").End(xlUp).row '��������� ������� ������
 K = Cells(2, Columns.Count).End(xlToLeft).column '��������� ������� ������
 
 ' ���� ������ �������
For I = 1 To K
    '���������� ������� ���������� �����
    If Worksheets(Y).Cells(2, I) = "������" Then
        zp = I
    End If
    '���������� ������� ���������� ������
    If Worksheets(Y).Cells(2, I) = "% ��������� �������" Then
        sp = I
    End If
    '���������� ������� ����
    If Worksheets(Y).Cells(2, I) = "����" Then
        yr = I
    End If
    '���������� ������� ��������
    If Worksheets(Y).Cells(2, I) = "��������" Then
        check = I
    End If
Next I

'������� ���������� ��������
    ThisWorkbook.Sheets(X).Activate
    Range(check & "4:" & check & K).Clear
'main
For I = 3 To s
'skip constant rows
    If Worksheets(Y).Cells(I, 2).Value2 = "�����" Then
        GoTo ExitHandler
    End If
    If Worksheets(Y).Cells(I, 2).Value2 = Worksheets(Y).Cells(I, 3).Value2 Then
        I = I + 1
    End If
'��������� �������� �������� � ���� ��������
    Worksheets(Y).Cells(I, zp).Copy
    ThisWorkbook.Sheets(X).Activate
    Range("B4").Select
    With Selection
        .PasteSpecial Paste:=xlPasteValues
    End With
    
'������ �������� ����� � ����� � ������������ ���������� ������
    temp = Worksheets(Y).Cells(I, yr).Value2
    ThisWorkbook.Sheets(X).Activate
    Range("J4:J15").Clear
        For c = 4 To 15
            temp1 = Cells(c, 1).Value2
            Cells(c, 10) = temp1 & " " & temp
        Next c
 '���� ���������� �������
    For j = 4 To 15
'        a = ThisWorkbook.Sheets(x).Cells(j, 10).Value2
'        b = Worksheets(y).Cells(i, 3).Value2
        If ThisWorkbook.Sheets(X).Cells(j, 10).Value2 = Worksheets(Y).Cells(I, 3).Value2 Then
            destinct = j '������ ��� ��� �������� �������� � ��������� ���������
        End If
    Next j
'������� �������� ����� �� 7,8
    ThisWorkbook.Sheets(X).Activate
    Range("J4:J15").Clear
 '��������� % ��� ������ � ��������� ���������
 ThisWorkbook.Sheets(X).Activate
 Cells(destinct, 9).Copy
 ThisWorkbook.Sheets(Y).Activate
 Cells(I, check).Select
    With Selection
        .PasteSpecial Paste:=xlPasteValues
    End With
Next I


ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub


Sub RefreshOrder()
Dim objWord As Object
Dim FileStart
Dim FileNew

Set objWord = CreateObject("Word.Application")

    FileSt = "D:\���\���\022-7\������.docx"
    FileNew = "D:\���\���\022-7\������1.docx"

    objWord.Documents.Open FileSt
                
    For Each MyLink In objWord.ActiveDocument.Fields
        MyLink.Update
        MyLink.Unlink
    Next MyLink

    objWord.ActiveDocument.SaveAs _
            Filename:=FileNew, _
            FileFormat:=wdFormatDocument, _
            Password:="", _
            AddToRecentFiles:=True, _
            WritePassword:="", _
            ReadOnlyRecommended:=False
objWord.Quit
End Sub

Sub Budget()

 Dim ThisWorkbook, importWB As Workbook
 Dim FilesToOpen
' Dim MyRange, MyCell As range
 Dim key As String
 Ye_ar = 2022 '��������� ��� �������
 X = 4 '���������� ������ ��� �������
 DataTab = "������" '���� ������
 WorkTab = "��" '������� ����
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 
 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="���� ��� �������")
 
 If TypeName(FilesToOpen) = "Boolean" Then
 GoTo ExitHandler
 End If
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 
 '����� ������ ������� ������
 ThisWorkbook.Sheets(DataTab).Activate
 FirstRowData = Columns(1).Find("*", LookIn:=xlValues).row '��� ������� ��������
 LastRowData = Cells(Rows.Count, 2).End(xlUp).row '��������� ��� ������
 LastColumnData = Cells(FirstRowData, Columns.Count).End(xlToLeft).column '��������� ������� ������
  
 '���� ������ ������� � ����� ������
For I = 1 To LastColumnData
    '���������� ������� �����1
    If Worksheets(DataTab).Cells(FirstRowData, I) = "����1" Then
        Key1ColData = I
    End If
    '���������� ������� ����� ����������
    If Worksheets(DataTab).Cells(FirstRowData, I) = "���������" Then
        EmployeeColData = I
    End If
    '���������� ������� ���������
    If Worksheets(DataTab).Cells(FirstRowData, I) = "���������" Then
        PositionColData = I
    End If
    '���������� ������� ����
    If Worksheets(DataTab).Cells(FirstRowData, I) = "���" Then
        YearColData = I
    End If
    '���������� ������� ������
    If Worksheets(DataTab).Cells(FirstRowData, I) = "������" Then
        PrizeColData = I
    End If
    '���������� ������� �����2
    If Worksheets(DataTab).Cells(FirstRowData, I) = "����2" Then
        Key2ColData = I
    End If
    '���������� ������� ������ ���������
    If Worksheets(DataTab).Cells(FirstRowData, I) = "���������2" Then
        Position2ColData = I
    End If
    '���������� ������� ������ � �������� ���������
    If Worksheets(DataTab).Cells(FirstRowData, I) = "���" Then
        MonthNumberColData = I
    End If
    '���������� ������� �������
    If Worksheets(DataTab).Cells(FirstRowData, I) = "�������" Then
        DecemberColData = I
    End If
Next I
 
 '�������� ���������� ������
 Range(Cells(FirstRowData + 1, Key1ColData), Cells(LastRowData, LastColumnData)).Select
 With Selection
        .ClearContents
 End With

'������� ������
Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))
For I = 1 To X
    On Error Resume Next
     importWB.Sheets(I).Activate
     lLastRow = Cells(Rows.Count, 1).End(xlUp).row
     j = lLastRow
     
     importWB.Sheets(I).Activate
     Range("A3:N" & j).Select
     Range("A3:N" & j).Copy
     ThisWorkbook.Sheets(DataTab).Activate
     '���� ����� �������� ���������� ����
     lLastRow = Cells(Rows.Count, 3).End(xlUp).row
     jRenow = lLastRow
     Range(Cells(jRenow + 1, EmployeeColData), Cells(jRenow + j, DecemberColData)).Select
     With Selection
            .PasteSpecial Paste:=xlPasteAll
            .UnMerge
            .Font.Name = "Times New Roman"
            .WrapText = False
            .MergeCells = False
            .Font.Size = 8
     End With
     
     '���������� �������������� ���� � ������� ������
     lLastRow = Cells(Rows.Count, 2).End(xlUp).row
     jNew = lLastRow
     Range(Cells(jRenow + 1, PrizeColData), Cells(jNew, PrizeColData)).Value2 = I

Next I
'�������� ����� ������
importWB.Close

'������� ������� ������� ����
 Range(Cells(FirstRowData + 1, YearColData), Cells(jNew, YearColData)).FormulaR1C1 = _
    "=IF(RC[1]=1,2022,IF(RC[1]=2,2023,IF(RC[1]=3,2022,IF(RC[1]=4,2023))))"
    '������� �������� ������ �������
    Range(Cells(FirstRowData + 1, YearColData), Cells(jNew, YearColData)).Select
    Range(Cells(FirstRowData + 1, YearColData), Cells(jNew, YearColData)).Copy
    Range(Cells(FirstRowData + 1, YearColData), Cells(jNew, YearColData)).Select
 With Selection
        .PasteSpecial Paste:=xlPasteValues
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
        .Font.Size = 8
 End With

'���������� �������������� ������ � ������� �����2
    Range(Cells(FirstRowData + 1, Key2ColData), Cells(jNew, Key2ColData)).FormulaR1C1 = "=IF(OR(RC[-1]=3,RC[-1]=4),""������"","""")"
    '������� �������� ������ � �������
    Range(Cells(FirstRowData + 1, Key2ColData), Cells(jNew, Key2ColData)).Select
    Range(Cells(FirstRowData + 1, Key2ColData), Cells(jNew, Key2ColData)).Copy
    Range(Cells(FirstRowData + 1, PrizeColData), Cells(jNew, PrizeColData)).Select
 With Selection
        .PasteSpecial Paste:=xlPasteValues
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
        .Font.Size = 8
 End With
 
 '���������� ������� ����� � ������ �������
Range(Cells(FirstRowData + 1, 1), Cells(jNew, 1)).FormulaR1C1 = "=CONCATENATE(RC[2],RC[16],RC[17])"

'���������� ������� �����1
Range(Cells(FirstRowData + 1, Key1ColData), Cells(jNew, Key1ColData)).FormulaR1C1 = "=CONCATENATE(RC[1],RC[15],RC[2],RC[16])"

'���������� ������� �����2
Range(Cells(FirstRowData + 1, Key2ColData), Cells(jNew, Key2ColData)).FormulaR1C1 = "=RC[-18]"

'���������� ������� ���������2
Range(Cells(FirstRowData + 1, Key2ColData), Cells(jNew, Key2ColData)).FormulaR1C1 = _
        "=IF(AND(RC[-17]=R[1]C[-17],RC[-16]<>R[1]C[-16]),R[1]C[-16],"""")"
        
'���������� ������� ������ � �������� ��������� ��� �������� � ������� �����������
Range(Cells(FirstRowData + 1, Key2ColData), Cells(jNew, Key2ColData)).FormulaArray = _
        "=IF(RC[-1]="""","""",MATCH(TRUE(),(RC[-16]:RC[-5]=""""),FALSE()))"
 
' '�������� �������� �� ������� �����2
' range(Cells(FirstRowData + 1, LastColumnData + 1), Cells(jNew, LastColumnData + 1)).ClearContents

ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
'    ThisWorkbook.Sheets("Preferences").Activate
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub

Sub RefreshPivots()
Dim pt As PivotTable
Dim ws As Worksheet
Dim ThisWorkbook As Workbook

Set ThisWorkbook = ActiveWorkbook
On Error GoTo ExitHandler

Application.ScreenUpdating = False
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

For Each ws In ThisWorkbook
pt.Refresh
Next ws

ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
' ThisWorkbook.Sheets("Preferences").Activate
 Exit Sub
End Sub

Sub ShowTabs()
 Dim tb
 For Each tb In Worksheets
 tb.Visible = True
 Next
 ActiveWorkbook.Sheets("Preferences").Activate
End Sub

Sub HideSys()
 Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Not ws.Range("ZZ500").Value2 = Empty Then
            ws.Visible = False
        End If
    Next ws
 ActiveWorkbook.Sheets("Preferences").Activate
End Sub

Sub UnhideSys()
 Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Not ws.Range("ZZ500").Value2 = Empty Then
            ws.Visible = True
        End If
    Next ws
 ActiveWorkbook.Sheets("Preferences").Activate
End Sub

Sub HideEmpty()
 Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Not ws.Range("ZZ501").Value2 = Empty Then
            ws.Visible = True
            ws.Select
            With ws.Tab
                .ColorIndex = xlNone
                .TintAndShade = 0
            End With
            ws.Visible = False
        End If
    Next ws
 ActiveWorkbook.Sheets("Preferences").Activate
End Sub

Sub Protect()
 Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect Password:="123"
    Next ws
    ThisWorkbook.Protect Password:="123"
 ActiveWorkbook.Sheets("Preferences").Activate
End Sub

Sub UnProtect()
 Dim ws As Worksheet
 On Error GoTo errorhandler
    For Each ws In ThisWorkbook.Worksheets
        ws.UnProtect Password:="123"
    Next ws
    ThisWorkbook.UnProtect Password:="123"

errorhandler:
 MsgBox ("��� ����� ��������������")
ActiveWorkbook.Sheets("Preferences").Activate

End Sub


Sub button1()
    [J61] = 0
    Range("J60").GoalSeek Goal:=0, ChangingCell:=Range("J61")
End Sub
Sub button2()
    [J61] = 0
    Range("J60").GoalSeek Goal:=0, ChangingCell:=Range("J61")
End Sub

