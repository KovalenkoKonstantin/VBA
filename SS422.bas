Attribute VB_Name = "SS422"
Sub Insertion_���22()

 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim SheetName As String
 Dim ws As Worksheet
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 SheetName = "���22"
 Limit = 4 '��������� ������� ����
 begin = 15 '������ ��� �������
 
 Dim aw(1 To 4) As Variant
 Dim iw(1 To 4) As Variant
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False

ThisWorkbook.UnProtect Password:="123"
ThisWorkbook.Sheets(SheetName).Visible = True
ThisWorkbook.Sheets(SheetName).Activate

 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="�������� ���� � ������������ � ���������� ������ �� ��� �������������� �����������")

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

ThisWorkbook.Activate

'����������� ������� ������� �����
On Error Resume Next
For i = 1 To 15
    If Worksheets(SheetName).Cells(i, 1) = "���������" Then
        DataRow = i
    End If
Next i

For i = 1 To Limit
    If Worksheets(SheetName).Cells(DataRow, i) = "���������" Then
        aw(1) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ ���������" Then
        aw(2) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�������. �������." Then
        aw(3) = i
    End If
    If Worksheets(SheetName).Cells(DataRow + 1, i) = "�������. �������" Then
        aw(4) = i
    End If
Next i
 
 importWB.Sheets(1).Activate

'����������� ������� ������������� �����
For i = 1 To 30
    If importWB.Sheets(1).Cells(i, 1) = "���������" Then
        ImportFirstDataRow = i
    End If
Next i

For i = 1 To 30
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "���������" Then '-
        iw(1) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ ���������" Then
        iw(2) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������. �������." Then
        iw(3) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow + 1, i) = "�������. �������" Then
        iw(4) = i
    End If
Next i

'�������� ���������� ������
 ThisWorkbook.Sheets(SheetName).Activate
awLastRow = Cells(Rows.Count, "A").End(xlUp).row
'awLastCol = Cells(begin, Columns.Count).End(xlUp).column
Range(Cells(begin, 1), Cells(awLastRow + 3, Limit)).Select
 With Selection
        .Clear
 End With

 '������ ���
Application.StatusBar = "������� ������"

 '������� ������
 importWB.Sheets(1).Activate
 iwLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(1, 1), Cells(iwLastRow, 30)).Select
 With Selection
    .UnMerge
 End With

For i = 1 To 4
''������ ���
'Application.StatusBar = "������������� ����. ���������: " & Int(100 * i / Limit) & "%." & _
'" ����� ��������: " & Int(87 * i / Limit) & "%" & _
'" ��������� ����� �� ����� ���������� ���������: " & _
'Int((100 - Int(87 * i / Limit)) * (((Now() - Start) * 24 * 60 * 60) / (Int(87 * i / Limit)))) & " ������"
 '��� �������
 importWB.Activate
 Range(Cells(begin - 1, iw(i)), Cells(iwLastRow, iw(i))).Copy

 ThisWorkbook.Sheets(SheetName).Activate
 Range(Cells(begin, aw(i)), Cells(iwLastRow, aw(i))).Select
    With Selection
           .PasteSpecial Paste:=xlPasteAll
           .UnMerge
           .Font.Name = "Times New Roman"
           .WrapText = False
           .MergeCells = False
           .Font.Size = 8
    End With
Next i

ThisWorkbook.Sheets(SheetName).Range("A2") = importWB.Sheets(1).Range("C4").Value2

'����������
importWB.Close

ThisWorkbook.Sheets(SheetName).Protect Password:="123"
ThisWorkbook.Sheets(SheetName).Visible = False

MsgBoxEx "������ c ������������ �� ��������" _
& vbCr & ThisWorkbook.Sheets(SheetName).Range("A2").Value2 _
& vbCr & "��������� �������", 0, "���������", 20

ExitHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    ThisWorkbook.Sheets("Preferences").UnProtect Password:="123"
    Rows("81:91").EntireRow.AutoFit
    ThisWorkbook.Sheets("Preferences").Protect Password:="123"
    ThisWorkbook.Protect Password:="123"

 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler
End Sub




