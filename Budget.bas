Attribute VB_Name = "Budget"
Sub BudgetInsertion()
 
'Dim I As Worksheet
' '������� ������ � ������ ������
' For Each I In importWB.Sheets
'     I.Activate
'     iwLastRow = Cells(Rows.Count, SearchRow).End(xlUp).row
'     importWB.Activate
'     Range(Cells(1, 1), Cells(iwLastRow, awLastCol)).Copy
'
'     ThisWorkbook.Sheets(SheetName).Activate
'     awLastRow = Cells(Rows.Count, SearchRow).End(xlUp).row
'     Range(Cells(awLastRow, 1), Cells(iwLastRow + awLastRow - 1, awLastCol)).Select
'        With Selection
'               .PasteSpecial Paste:=xlPasteAll
'               .UnMerge
'               .Font.Name = "Times New Roman"
'               .WrapText = False
'               .MergeCells = False
'               .Font.Size = 10
'        End With
' Next I

'����������


Start = Now()
 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim ws, this As Worksheet
 Dim MyRange, MyCell As Range
 Dim SheetName, CompanyName As String
  
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 SheetName = "������"
 DistinctYear = 2021
 SearchRow = "A"
 Limit = 54 '��������� ������� �����
 begin = 5 '������ ��� �������
 CompanyName = ThisWorkbook.Sheets("Preferences").Range("C7").Value2 '��� �������
 
 Dim aw(1 To 54) As Variant
 Dim iw(1 To 54) As Variant
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False

FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="�������� ��������� ��������� �� �������� " & CompanyName & " �� " & DistinctYear & " ���")
 
 '������ ���
Application.StatusBar = "������ ������..."

 If TypeName(FilesToOpen) = "Boolean" Then ',���� ������ ������ ������ ����� �� ���������
 GoTo ExitHandler
 End If

'������� ������
Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

On Error Resume Next

 '�������� ������������ ������ ������
 importWB.Sheets(1).Activate
 Range("A2").Select
 ActiveCell.FormulaR1C1 = "=COUNTIF(RC[4]:RC[51],""<>"""""")"
 If Range("A2").Value2 <> 48 Or Range("C3").Value2 <> CompanyName Then
    Range("A2").Select
    With Selection:
        .Clear
    End With
    MsgBoxEx "������ �������������� ����." _
    & vbCr & "������� �������.", vbCritical, "Bad Day", 20
    GoTo ExitHandler
 Else
 Range("A2").Select
    With Selection:
        .Clear
    End With
    MsgBoxEx "������ ���������� ���� � ��������" _
    & vbCr & "����������.", 0, "Succes", 5
 End If

ThisWorkbook.Sheets(SheetName).Activate
On Error Resume Next
ActiveSheet.ShowAllData

ThisWorkbook.Activate

'����������� ������� ������� �����
On Error Resume Next
For I = 1 To 10
    If Worksheets(SheetName).Cells(I, 1) = "���������" Then
        DataRow = I
    End If
Next I

For I = 1 To Limit
    If Worksheets(SheetName).Cells(DataRow, I) = "���������" Then
        aw(1) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "����������" Then
        aw(2) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�����������" Then
        aw(3) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���������" Then
        aw(4) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������" Then
        aw(5) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ ������" Then
        aw(6) = I
    End If
' ------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, I) = "������ 2021" Then
        aw(7) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� 2021" Then
        aw(8) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���� 2021" Then
        aw(9) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ 2021" Then
        aw(10) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "��� 2021" Then
        aw(11) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���� 2021" Then
        aw(12) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���� 2021" Then
        aw(13) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ 2021" Then
        aw(14) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������� 2021" Then
        aw(15) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� 2021" Then
        aw(16) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ 2021" Then
        aw(17) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� 2021" Then
        aw(18) = I
    End If
' ------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, I) = "������ 2022" Then
        aw(19) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� 2022" Then
        aw(20) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���� 2022" Then
        aw(21) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ 2022" Then
        aw(22) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "��� 2022" Then
        aw(23) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���� 2022" Then
        aw(24) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���� 2022" Then
        aw(25) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ 2022" Then
        aw(26) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������� 2022" Then
        aw(27) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� 2022" Then
        aw(28) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ 2022" Then
        aw(29) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� 2022" Then
        aw(30) = I
    End If
' ------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, I) = "������ 2023" Then
        aw(31) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� 2023" Then
        aw(32) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���� 2023" Then
        aw(33) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ 2023" Then
        aw(34) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "��� 2023" Then
        aw(35) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���� 2023" Then
        aw(36) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���� 2023" Then
        aw(37) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ 2023" Then
        aw(38) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������� 2023" Then
        aw(39) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� 2023" Then
        aw(40) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ 2023" Then
        aw(41) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� 2023" Then
        aw(42) = I
    End If
' ------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, I) = "������ 2024" Then
        aw(43) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� 2024" Then
        aw(44) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���� 2024" Then
        aw(45) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ 2024" Then
        aw(46) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "��� 2024" Then
        aw(47) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���� 2024" Then
        aw(48) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���� 2024" Then
        aw(49) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ 2024" Then
        aw(50) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������� 2024" Then
        aw(51) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� 2024" Then
        aw(52) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ 2024" Then
        aw(53) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� 2024" Then
        aw(54) = I
    End If
    
Next I
 
 importWB.Sheets(1).Activate

'����������� ������� ������������� �����
For I = 1 To 20
    If importWB.Sheets(1).Cells(I, 1) = "�����������" Then
        ImportFirstDataRow = I
    End If
Next I
'For I = 1 To 20
'    If importWB.Sheets(1).Cells(I, 1) = "���������" Then
'        ImportSecondDataRow = I
'    End If
'Next I

For I = 1 To Limit
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "���������" Then
        aw(1) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "����������" Then
        aw(2) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "�����������" Then
        aw(3) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "���������" Then
        aw(4) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������" Then
        aw(5) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������ ������" Then
        aw(6) = I
    End If
' ------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������ 2021" Then
        aw(7) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������� 2021" Then
        aw(8) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "���� 2021" Then
        aw(9) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������ 2021" Then
        aw(10) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "��� 2021" Then
        aw(11) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "���� 2021" Then
        aw(12) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "���� 2021" Then
        aw(13) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������ 2021" Then
        aw(14) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "�������� 2021" Then
        aw(15) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������� 2021" Then
        aw(16) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������ 2021" Then
        aw(17) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������� 2021" Then
        aw(18) = I
    End If
' ------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������ 2022" Then
        aw(19) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������� 2022" Then
        aw(20) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "���� 2022" Then
        aw(21) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������ 2022" Then
        aw(22) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "��� 2022" Then
        aw(23) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "���� 2022" Then
        aw(24) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "���� 2022" Then
        aw(25) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������ 2022" Then
        aw(26) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "�������� 2022" Then
        aw(27) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������� 2022" Then
        aw(28) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������ 2022" Then
        aw(29) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������� 2022" Then
        aw(30) = I
    End If
' ------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������ 2023" Then
        aw(31) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������� 2023" Then
        aw(32) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "���� 2023" Then
        aw(33) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������ 2023" Then
        aw(34) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "��� 2023" Then
        aw(35) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "���� 2023" Then
        aw(36) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "���� 2023" Then
        aw(37) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������ 2023" Then
        aw(38) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "�������� 2023" Then
        aw(39) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������� 2023" Then
        aw(40) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������ 2023" Then
        aw(41) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������� 2023" Then
        aw(42) = I
    End If
' ------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������ 2024" Then
        aw(43) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������� 2024" Then
        aw(44) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "���� 2024" Then
        aw(45) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������ 2024" Then
        aw(46) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "��� 2024" Then
        aw(47) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "���� 2024" Then
        aw(48) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "���� 2024" Then
        aw(49) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������ 2024" Then
        aw(50) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "�������� 2024" Then
        aw(51) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������� 2024" Then
        aw(52) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������ 2024" Then
        aw(53) = I
    End If
    If Worksheets(SheetName).Cells(ImportFirstDataRow, I) = "������� 2024" Then
        aw(54) = I
    End If

Next I

'�������� ���������� ������
 ThisWorkbook.Sheets(SheetName).Activate
awLastRow = Cells(Rows.Count, SearchRow).End(xlUp).row
awLastCol = Cells(begin, Columns.Count).End(xlUp).column
Range(Cells(begin, 1), Cells(awLastRow, Limit)).Select
 With Selection
        .Clear
 End With

 '����������� ���������� ���� IWB
 importWB.Sheets(1).Activate
 iwLastRow = Cells(Rows.Count, SearchRow).End(xlUp).row

For I = 1 To Limit
'������ ���
Application.StatusBar = "���������� ����� ������: " & Int(100 * I / Limit) & "%."
 
 '����������
 importWB.Activate
 Range(Cells(begin - 2, iw(I)), Cells(iwLastRow, iw(I))).Copy

 ThisWorkbook.Sheets(SheetName).Activate
 Range(Cells(begin, aw(I)), Cells(iwLastRow + 2, aw(I))).Select
    With Selection
           .PasteSpecial Paste:=xlPasteAll
           .UnMerge
           .Font.Name = "Times New Roman"
           .WrapText = False
           .MergeCells = False
           .Font.Size = 8
    End With
Next I

'������ ���
Application.StatusBar = "���������: 95 %"

'�������
ThisWorkbook.Sheets(SheetName).Activate
Columns("E:AZ").Select
    Selection.NumberFormat = _
        "_-* #,##0.00 _?_-;-* #,##0.00 _?_-;_-* ""-""?? _?_-;_-@_-"

ExitHandler:
    On Error Resume Next
    importWB.Close
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
 ThisWorkbook.Sheets("Preferences").Activate
 Exit Sub
  
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler



End Sub
