Attribute VB_Name = "Expenditures"
Sub Overheads()
 
 Start = Now()
 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim ws, this As Worksheet
 Dim MyRange, MyCell As Range
 Dim SheetName, CompanyName As String
  
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 SheetName = "Expenditures"
' DistinctYear = 2021
 Limit = 138 '��������� ������� ����
 begin = 12 '������ ��� �������
 CompanyName = ThisWorkbook.Sheets("Preferences").Range("C7").Value2 '��� �������
 
 Dim aw(1 To 138) As Variant
 Dim iw(1 To 138) As Variant
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual

FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="�������� ��������� ��������� �� �������� " & CompanyName)
 
 MsgBoxEx "��������� ��������� ������ ���� � ������ ����!" _
 & vbCr & "� ��������� ������, �� ��� �������� ���� ����������� ���������." _
    & vbCr & "...", vbCritical, "Pay attention", 20
 
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
 Range("G2").Select
 ActiveCell.FormulaR1C1 = "=YEAR(MID(RC[-4],SEARCH("" "",RC[-4],1)+1,10))"
' If Range("G2").Value2 <> DistinctYear Or Range("A11").Value2 <> CompanyName Then
 If Range("A11").Value2 <> CompanyName Then
'    Range("G2").Select
'    With Selection:
'        .Clear
'    End With
    MsgBoxEx "������� ������������ ��������� ���������." _
    & vbCr & "������� �������.", vbCritical, "Bad Day", 20
    GoTo ExitHandler
' ElseIf Range("G2").Value2 = DistinctYear Then
' Range("G2").Select
'    With Selection:
'        .Clear
'    End With
 End If
 
  Range("G2").Select
    With Selection:
        .Clear
    End With

ThisWorkbook.Sheets(SheetName).Activate
On Error Resume Next
ActiveSheet.ShowAllData

ThisWorkbook.Activate

'����������� ������� ������� �����
On Error Resume Next
For i = 1 To 20
    If Worksheets(SheetName).Cells(i, 1) = "���������" Then
        DataRow = i
    End If
Next i

For i = 1 To Limit
    If Worksheets(SheetName).Cells(DataRow, i) = "���������" Then
        aw(1) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�����" Then
        aw(2) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "��������� ����� �����" Then
        aw(3) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ �����" Then
        aw(4) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "��������� ����� ����" Then
        aw(5) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ ����" Then
        aw(6) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "���������� ���� ����� 20,26,44 �����" Then
        aw(7) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "��� ��������" Then
        aw(8) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ ��������� �������" Then
        aw(9) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ �� ����������.������ �� ��" Then
        aw(10) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������������� �������" Then
        aw(11) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "���������" Then
        aw(12) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "��� ���������" Then
        aw(13) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "���� ��������" Then
        aw(14) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ ��������� �������� � ��� �����" Then
        aw(15) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ ������" Then
        aw(16) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "����� ����" Then
        aw(17) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "����� �����" Then
        aw(18) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "���������� ����" Then
        aw(19) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "���������� �����" Then
        aw(20) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, i) = "���������" Then
        aw(21) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ ���������� ������ �� ���� ������������" Then
        aw(22) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ ���������� ������" Then
        aw(23) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "����� �� ����" Then
        aw(24) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "���������� �� ������� (���������� ��� �� ������)" Then
        aw(25) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�������� �� ��������� � �������������" Then
        aw(26) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ ������� �� ����������� ����" Then
        aw(27) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ �� ����� ������" Then
        aw(28) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�������� �� ��������������� � �������� ����������� �����" Then
        aw(29) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ �� ������ ����" Then
        aw(30) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ �� ���� ����" Then
        aw(31) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "����������� ������� (������ ��������)" Then
        aw(32) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "����������� ������" Then
        aw(33) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ �������" Then
        aw(34) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�������� ������� ��� ����������" Then
        aw(35) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������� �� ����� �� �������� �� �������� ���" Then
        aw(36) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "����� �� ���� (��������������� ������������ ����)" Then
        aw(37) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�������� �� ��������������� � �������� ����������� ����� (��������������� ������������ ����)" Then
        aw(38) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�������� �� ������ �� ����������, ������������� ��������������� �����" Then
        aw(39) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "����� �� ���� 26 �����������" Then
        aw(40) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ �� �����������" Then
        aw(41) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�������� ��������" Then
        aw(42) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�������� �����������" Then
        aw(43) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ � ������ ��������� ������������(�����������)" Then
        aw(44) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "��������� �������������� ������������ ������" Then
        aw(45) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ �� ������ ���� � ������ ��" Then
        aw(46) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�������������� �� �����������" Then
        aw(47) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������� (������, ������)" Then
        aw(48) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "����������� �� ������" Then
        aw(49) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ � ������ ��������� ������������ (��������)" Then
        aw(50) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������" Then
        aw(51) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������� �� ���������� ����������, ���������� ������������" Then
        aw(52) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�������� ������" Then
        aw(53) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������� �����" Then
        aw(54) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�������������� ������� ������ (������������)" Then
        aw(55) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ ������ � ����������� � �������� ���" Then
        aw(56) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�������������� ������ ������ ����������" Then
        aw(57) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ �������� �.�" Then
        aw(58) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "���������� �� ������������ �������" Then
        aw(59) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ ��� ������ �������� �� ��" Then
        aw(60) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ ���� ����� �� ������-����������" Then
        aw(61) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������� �� ����� �� ������� �� 3 ��� ��� ������" Then
        aw(62) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "����� �� ����� (������������� ������������� �������)" Then
        aw(63) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�������� �� ��������� � ������������� (�� ����� ��������������� ������������ �������)" Then
        aw(64) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�������� �� ��������������� � �������� ����������� ����� (�� ����� ������� ����� �������)" Then
        aw(65) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "���������� �� �������" Then
        aw(66) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������� �� ������ � ������ �����" Then
        aw(67) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�� ������ ������ �� ���" Then
        aw(68) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ ���������� (���������, �����, ����������, �������)" Then
        aw(69) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�������� �� ���������� ������������������ � ��������� �" Then
        aw(70) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ �� ������������ � �����" Then
        aw(71) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������� �� ������ � ����������� ��� (������� �����)" Then
        aw(72) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������� �� ������ � ����������� ��� (������ �����)" Then
        aw(73) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������� �� ����������� ��� ������������� ����� �������� �������" Then
        aw(74) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�������������� ��� ������� �� ������������ � �����" Then
        aw(75) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������� �� ��������������� ������� ����" Then
        aw(76) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������������ ������" Then
        aw(77) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, i) = "����� �� ����� � ����������� �����" Then
        aw(78) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������� �� ���� ������� �����������" Then
        aw(79) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "����� � ����������� ����� (���. ����)" Then
        aw(80) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "���������" Then
        aw(81) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "���� � ������ � ���� ������� ��������" Then
        aw(82) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������� ������������" Then
        aw(83) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�������" Then
        aw(84) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "��������� ��������" Then
        aw(85) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������� ������� ����������� ����" Then
        aw(86) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������� �����" Then
        aw(87) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, i) = "��������" Then
        aw(88) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "����" Then
        aw(89) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "���� � ����������" Then
        aw(90) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�������� �� �������� �� ������ ������������" Then
        aw(91) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "��������� �� �/� ��������������� ��������" Then
        aw(92) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "��������� �� ���. ����� ���������" Then
        aw(93) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "��������� �� ���������" Then
        aw(94) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������� ������������� �������� ����������" Then
        aw(95) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�������� �� �������� �� ������ ������������-��������������� �����" Then
        aw(96) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "��������� �� �������� �� ��������" Then
        aw(97) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "��������� �� �������� 73,03" Then
        aw(98) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "��������� �� ���. ����� ����. ������" Then
        aw(99) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "��������� �� ��������������� ���������" Then
        aw(100) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, i) = "������" Then
        aw(101) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "��� �� ����������" Then
        aw(102) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "��� � ����������" Then
        aw(103) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�����" Then
        aw(104) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "���" Then
        aw(105) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "��� ��" Then
        aw(106) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, i) = "% ��������� �������" Then
        aw(107) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "���� �������" Then
        aw(108) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, i) = "�������� �� ���� ������ �� ������ ��������������� �����" Then
        aw(109) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ �� �������������� ���� (���) ������ ������" Then
        aw(110) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������������" Then
        aw(111) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ ����������� ����� ������ (��)" Then
        aw(112) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ ��������" Then
        aw(113) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ �� ������ (�� �����)" Then
        aw(114) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "����� �� �����" Then
        aw(115) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�������� �� ��������� � ������������� (�� �����)" Then
        aw(116) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�������� �� ��������� � ������������� (��������������� ������������ ����)" Then
        aw(117) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "����������� ������" Then
        aw(118) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "����������� �������� �� ��������� �������" Then
        aw(119) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ �������� (� ������ ��)" Then
        aw(120) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ �����������" Then
        aw(121) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ ����������� (� ������ ��)" Then
        aw(122) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ �� ������ ���� (� ������ ��)" Then
        aw(123) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�������� �� ��������� � ������������� (�� ����� �������. ������������� �������)" Then
        aw(124) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ ������� (� ������ ��)" Then
        aw(125) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "���� ����������" Then
        aw(126) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ ��������" Then
        aw(127) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "���" Then
        aw(128) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������� �� ������ � ������ ����� (����������� � �������� ���)" Then
        aw(134) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ ����������� (� ������ ��)" Then
        aw(135) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������ �����������" Then
        aw(136) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "�������� �� ��������������� � �������� ����������� ����� (�� ����� �������. �����. �������)" Then
        aw(137) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "������� �����" Then
        aw(138) = i
    End If
    
Next i
 
 importWB.Sheets(1).Activate

'����������� ������� ������������� �����
For i = 1 To 20
    If importWB.Sheets(1).Cells(i, 1) = "�����������" Then
        ImportFirstDataRow = i
    End If
Next i
For i = 1 To 20
    If importWB.Sheets(1).Cells(i, 1) = "���������" Then
        ImportSecondDataRow = i
    End If
Next i

For i = 1 To Limit
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "���������" Then '-
        iw(1) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�����" Then
        iw(2) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "��������� ����� �����" Then
        iw(2) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ �����" Then
        iw(4) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "��������� ����� ����" Then
        iw(5) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ ����" Then
        iw(6) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "���������� ���� ����� 20,26,44 �����" Then
        iw(7) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "��� ��������" Then
        iw(8) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ ��������� �������" Then
        iw(9) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "������ �� ����������.������ �� ��" Then
        iw(10) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "������������� �������" Then '-
        iw(11) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "���������" Then '-
        iw(12) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "��� ���������" Then '-
        iw(13) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "���� ��������" Then '-
        iw(14) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "������ ��������� �������� � ��� �����" Then '-
        iw(15) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "������ ������" Then '-
        iw(16) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "����� ����" Then '-
        iw(17) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "����� �����" Then '-
        iw(18) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "����" Then '-
        iw(19) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "�����" Then '-
        iw(20) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "���������" Then
        iw(21) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ ���������� ������ �� ���� ������������" Then
        iw(22) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ ���������� ������" Then
        iw(23) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "����� �� ����" Then
        iw(24) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "���������� �� ������� (���������� ��� �� ������)" Then
        iw(25) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������� �� ��������� � �������������" Then
        iw(26) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ ������� �� ����������� ����" Then
        iw(27) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ �� ����� ������" Then
        iw(28) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������� �� ��������������� � �������� ����������� �����" Then
        iw(29) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ �� ������ ����" Then
        iw(30) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ �� ���� ����" Then
        iw(31) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "����������� ������� (������ ��������)" Then
        iw(32) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "����������� ������" Then
        iw(33) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ �������" Then
        iw(34) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������� ������� ��� ����������" Then
        iw(35) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������� �� ����� �� �������� �� �������� ���" Then
        iw(36) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "����� �� ���� (��������������� ������������ ����)" Then
        iw(37) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������� �� ��������������� � �������� ����������� ����� (��������������� ������������ ����)" Then
        iw(38) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������� �� ������ �� ����������, ������������� ��������������� �����" Then
        iw(39) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "����� �� ���� 26 �����������" Then
        iw(40) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ �� �����������" Then
        iw(41) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������� ��������" Then
        iw(42) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������� �����������" Then
        iw(43) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ � ������ ��������� ������������(�����������)" Then
        iw(44) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "��������� �������������� ������������ ������" Then
        iw(45) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ �� ������ ���� � ������ ��" Then
        iw(46) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������������� �� �����������" Then
        iw(47) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������� (������, ������)" Then
        iw(48) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "����������� �� ������" Then
        iw(49) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ � ������ ��������� ������������ (��������)" Then
        iw(50) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������" Then
        iw(51) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������� �� ���������� ����������, ���������� ������������" Then
        iw(52) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������� ������" Then
        iw(53) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������� �����" Then
        iw(54) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������������� ������� ������ (������������)" Then
        iw(55) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ ������ � ����������� � �������� ���" Then
        iw(56) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������������� ������ ������ ����������" Then
        iw(57) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ �������� �.�" Then
        iw(58) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "���������� �� ������������ �������" Then
        iw(59) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ ��� ������ �������� �� ��" Then
        iw(60) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ ���� ����� �� ������-����������" Then
        iw(61) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������� �� ����� �� ������� �� 3 ��� ��� ������" Then
        iw(62) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "����� �� ����� (������������� ������������� �������)" Then
        iw(63) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������� �� ��������� � ������������� (�� ����� ��������������� ������������ �������)" Then
        iw(64) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������� �� ��������������� � �������� ����������� ����� (�� ����� ������� ����� �������)" Then
        iw(65) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "���������� �� �������" Then
        iw(66) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������� �� ������ � ������ �����" Then
        iw(67) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�� ������ ������ �� ���" Then
        iw(68) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ ���������� (���������, �����, ����������, �������)" Then
        iw(69) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������� �� ���������� ������������������ � ��������� �" Then
        iw(70) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ �� ������������ � �����" Then
        iw(71) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������� �� ������ � ����������� ��� (������� �����)" Then
        iw(72) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������� �� ������ � ����������� ��� (������ �����)" Then
        iw(73) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������� �� ����������� ��� ������������� ����� �������� �������" Then
        iw(74) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������������� ��� ������� �� ������������ � �����" Then
        iw(75) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������� �� ��������������� ������� ����" Then
        iw(76) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������������ ������" Then
        iw(77) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "����� �� ����� � ����������� �����" Then
        iw(78) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������� �� ���� ������� �����������" Then
        iw(79) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "����� � ����������� ����� (���. ����)" Then
        iw(80) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "���������" Then
        iw(81) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "���� � ������ � ���� ������� ��������" Then
        iw(82) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������� ������������" Then
        iw(83) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������" Then
        iw(84) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "��������� ��������" Then
        iw(85) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������� ������� ����������� ����" Then
        iw(86) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������� �����" Then
        iw(87) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "��������" Then
        iw(88) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "����" Then
        iw(89) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "���� � ����������" Then
        iw(90) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������� �� �������� �� ������ ������������" Then
        iw(91) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "��������� �� �/� ��������������� ��������" Then
        iw(92) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "��������� �� ���. ����� ���������" Then
        iw(93) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "��������� �� ���������" Then
        iw(94) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������� ������������� �������� ����������" Then
        iw(95) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������� �� �������� �� ������ ������������-��������������� �����" Then
        iw(96) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "��������� �� �������� �� ��������" Then
        iw(97) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "��������� �� �������� 73,03" Then
        iw(98) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "��������� �� ���. ����� ����. ������" Then
        iw(99) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "��������� �� ��������������� ���������" Then
        iw(100) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������" Then
        iw(101) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "��� �� ����������" Then
        iw(102) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "��� � ����������" Then
        iw(103) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�����" Then
        iw(104) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "���" Then
        iw(105) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "��� ��" Then
        iw(106) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "% ��������� �������" Then
        iw(107) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "���� �������" Then '-
        iw(108) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������� �� ���� ������ �� ������ ��������������� �����" Then '-
        iw(109) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ �� �������������� ���� (���) ������ ������" Then '-
        iw(110) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������������" Then '-
        iw(111) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ ����������� ����� ������ (��)" Then '-
        iw(112) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ ��������" Then '-
        iw(113) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ �� ������ (�� �����)" Then '-
        iw(114) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "����� �� �����" Then '-
        iw(115) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������� �� ��������� � ������������� (�� �����)" Then '-
        iw(116) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������� �� ��������� � ������������� (��������������� ������������ ����)" Then '-
        iw(117) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "����������� ������" Then '-
        iw(118) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "����������� �������� �� ��������� �������" Then '-
        iw(119) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ �������� (� ������ ��)" Then '-
        iw(120) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ �����������" Then '-
        iw(121) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ ����������� (� ������ ��)" Then '-
        iw(122) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ �� ������ ���� (� ������ ��)" Then '-
        iw(123) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������� �� ��������� � ������������� (�� ����� �������. ������������� �������)" Then '-
        iw(124) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ ������� (� ������ ��)" Then '-
        iw(125) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "���� ����������" Then '-
        iw(126) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ ��������" Then '-
        iw(127) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������� �� ������ � ������ ����� (����������� � �������� ���)" Then '-
        iw(134) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ ����������� (� ������ ��)" Then '-
        iw(135) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������ �����������" Then '-
        iw(136) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "�������� �� ��������������� � �������� ����������� ����� (�� ����� �������. �����. �������)" Then '-
        iw(137) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "������� �����" Then '-
        iw(138) = i
    End If
    

Next i

'�������� ���������� ������
 ThisWorkbook.Sheets(SheetName).Activate
awLastRow = Cells(Rows.Count, "A").End(xlUp).row
awLastCol = Cells(begin, Columns.Count).End(xlUp).column
Range(Cells(begin, 1), Cells(awLastRow, Limit)).Select
 With Selection
        .Clear
 End With

 '������ ���
Application.StatusBar = "������� ������"

 '������� ������
 importWB.Sheets(1).Activate
 iwLastRow = Cells(Rows.Count, "A").End(xlUp).row

For i = 1 To Limit
'������ ���
Application.StatusBar = "������������� ����. ���������: " & Int(100 * i / Limit) & "%." & _
" ����� ��������: " & Int(87 * i / Limit) & "%" & _
" ��������� ����� �� ����� ���������� �����: " & _
Int((100 - Int(87 * i / Limit)) * (((Now() - Start) * 24 * 60 * 60) / (Int(87 * i / Limit)))) & " ������"
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
'    If I = Int(Limit / 4) Then
'        '���������
'        MsgBoxEx "������� ������� ..." _
'    & vbCr & "��� �������, ����������� ��������, �������� ���������", 0, "��������� " & _
'    Int(87 * I / Limit) & "%", 5
'    End If
'
'    If I = Int(Limit / 2) Then
'        '���������
'        MsgBoxEx "�� � �������" _
'    & vbCr & "��������� " & Int(87 * I / Limit) & "%", 0, "�������� �������� �������������� �����", 5
'    End If
'
'    If I = Int(Limit / 4 * 3) Then
'        '���������
'        MsgBoxEx "�������������� " & iwLastRow & " ����� � " & Limit & " �������" _
'    & vbCr & "��������� " & Int(87 * I / Limit) & "%", 0, "�����...", 5
'    End If

Next i

'������ ���
Application.StatusBar = "�������������� �����. ���������: 87 %"

'�������
ThisWorkbook.Sheets(SheetName).Activate
Columns("Q:DD").Select
    Selection.NumberFormat = _
        "_-* #,##0.00 _?_-;-* #,##0.00 _?_-;_-* ""-""?? _?_-;_-@_-"
Columns("ED").Select
    Selection.NumberFormat = _
        "_-* #,##0.00 _?_-;-* #,##0.00 _?_-;_-* ""-""?? _?_-;_-@_-"

'������� ����������� ������
ThisWorkbook.Sheets(SheetName).Activate
'������ ���
Application.StatusBar = "���������� ����������� ������ ������. ���������: 88 %"
'�����
Cells(begin, aw(2)).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(IFERROR(SEARCH("" 20"",RC[-1],1)>0,FALSE)," _
        & "TRIM(MID(IF(IFERROR(SEARCH("" 20"",RC[-1],1)>0,FALSE)," _
        & "RC[-1],R[-1]C),1,SEARCH("" "",IF(IFERROR(SEARCH("" 20"",RC[-1],1)>0,FALSE)," _
        & "RC[-1],R[-1]C),1)-1)),R[-1]C)"
    Cells(begin, aw(2)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(2)), Cells(iwLastRow, aw(2)))

'������ ���
Application.StatusBar = "���������� ������ ��������� ����� �����. ���������: 89 %"
'��������� ����� �����
K = 3
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(CONCATENATE(RC[-1],"" "",RC[5])=RC[-2]," _
        & "VLOOKUP(RC[-2],RC1:RC114,MATCH(R8C4,R9C1:R9C114,0),0)>0," _
        & "VLOOKUP(RC[-2],RC1:RC114,MATCH(R7C4,R9C1:R9C114,0),0)>0," _
        & "RC[4]=TRUE),"""",VLOOKUP(RC[-1],INDIRECT(CONCATENATE(""'"",VALUE(RC[5])," _
        & """ �����. ���������'!$A:$BM"")),HLOOKUP(RC[20]," _
        & "INDIRECT(CONCATENATE(""'"",VALUE(RC[5]),"" �����. ���������'!$2:$3"")),2,0),0))"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
    
'������ ���
Application.StatusBar = "���������� ������ ������� ����� �����. ���������: 90 %"
'������ �����
K = 4
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
        "=OR(RC[-1]="""",OR(RC[-1]=VALUE(RC[21]),VLOOKUP(RC[-3],RC1:RC114," _
        & "MATCH(R8C4,R9C1:R9C114,0),0)>0))"
    Range("A1").Copy
    Cells(begin, aw(K)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
   
'������ ���
Application.StatusBar = "���������� ������ ����� ����. ���������: 91 %"
'��������� ����� ����
K = 5
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(CONCATENATE(RC[-3],"" "",RC[3])=RC[-4]," _
        & "VLOOKUP(RC[-4],RC1:RC114,MATCH(R8C4,R9C1:R9C114,0),0)>0," _
        & "RC[2]=TRUE),"""",VLOOKUP(RC[-3],INDIRECT(CONCATENATE(""'"",VALUE(RC[3])," _
        & """ �����. ���������'!$A$18:$BM$31"")),HLOOKUP(RC[18]," _
        & "INDIRECT(CONCATENATE(""'"",VALUE(RC[3]),"" �����. ���������'!$18:$19"")),2,0),0))"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
    
'������ ���
Application.StatusBar = "���������� ������ ������� ����� ����. ���������: 92 %"
'������ ����
K = 6
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
        "=OR(RC[-3]="""",OR(RC[-1]=VALUE(RC[18])," _
        & "VLOOKUP(RC[-5],RC1:RC114,MATCH(R8C4,R9C1:R9C114,0),0)>0))"
    Range("A1").Copy
    Cells(begin, aw(K)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))

'������ ���
Application.StatusBar = "���������� ������ ���������� �� ������� ���� ����� 20,26,44 ������. ���������: 93 %"
'���������� ���� ����� 20,26,44 �����
K = 7
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
       "=NOT(OR(IFERROR(SEARCH(20,RC[15],1),FALSE)," _
       & "IFERROR(SEARCH(26,RC[15],1),FALSE),IFERROR(SEARCH(44,RC[15],1),FALSE)))"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
    
'������ ���
Application.StatusBar = "���������� ������ ��������� ����� �������� �����������. ���������: 94 %"
'��� ��������
K = 8
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(CONCATENATE(RC[-12],"" "",RC[-6])=RC[-13],""""," _
        & "(MID(RC[-13],SEARCH("" "",RC[-13],1),LEN(RC[-13]))))"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))

'������ ���
Application.StatusBar = "���������� ������ � �������������� ������� ��������� �������. ���������: 95 %"
'������ ��������� �������
K = 9
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-14],RC[-13],RC[5])"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
        
Columns("O:O").Select
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Font
        .Color = -16754788
'        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10284031
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False


'������ ���
Application.StatusBar = "���������� ������ �� ������ � ����������� �����. ���������: 96 %"
'����� �� ����� � ����������� �����
K = 78
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[1]:RC[9])"
    Cells(begin, aw(K)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 12517371
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = -13680896
        .TintAndShade = 0
    End With
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
    
'������ ���
Application.StatusBar = "���������� ������ ��������� ����. ���������: 97 %"
'���
K = 128
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
       "=VALUE(IF(IFERROR(SEARCH("" 20"",RC[-7],1)>0,FALSE)," _
        & "MID(RC[-7],SEARCH("" "",RC[-7],1)+1,4),R[-1]C))"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))


'������ ���
Application.StatusBar = "�������������� ����������. ���������: 98 %"

Columns("A:A").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = -1
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

Range("C:C,E:E").Select
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    
'������ ���
Application.StatusBar = "�������������� ����������. ���������: 99 %"
' �������������� ������� ���� ��������� �������
Columns("EC:EC").Select
    Selection.NumberFormat = _
        "_-* #,##0.00 _?_-;-* #,##0.00 _?_-;_-* ""-""?? _?_-;_-@_-"
Columns("EF:EF").Select
    Selection.NumberFormat = _
        "_-* #,##0.00 _?_-;-* #,##0.00 _?_-;_-* ""-""?? _?_-;_-@_-"

'������ ���
Application.StatusBar = "���������: 100 %"

'����������
ThisWorkbook.Sheets(SheetName).Activate
MsgBoxEx "��������� ��������� " _
    & "�� �������� " & vbCr & ThisWorkbook.Sheets("Preferences").Range("C7").Value2 _
    & vbCr & "�� " & ThisWorkbook.Sheets(SheetName).Range("B2").Value2 & " ���" _
    & vbCr & "��������� �������", 0, "���������", 25

ThisWorkbook.Sheets("Calculation22").Activate

ExitHandler:
    On Error Resume Next
    importWB.Close
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    Application.Calculation = xlAutomatic
 ThisWorkbook.Sheets("Preferences").Activate
 Exit Sub
  
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub





















