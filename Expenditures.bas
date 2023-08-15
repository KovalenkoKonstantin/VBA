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
 Limit = 134 '��������� ������� ����
 begin = 12 '������ ��� �������
 CompanyName = ThisWorkbook.Sheets("Preferences").Range("C7").Value2 '��� �������
 
 Dim aw(1 To 134) As Variant
 Dim iw(1 To 134) As Variant
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False

FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="�������� ��������� ��������� �� �������� " & CompanyName)
 
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
For I = 1 To 20
    If Worksheets(SheetName).Cells(I, 1) = "���������" Then
        DataRow = I
    End If
Next I

For I = 1 To Limit
    If Worksheets(SheetName).Cells(DataRow, I) = "���������" Then
        aw(1) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�����" Then
        aw(2) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "��������� ����� �����" Then
        aw(3) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ �����" Then
        aw(4) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "��������� ����� ����" Then
        aw(5) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ ����" Then
        aw(6) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���������� ���� ����� 20,26,44 �����" Then
        aw(7) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "��� ��������" Then
        aw(8) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ ��������� �������" Then
        aw(9) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ �� ����������.������ �� ��" Then
        aw(10) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������������� �������" Then
        aw(11) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���������" Then
        aw(12) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "��� ���������" Then
        aw(13) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���� ��������" Then
        aw(14) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ ��������� �������� � ��� �����" Then
        aw(15) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ ������" Then
        aw(16) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "����� ����" Then
        aw(17) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "����� �����" Then
        aw(18) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���������� ����" Then
        aw(19) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���������� �����" Then
        aw(20) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, I) = "���������" Then
        aw(21) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ ���������� ������ �� ���� ������������" Then
        aw(22) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ ���������� ������" Then
        aw(23) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "����� �� ����" Then
        aw(24) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���������� �� ������� (���������� ��� �� ������)" Then
        aw(25) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������� �� ��������� � �������������" Then
        aw(26) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ ������� �� ����������� ����" Then
        aw(27) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ �� ����� ������" Then
        aw(28) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������� �� ��������������� � �������� ����������� �����" Then
        aw(29) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ �� ������ ����" Then
        aw(30) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ �� ���� ����" Then
        aw(31) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "����������� ������� (������ ��������)" Then
        aw(32) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "����������� ������" Then
        aw(33) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ �������" Then
        aw(34) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������� ������� ��� ����������" Then
        aw(35) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� �� ����� �� �������� �� �������� ���" Then
        aw(36) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "����� �� ���� (��������������� ������������ ����)" Then
        aw(37) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������� �� ��������������� � �������� ����������� ����� (��������������� ������������ ����)" Then
        aw(38) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������� �� ������ �� ����������, ������������� ��������������� �����" Then
        aw(39) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "����� �� ���� 26 �����������" Then
        aw(40) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ �� �����������" Then
        aw(41) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������� ��������" Then
        aw(42) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������� �����������" Then
        aw(43) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ � ������ ��������� ������������(�����������)" Then
        aw(44) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "��������� �������������� ������������ ������" Then
        aw(45) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ �� ������ ���� � ������ ��" Then
        aw(46) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������������� �� �����������" Then
        aw(47) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� (������, ������)" Then
        aw(48) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "����������� �� ������" Then
        aw(49) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ � ������ ��������� ������������ (��������)" Then
        aw(50) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������" Then
        aw(51) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� �� ���������� ����������, ���������� ������������" Then
        aw(52) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������� ������" Then
        aw(53) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� �����" Then
        aw(54) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������������� ������� ������ (������������)" Then
        aw(55) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ ������ � ����������� � �������� ���" Then
        aw(56) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������������� ������ ������ ����������" Then
        aw(57) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ �������� �.�" Then
        aw(58) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���������� �� ������������ �������" Then
        aw(59) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ ��� ������ �������� �� ��" Then
        aw(60) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ ���� ����� �� ������-����������" Then
        aw(61) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� �� ����� �� ������� �� 3 ��� ��� ������" Then
        aw(62) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "����� �� ����� (������������� ������������� �������)" Then
        aw(63) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������� �� ��������� � ������������� (�� ����� ��������������� ������������ �������)" Then
        aw(64) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������� �� ��������������� � �������� ����������� ����� (�� ����� ������� ����� �������)" Then
        aw(65) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���������� �� �������" Then
        aw(66) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� �� ������ � ������ �����" Then
        aw(67) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�� ������ ������ �� ���" Then
        aw(68) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ ���������� (���������, �����, ����������, �������)" Then
        aw(69) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������� �� ���������� ������������������ � ��������� �" Then
        aw(70) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ �� ������������ � �����" Then
        aw(71) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� �� ������ � ����������� ��� (������� �����)" Then
        aw(72) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� �� ������ � ����������� ��� (������ �����)" Then
        aw(73) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� �� ����������� ��� ������������� ����� �������� �������" Then
        aw(74) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������������� ��� ������� �� ������������ � �����" Then
        aw(75) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� �� ��������������� ������� ����" Then
        aw(76) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������������ ������" Then
        aw(77) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, I) = "����� �� ����� � ����������� �����" Then
        aw(78) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� �� ���� ������� �����������" Then
        aw(79) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "����� � ����������� ����� (���. ����)" Then
        aw(80) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���������" Then
        aw(81) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���� � ������ � ���� ������� ��������" Then
        aw(82) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� ������������" Then
        aw(83) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������" Then
        aw(84) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "��������� ��������" Then
        aw(85) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� ������� ����������� ����" Then
        aw(86) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� �����" Then
        aw(87) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, I) = "��������" Then
        aw(88) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "����" Then
        aw(89) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���� � ����������" Then
        aw(90) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������� �� �������� �� ������ ������������" Then
        aw(91) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "��������� �� �/� ��������������� ��������" Then
        aw(92) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "��������� �� ���. ����� ���������" Then
        aw(93) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "��������� �� ���������" Then
        aw(94) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� ������������� �������� ����������" Then
        aw(95) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������� �� �������� �� ������ ������������-��������������� �����" Then
        aw(96) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "��������� �� �������� �� ��������" Then
        aw(97) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "��������� �� �������� 73,03" Then
        aw(98) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "��������� �� ���. ����� ����. ������" Then
        aw(99) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "��������� �� ��������������� ���������" Then
        aw(100) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, I) = "������" Then
        aw(101) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "��� �� ����������" Then
        aw(102) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "��� � ����������" Then
        aw(103) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�����" Then
        aw(104) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���" Then
        aw(105) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "��� ��" Then
        aw(106) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, I) = "% ��������� �������" Then
        aw(107) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���� �������" Then
        aw(108) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, I) = "�������� �� ���� ������ �� ������ ��������������� �����" Then
        aw(109) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ �� �������������� ���� (���) ������ ������" Then
        aw(110) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������������" Then
        aw(111) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ ����������� ����� ������ (��)" Then
        aw(112) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ ��������" Then
        aw(113) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ �� ������ (�� �����)" Then
        aw(114) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "����� �� �����" Then
        aw(115) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������� �� ��������� � ������������� (�� �����)" Then
        aw(116) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������� �� ��������� � ������������� (��������������� ������������ ����)" Then
        aw(117) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "����������� ������" Then
        aw(118) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "����������� �������� �� ��������� �������" Then
        aw(119) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ �������� (� ������ ��)" Then
        aw(120) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ �����������" Then
        aw(121) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ ����������� (� ������ ��)" Then
        aw(122) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ �� ������ ���� (� ������ ��)" Then
        aw(123) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "�������� �� ��������� � ������������� (�� ����� �������. ������������� �������)" Then
        aw(124) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ ������� (� ������ ��)" Then
        aw(125) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���� ����������" Then
        aw(126) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������ ��������" Then
        aw(127) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "���" Then
        aw(128) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "������� �� ������ � ������ ����� (����������� � �������� ���)" Then
        aw(134) = I
    End If
    
Next I
 
 importWB.Sheets(1).Activate

'����������� ������� ������������� �����
For I = 1 To 20
    If importWB.Sheets(1).Cells(I, 1) = "�����������" Then
        ImportFirstDataRow = I
    End If
Next I
For I = 1 To 20
    If importWB.Sheets(1).Cells(I, 1) = "���������" Then
        ImportSecondDataRow = I
    End If
Next I

For I = 1 To Limit
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "���������" Then '-
        iw(1) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�����" Then
        iw(2) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "��������� ����� �����" Then
        iw(2) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ �����" Then
        iw(4) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "��������� ����� ����" Then
        iw(5) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ ����" Then
        iw(6) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "���������� ���� ����� 20,26,44 �����" Then
        iw(7) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "��� ��������" Then
        iw(8) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ ��������� �������" Then
        iw(9) = I
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "������ �� ����������.������ �� ��" Then
        iw(10) = I
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "������������� �������" Then '-
        iw(11) = I
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "���������" Then '-
        iw(12) = I
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "��� ���������" Then '-
        iw(13) = I
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "���� ��������" Then '-
        iw(14) = I
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "������ ��������� �������� � ��� �����" Then '-
        iw(15) = I
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "������ ������" Then '-
        iw(16) = I
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "����� ����" Then '-
        iw(17) = I
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "����� �����" Then '-
        iw(18) = I
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "����" Then '-
        iw(19) = I
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "�����" Then '-
        iw(20) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "���������" Then
        iw(21) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ ���������� ������ �� ���� ������������" Then
        iw(22) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ ���������� ������" Then
        iw(23) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "����� �� ����" Then
        iw(24) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "���������� �� ������� (���������� ��� �� ������)" Then
        iw(25) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�������� �� ��������� � �������������" Then
        iw(26) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ ������� �� ����������� ����" Then
        iw(27) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ �� ����� ������" Then
        iw(28) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�������� �� ��������������� � �������� ����������� �����" Then
        iw(29) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ �� ������ ����" Then
        iw(30) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ �� ���� ����" Then
        iw(31) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "����������� ������� (������ ��������)" Then
        iw(32) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "����������� ������" Then
        iw(33) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ �������" Then
        iw(34) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�������� ������� ��� ����������" Then
        iw(35) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������� �� ����� �� �������� �� �������� ���" Then
        iw(36) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "����� �� ���� (��������������� ������������ ����)" Then
        iw(37) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�������� �� ��������������� � �������� ����������� ����� (��������������� ������������ ����)" Then
        iw(38) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�������� �� ������ �� ����������, ������������� ��������������� �����" Then
        iw(39) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "����� �� ���� 26 �����������" Then
        iw(40) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ �� �����������" Then
        iw(41) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�������� ��������" Then
        iw(42) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�������� �����������" Then
        iw(43) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ � ������ ��������� ������������(�����������)" Then
        iw(44) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "��������� �������������� ������������ ������" Then
        iw(45) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ �� ������ ���� � ������ ��" Then
        iw(46) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�������������� �� �����������" Then
        iw(47) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������� (������, ������)" Then
        iw(48) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "����������� �� ������" Then
        iw(49) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ � ������ ��������� ������������ (��������)" Then
        iw(50) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������" Then
        iw(51) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������� �� ���������� ����������, ���������� ������������" Then
        iw(52) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�������� ������" Then
        iw(53) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������� �����" Then
        iw(54) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�������������� ������� ������ (������������)" Then
        iw(55) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ ������ � ����������� � �������� ���" Then
        iw(56) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�������������� ������ ������ ����������" Then
        iw(57) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ �������� �.�" Then
        iw(58) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "���������� �� ������������ �������" Then
        iw(59) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ ��� ������ �������� �� ��" Then
        iw(60) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ ���� ����� �� ������-����������" Then
        iw(61) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������� �� ����� �� ������� �� 3 ��� ��� ������" Then
        iw(62) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "����� �� ����� (������������� ������������� �������)" Then
        iw(63) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�������� �� ��������� � ������������� (�� ����� ��������������� ������������ �������)" Then
        iw(64) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�������� �� ��������������� � �������� ����������� ����� (�� ����� ������� ����� �������)" Then
        iw(65) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "���������� �� �������" Then
        iw(66) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������� �� ������ � ������ �����" Then
        iw(67) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�� ������ ������ �� ���" Then
        iw(68) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ ���������� (���������, �����, ����������, �������)" Then
        iw(69) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�������� �� ���������� ������������������ � ��������� �" Then
        iw(70) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ �� ������������ � �����" Then
        iw(71) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������� �� ������ � ����������� ��� (������� �����)" Then
        iw(72) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������� �� ������ � ����������� ��� (������ �����)" Then
        iw(73) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������� �� ����������� ��� ������������� ����� �������� �������" Then
        iw(74) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�������������� ��� ������� �� ������������ � �����" Then
        iw(75) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������� �� ��������������� ������� ����" Then
        iw(76) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������������ ������" Then
        iw(77) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "����� �� ����� � ����������� �����" Then
        iw(78) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������� �� ���� ������� �����������" Then
        iw(79) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "����� � ����������� ����� (���. ����)" Then
        iw(80) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "���������" Then
        iw(81) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "���� � ������ � ���� ������� ��������" Then
        iw(82) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������� ������������" Then
        iw(83) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�������" Then
        iw(84) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "��������� ��������" Then
        iw(85) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������� ������� ����������� ����" Then
        iw(86) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������� �����" Then
        iw(87) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "��������" Then
        iw(88) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "����" Then
        iw(89) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "���� � ����������" Then
        iw(90) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�������� �� �������� �� ������ ������������" Then
        iw(91) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "��������� �� �/� ��������������� ��������" Then
        iw(92) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "��������� �� ���. ����� ���������" Then
        iw(93) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "��������� �� ���������" Then
        iw(94) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������� ������������� �������� ����������" Then
        iw(95) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�������� �� �������� �� ������ ������������-��������������� �����" Then
        iw(96) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "��������� �� �������� �� ��������" Then
        iw(97) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "��������� �� �������� 73,03" Then
        iw(98) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "��������� �� ���. ����� ����. ������" Then
        iw(99) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "��������� �� ��������������� ���������" Then
        iw(100) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������" Then
        iw(101) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "��� �� ����������" Then
        iw(102) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "��� � ����������" Then
        iw(103) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�����" Then
        iw(104) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "���" Then
        iw(105) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "��� ��" Then
        iw(106) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "% ��������� �������" Then
        iw(107) = I
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "���� �������" Then '-
        iw(108) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�������� �� ���� ������ �� ������ ��������������� �����" Then '-
        iw(109) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ �� �������������� ���� (���) ������ ������" Then '-
        iw(110) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������������" Then '-
        iw(111) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ ����������� ����� ������ (��)" Then '-
        iw(112) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ ��������" Then '-
        iw(113) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ �� ������ (�� �����)" Then '-
        iw(114) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "����� �� �����" Then '-
        iw(115) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�������� �� ��������� � ������������� (�� �����)" Then '-
        iw(116) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�������� �� ��������� � ������������� (��������������� ������������ ����)" Then '-
        iw(117) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "����������� ������" Then '-
        iw(118) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "����������� �������� �� ��������� �������" Then '-
        iw(119) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ �������� (� ������ ��)" Then '-
        iw(120) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ �����������" Then '-
        iw(121) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ ����������� (� ������ ��)" Then '-
        iw(122) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ �� ������ ���� (� ������ ��)" Then '-
        iw(123) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "�������� �� ��������� � ������������� (�� ����� �������. ������������� �������)" Then '-
        iw(124) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ ������� (� ������ ��)" Then '-
        iw(125) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "���� ����������" Then '-
        iw(126) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������ ��������" Then '-
        iw(127) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "������� �� ������ � ������ ����� (����������� � �������� ���)" Then '-
        iw(134) = I
    End If
    

Next I

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

For I = 1 To Limit
'������ ���
Application.StatusBar = "������������� ����. ���������: " & Int(100 * I / Limit) & "%." & _
" ����� ��������: " & Int(87 * I / Limit) & "%" & _
" ��������� ����� �� ����� ���������� ���������: " & _
Int((100 - Int(87 * I / Limit)) * (((Now() - Start) * 24 * 60 * 60) / (Int(87 * I / Limit)))) & " ������"
 importWB.Activate
 Range(Cells(begin - 1, iw(I)), Cells(iwLastRow, iw(I))).Copy

 ThisWorkbook.Sheets(SheetName).Activate
 Range(Cells(begin, aw(I)), Cells(iwLastRow, aw(I))).Select
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

Next I

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
Columns("DW:DW").Select
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
 ThisWorkbook.Sheets("Preferences").Activate
 Exit Sub
  
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub





















