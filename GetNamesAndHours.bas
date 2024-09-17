Attribute VB_Name = "GetNamesAndHours"
Public Sub GetDistinctNames()

'' ��������� ������� ����� ����������
'    Dim originalCalculationMode As XlCalculation
'    originalCalculationMode = Application.Calculation

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
'Application.Calculation = xlManual

    Dim FullNameColumn As Range
    Dim dimension As Integer
    Dim ThisWorkbook, importWB As Workbook
    Dim DistinctList As Variant
    Dim FilesToOpen
    Dim WorkbookLinks As Variant
    Dim Wb As Workbook
'    Dim n As Variant
    Dim i As Long
    
FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsb), *.xlsb", _
 MultiSelect:=True, Title:="�������� ������ �� �����������")
    
    'Set FullNameColumn = ActiveSheet.UsedRange.Columns(4) ' �������� ������ �������.
    Set ThisWorkbook = ActiveWorkbook
    SheetName = "��"
    ThisWorkbook.Sheets(SheetName).Activate
    On Error Resume Next
    ActiveSheet.ShowAllData
    Range("B4:B103").ClearContents
    
Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

On Error Resume Next
importWB.Activate
Set FullNameColumn = ActiveSheet.UsedRange.Range("D5:D150") ' �������� ��������.

    DistinctList = GetDistinctItems(FullNameColumn) ' �������� �������� � �������.
'    Debug.Print Join(DistinctList, vbCrLf) ' ������� ���������.
    dimension = UBound(DistinctList, 1) '������ �������

ThisWorkbook.Sheets(SheetName).Activate
  
  '��� ������ �� ��������� �������
  For i = 1 To dimension
  '����������� ��� ����� ����� �� ����� ���� ������ 5 ������
    If Len(DistinctList(i)) > 5 Then
    '��������� �������� � ����� ������
      j = j + 1
    DistinctList(j) = DistinctList(i)
    Debug.Print DistinctList(j)

'������ �������� �� ���� � ������ B4
    Range("B" & i + 3).Value2 = DistinctList(j)
    End If
  Next

'���������� ������� � ������ ������
Range("J4").Select
ActiveCell.FormulaR1C1 = "=SUMIFS('[" & importWB.Name & "]��������������'!C6," _
& "'[" & importWB.Name & "]��������������'!C10,RC[12]," _
& "'[" & importWB.Name & "]��������������'!C11,RC[9]," _
& "'[" & importWB.Name & "]��������������'!C4,RC[-8])"

'������� ������ � �������� � ������� �������
Range("J4").Select
Selection.Copy

'��������� ������� ������� ������
Range("J4:J103,J105:J204,J206:J305,J307:J406," _
& "J408:J507,J509:J608,J610:J709,J711:J810,J812:J911," _
& "J913:J1012,J1014:J1113,J1115:J1214,J1216:J1315,J1317:J1416," _
& "J1418:J1517,J1519:J1618,J1620:J1719,J1721:J1820,J1822:J1921," _
& "J1923:J2022,J2024:J2123,J2125:J2224,J2226:J2325").Select

'���������� �������� � ���� ���������� ��������
Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
' ������������� ���������� ��������
'    Set selectedRange = Range("K4:K2428")
'' ��������� �������� ������ ��� ����������� ���������
'    If Not selectedRange Is Nothing Then
'        selectedRange.Calculate
'    End If
'
'������� ������ ������������ ������ �� ������� ���������
With ActiveSheet.ListObjects("��").Range
    .AutoFilterMode = False ' ����� �������� �������
    .AutoFilter Field:=11, Criteria1:="<>0" _
        , Operator:=xlAnd ' ��������� �� 11-�� ����
End With
    

'�������� ����� � �������� �������
Set Wb = ActiveWorkbook
WorkbookLinks = Wb.LinkSources(Type:=xlLinkTypeExcelLinks)
If Not IsEmpty(WorkbookLinks) Then
    For i = LBound(WorkbookLinks) To UBound(WorkbookLinks)
        On Error Resume Next ' ���������� ������
        Wb.BreakLink Name:=WorkbookLinks(i), Type:=xlLinkTypeExcelLinks
        On Error GoTo 0 ' ���������� ���������� ��������� ������
    Next i
End If
  
''importWB.Close
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
'    Application.Calculation = xlAutomatic
 ThisWorkbook.Sheets("Preferences").Activate
End Sub

Public Function GetDistinctItems(ByRef Range As Range) As Variant
    Dim Data As Variant: Data = Range.Value ' ����������� �������� � ������.
    Dim Buffer As Object: Set Buffer = CreateObject("System.Collections.ArrayList") ' ������� ������ ArrayList.

    Dim Item
    For Each Item In Data
        If Not Buffer.Contains(Item) Then Buffer.Add Item ' ��������� ������� �������� � ��������� ���� �����������.
    Next

    Buffer.Sort ': Buffer.Reverse ' ��������� �� �����������, � ����� �������������� (�� ��������).
    GetDistinctItems = Buffer.ToArray() ' ��������� � ���� �������.
End Function
