Attribute VB_Name = "GetNames"
Public Sub GetDistinctNames()
    Dim FullNameColumn As Range
    Dim dimension As Integer
    Dim ThisWorkbook, importWB As Workbook
    Dim DistinctList As Variant
    
    'Set FullNameColumn = ActiveSheet.UsedRange.Columns(4) ' �������� ������ �������.
    
    Set ThisWorkbook = ActiveWorkbook
    SheetName = "��"
    ThisWorkbook.Sheets(SheetName).Activate
    Range("B4:B103").ClearContents
    
FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="�������� ������ �� �����������")

Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

On Error Resume Next
importWB.Activate
Set FullNameColumn = ActiveSheet.UsedRange.Range("D5:D100") ' �������� ��������.

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
'    DistinctList(j) = DistinctList(i)
'    Debug.Print DistinctList(j)

'������ �������� �� ���� � ������ B4
    Range("B" & i + 3).Value2 = DistinctList(j)
    End If
    
  Next
'    Debug.Print DistinctList
    
'    Range("E1").Resize(j + 1, 1) = WorksheetFunction.Transpose(DistinctList)
'    DistinctList = Range(Range("E1"), Range("E" & Rows.Count).End(xlUp)).Resize(, 1).Value
'    Range("E1").Resize(dimension - 2, 1).Select
'    Range("E1").Resize(dimension, 1).Value2 = WorksheetFunction.Transpose(DistinctList)
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
