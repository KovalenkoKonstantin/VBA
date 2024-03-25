Attribute VB_Name = "GetNames"
Public Sub GetDistinctNames()

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
    
FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="�������� ������ �� �����������")
    
    'Set FullNameColumn = ActiveSheet.UsedRange.Columns(4) ' �������� ������ �������.
    Set ThisWorkbook = ActiveWorkbook
    SheetName = "��"
    ThisWorkbook.Sheets(SheetName).Activate
    Range("B4:B103").ClearContents
    
Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

On Error Resume Next
importWB.Activate
Set FullNameColumn = ActiveSheet.UsedRange.Range("D5:D150") ' �������� ��������.

    DistinctList = GetDistinctItems(FullNameColumn) ' �������� �������� � �������.
'    Debug.Print Join(DistinctList, vbCrLf) ' ������� ���������.
    dimension = UBound(DistinctList, 1) '������ �������

ThisWorkbook.Sheets(SheetName).Activate
    
  j = -1
  '��� ������ �� ��������� �������
  For i = 0 To dimension
  '����������� ��� ����� ����� �� ����� ���� ������ 5 ������
    If Len(DistinctList(i)) > 5 Then
    '��������� �������� � ����� ������
      j = j + 1
    DistinctList(j) = DistinctList(i)
    Debug.Print DistinctList(j)

'������ �������� �� ���� � ������ B4
    Range("B" & i + 4).Value2 = DistinctList(j)
    End If
  Next
  
importWB.Close
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
