Attribute VB_Name = "CopyWB"
Sub Copy_W()

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

Dim WbLinks
Dim SaveName As String
Dim DistinctList As Variant
Dim FullNameColumn As Range

Path = ActiveWorkbook.Path
SaveName = ActiveSheet.Range("H30").Text
ThisWorkbook.Sheets("Preferences").Activate
Set FullNameColumn = ActiveSheet.UsedRange.Range("I2:I20") ' �������� ��������. � ������� ��������
DistinctList = GetDistinctItems(FullNameColumn) ' �������� �������� � �������.


'������ �������� ������� ���������� ������ ������
n = LBound(DistinctList) ' ��������� ������� �� ������ �.�. ������ ������������ ��������
For i = n To UBound(DistinctList) - 1
    DistinctList(i) = DistinctList(i + 1)
Next
ReDim Preserve DistinctList(LBound(DistinctList) To i - 1)


'�����
Debug.Print Join(DistinctList, vbCrLf) ' ������� ���������.
Debug.Print ("____________________________")
'�������� ����� ������� �������
ReDim Preserve DistinctList(UBound(DistinctList) + 1)
'����� ��� ������ ��������
DistinctList(UBound(DistinctList)) = "Ninth"
'�����
Debug.Print Join(DistinctList, vbCrLf) ' ������� ���������.

ActiveWorkbook.Sheets(DistinctList).Copy
'�������� ��� �� ������
ActiveWorkbook.PrecisionAsDisplayed = True
'������ ��������� ������� �������
ReDim Preserve DistinctList(UBound(DistinctList) - 1)

Sheets(DistinctList).Select

'������ ����� �����
Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
'������ ������ �����
Sheets("Ninth").delete
Sheets("������").delete

'�������� �����
WbLinks = ActiveWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks)
If Not IsEmpty(WbLinks) Then
    For i = LBound(WbLinks) To UBound(WbLinks)
        ActiveWorkbook.BreakLink Name:=WbLinks(i), Type:=xlLinkTypeExcelLinks
    Next
Else
End If

'�� �������� ������� �������� ������� ��������� ���� �����
ActiveWorkbook.Sheets(DistinctList(LBound(DistinctList))).Select

'�������� ����� ���� ��� ����������
FilePath = Path & "\" & SaveName & ".xls"
If Dir(FilePath) <> "" Then
    Kill FilePath
    ActiveWorkbook.SaveAs Filename:=Path & "\" & _
    SaveName & ".xls"
Else
    ActiveWorkbook.SaveAs Filename:=Path & "\" & _
    SaveName & ".xls"
End If

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True

'NewWb.Close
    
End Sub
