Attribute VB_Name = "CopyWB"
Sub Copy_W()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False

    Dim WbLinks As Variant
    Dim SaveName As String
    Dim DistinctList As Variant
    Dim FullNameColumn As Range
    Dim i As Long
    Dim Path As String
    Dim FilePath As String
    
    ' �������� ���� � ����� � ��� ��� ����������
    Path = ActiveWorkbook.Path
    SaveName = ActiveSheet.Range("H30").Text
    Set FullNameColumn = ThisWorkbook.Sheets("Preferences").Range("I2:I9").SpecialCells(xlCellTypeVisible) ' �������� �������� �������� ��� ������ �����
    
    ' �������� ���������� �������� �� ���������� ���������
    DistinctList = GetDistinctItems(FullNameColumn)
    If IsEmpty(DistinctList) Then Exit Sub ' ��������, ��� ������ �� ������

    ' ������� ������ �������� (������������� � ��������� ������� ������� ����� ����������)
    Dim NewList() As Variant
    Dim count As Long
    ReDim NewList(0)
    
    For i = LBound(DistinctList) To UBound(DistinctList)
        If Not IsEmpty(DistinctList(i)) Then
            NewList(count) = DistinctList(i)
            count = count + 1
            ReDim Preserve NewList(count) ' ����������� ������
        End If
    Next i
    If count > 0 Then ReDim Preserve NewList(count - 1) ' ������� ��������� �������� �������

    ' ��������� ����� ������� �������
    ReDim Preserve NewList(UBound(NewList) + 1)
    NewList(UBound(NewList)) = "Ninth"

    ' �������� ��������� �����
    ActiveWorkbook.Sheets(NewList).Copy
    ActiveWorkbook.PrecisionAsDisplayed = True

    
    ' ��������� �������� � ����� "�2 (1)" ����� ���������
    ActiveWorkbook.Sheets("�2 (1)").Activate
    Cells.Select
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    ' ��������� �������� � ����� "�� (1)" ����� ���������
    ActiveWorkbook.Sheets("�� (1)").Activate
    Cells.Select
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    
    ' ������� ������ �����
    On Error Resume Next ' ���������� ������ ��� ��������
    Sheets("Ninth").delete
    Sheets("��").delete
    On Error GoTo 0

    ' ��������� �����
    WbLinks = ActiveWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks)
    If Not IsEmpty(WbLinks) Then
        For i = LBound(WbLinks) To UBound(WbLinks)
            ActiveWorkbook.BreakLink Name:=WbLinks(i), Type:=xlLinkTypeExcelLinks
        Next i
    End If

    ' ��� ���������� �����
    FilePath = Path & "\" & SaveName & ".xls"
    If Dir(FilePath) <> "" Then Kill FilePath
    ActiveWorkbook.SaveAs Filename:=Path & "\" & SaveName & ".xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

    ' �������� "������" �� �������
    Dim val As String: val = "������"
    Dim FindIndex As Long
    FindIndex = -1

    For i = LBound(NewList) To UBound(NewList)
        If NewList(i) = val Then
            FindIndex = i
            Exit For
        End If
    Next i

    ' ������� ������� "������", ���� �� ��� ������
    If FindIndex <> -1 Then
        For i = FindIndex To UBound(NewList) - 1
            NewList(i) = NewList(i + 1)
        Next i
        ReDim Preserve NewList(LBound(NewList) To UBound(NewList) - 1)
    End If

    ' ����� �������� ���������� �����
    ActiveWorkbook.Sheets(1).Activate

    ' �������� ��� �������
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True

End Sub
