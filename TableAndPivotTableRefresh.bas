Attribute VB_Name = "TableAndPivotTableRefresh"
'Sub RefreshAllTables()
'    Dim ws As Worksheet
'    Dim lo As ListObject
'    Dim info As String
'    Dim pt As PivotTable
'
'    For Each ws In ThisWorkbook.Worksheets
'        For Each lo In ws.ListObjects
'        On Error Resume Next
'            info = "��� �������: " & lo.Name & vbCrLf
'            info = info & "����: " & ws.Name & vbCrLf
'            info = info & "���������� �����: " & lo.ListRows.Count & vbCrLf
'            info = info & "���������� ��������: " & lo.ListColumns.Count & vbCrLf
'            info = info & vbCrLf
'            Debug.Print info
'            lo.QueryTable.Refresh BackgroundQuery:=False
'
'            lo.TableObject.Refresh
'
'        Next lo
'    Next ws
'    For Each ws In ThisWorkbook
'        pt.Refresh
'    Next ws
'
'End Sub


Sub RefreshAllTables()
    ' ��������� ����������
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim info As String
    Dim pt As PivotTable
    
    ' �������� �� ���� ������ � ������� �����
    For Each ws In ThisWorkbook.Worksheets
        ' �������� �� ���� �������� (ListObjects) �� ������� �����
        For Each lo In ws.ListObjects
            ' ���������� ������, ����� ���������� ���������� ���� ���� ��� ������������� ������
            On Error Resume Next
            
            ' ��������� ������ � ����������� � �������
            info = "��� �������: " & lo.Name & vbCrLf
            info = info & "����: " & ws.Name & vbCrLf
            info = info & "���������� �����: " & lo.ListRows.Count & vbCrLf
            info = info & "���������� ��������: " & lo.ListColumns.Count & vbCrLf
            info = info & vbCrLf
            
            ' ������� ���������� � ���� �������
            Debug.Print info
            
            ' ��������� �������, ���� ��� ������� � ������� ���������� ������
            lo.QueryTable.Refresh BackgroundQuery:=False
            
            ' ��������� ������ �������
            lo.TableObject.Refresh
        Next lo
    Next ws
    
    ' �������� �� ���� ������ � ������� �����
    For Each ws In ThisWorkbook.Worksheets
        ' �������� �� ���� ������� �������� �� ������� �����
        For Each pt In ws.PivotTables
            ' ��������� ������� �������
            pt.RefreshTable
        Next pt
    Next ws
End Sub
