VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "����8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim dropdownCell As Range
    Dim conditionCell As Range


    ' ���������� ������ � ���������� ������� � ������ �������
    Set dropdownCell = Me.Range("E3") ' ������ � ���������� �������
    Set conditionCell = Me.Range("E1") ' ������ ��� �������� �������

    ' ���������, �������� �� ������ � ���������� �������
    If Not Intersect(Target, dropdownCell) Is Nothing Then
        ' ��������� �������� � ������ E1
        If conditionCell.Value = False Then
            ' ���� �������� � E1 ����� False, ��������� ��������
            RefreshAllTables
        End If
    End If
    
    ' ��������� �������� � ������� ��������
    With Sheets("������ �� ""��""")
                .Range("B4").AutoFill Destination:=.Range("B4:B500"), Type:=xlFillDefault
    End With
    Sheets("��").Select
    
End Sub
