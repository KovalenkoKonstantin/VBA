Attribute VB_Name = "PageCount"
Public Sub CountPages()
Dim ThisWorkbook As Workbook
Set ThisWorkbook = ActiveWorkbook

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

On Error Resume Next

X = 0
ThisWorkbook.Sheets("Preferences").Activate
DistinctColumn = Rows(1).Find("����������", LookIn:=xlValues).column '������� ����������
LogicColumn = Rows(1).Find("��������", LookIn:=xlValues).column '������� ����������
ThisWorkbook.Sheets("�����").Activate
NameColumn = 3 'Rows(5).Find("������������ ���������", LookIn:=xlValues).column '������� ������������

i = "1"
j = "�������� (������������) ���� (����������*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
    X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
'[d7] = X
X = 0

'i = "2"
j = "�������� ����������� ������ (�. � 2*"
'ThisWorkbook.Sheets("Preferences").Activate
'RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
'If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
'X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
'End If
'i = "2_22"
'ThisWorkbook.Sheets("Preferences").Activate
'RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
'If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
'X = X + ThisWorkbook.Sheets(i).PageSetup.Pages.Count
'End If
'i = "2_23"
'ThisWorkbook.Sheets("Preferences").Activate
'RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
'If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
'X = X + ThisWorkbook.Sheets(i).PageSetup.Pages.Count
'End If
'i = "2_24"
'ThisWorkbook.Sheets("Preferences").Activate
'RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
'If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
'X = X + ThisWorkbook.Sheets(i).PageSetup.Pages.Count
'End If
'i = "2_1"
'ThisWorkbook.Sheets("Preferences").Activate
'RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
'If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
'X = X + ThisWorkbook.Sheets(i).PageSetup.Pages.Count
'End If
'i = "2_1_22"
'ThisWorkbook.Sheets("Preferences").Activate
'RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
'If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
'X = X + ThisWorkbook.Sheets(i).PageSetup.Pages.Count
'End If
'i = "2_1_23"
'ThisWorkbook.Sheets("Preferences").Activate
'RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
'If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
'X = X + ThisWorkbook.Sheets(i).PageSetup.Pages.Count
'End If
i = "2_2"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
X = X + ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
i = "2_2_23"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
X = X + ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
i = "2_2_24"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
X = X + ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
'[d8] = X
X = 0

'[d9] = 0
'[d10] = 0
'[d11] = 0
'[d12] = 0

i = "7_2"
j = "����������� ������ �� ������� (�������), ����������� (�����������) ���������� ������������� (�. � 7 (7�) ��� (���)"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
    X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
i = "7_2_23"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
    X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
i = "7_2_24"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
    X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
'[d16] = X
X = 0


'[d17] = 0
'[d18] = 0
'[d19] = 0

i = "9_2"
j = "����������� �������� ���������� ����� (�. � 9 (9�)*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
i = "9_2_23"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
'[d20] = X
X = 0
i = "9_2_24"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
'[d20] = X
X = 0

i = "10"
j = "������-����������� ������ (*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
'[d21] = X
X = 0

'[d22] = 0

i = "12"
j = "����� � ������ ����������������� ������/���������������-�������������� ��������*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
'[d23] = X
X = 0
'
'[d24] = 0
'[d25] = 0
'[d26] = 0
'[d27] = 0
'[d28] = 0
'[d29] = 0
'[d30] = 0
'[d31] = 0
'[d32] = 0

i = "18"
j = "����������� ������ ������ ������ (�. � 18 (18�*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
'[d33] = 0
X = 0

'[d34] = 0

'i = "20"
j = "������ � ����������� ������� (�. � 20 (20�*"
'ThisWorkbook.Sheets("Preferences").Activate
'RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
'If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
'X = ThisWorkbook.Sheets("20").PageSetup.Pages.Count
'End If
'i = "20_1"
'ThisWorkbook.Sheets("Preferences").Activate
'RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
'If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
'    X = X + ThisWorkbook.Sheets(i).PageSetup.Pages.Count
'End If
i = "20_2"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
    X = X + ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
i = "20_2_23"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
    X = X + ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
i = "20_2_24"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
    X = X + ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
'[d35] = X
X = 0

i = "21�"
j = "�������� �� ������� �������� ���������, � �. �. �� ���, ������� ���������� �������� (�. � 21*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
    X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
'[d36] = X
X = 0

i = "22�"
j = "�������� � ���������� � ������������� ����������� �����������*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
    X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
X = 0

i = "23�"
j = "������ (�����������) ������������ (�. � 23)"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
'[d38] = X
X = 0

i = "������"
j = "������� (�������� ������������ ���������������*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
'[d39] = X
X = 0

i = "�5"
j = "������� ��������� � ��������� ���*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
'[d40] = X
X = 0

i = "��"
j = "�������� � ��������� ������� ���������*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
'[d41] = X
X = 0

i = "�6"
j = "������� �� ����������� � �����*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
'[d42] = X
X = 0

i = "�7"
j = "������� �� ������ *"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
'[d43] = X
X = 0

'[d44] = 0

j = "������� ��������"
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = 40

'[d46] = 0

i = "�8"
j = "����������� ����������������� ������ ��. *"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
'[d47] = X
X = 0

i = "������"
j = "������ �����*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
'[d48] = X
X = 0

i = "��"
j = "�������������*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
'[d49] = X
X = 0

i = "���_2"
j = "�������� �����*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
'[d50] = X
X = 0

i = "��1"
j = "������� ���������� �� 01.06.2022 �."
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
    X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
X = 0

i = "��2"
j = "������� ���������� �� 01.03.2023 �."
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
    X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
X = 0

i = "��3"
j = "������� ���������� �� 01.04.2023 �."
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(i, LookIn:=xlValues).row '��� ��������
If ThisWorkbook.Sheets("Preferences").Range(LogicColumn & RowData).Value2 = True Then
    X = ThisWorkbook.Sheets(i).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("�����").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row '��� �����
ThisWorkbook.Sheets("�����").Cells(RowKey, NameColumn + 1).Value2 = X
X = 0

'[d52] = 7
'[d53] = 6
'[d54] = 1
'[d55] = 0
'[d56] = 7
'[d57] = 2
'[d58] = 2
'[d59] = 17

Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True

End Sub

