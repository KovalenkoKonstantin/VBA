Attribute VB_Name = "CopyList"
Sub Clone9()
  Dim kolvo As Variant
  Dim i As Long
  Dim list As Worksheet
  Dim ThisWorkbook As Workbook
  Dim SheetName As String
  Dim Sht As Worksheet
  
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False
  
  Set ThisWorkbook = ActiveWorkbook
  SheetName = "9"
  kolvo = 14
  
  '�������� ���������� ������
  
  

  On Error Resume Next
  
  For i = 1 To kolvo
    Sheets("9" & i).delete
    '������ ���
    Application.StatusBar = "�������� ������. ������ ��������. " & _
    "���������: " & Int(100 * i / kolvo) & "%."
  Next i
  
  For i = 1 To 4
    Sheets("9_2" & i).delete
    '������ ���
    Application.StatusBar = "�������� ������. ������ ��������. " & _
    "���������: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = 1 To 2
    Sheets("9_" & i).delete
    '������ ���
    Application.StatusBar = "�������� ������. ������ ��������. " & _
    "���������: " & Int(100 * i / 2) & "%."
  Next i
  
  For i = 1 To 4
    Sheets("9_1_2" & i).delete
    '������ ���
    Application.StatusBar = "�������� ������. �������� ��������. " & _
    "���������: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = 1 To 4
    Sheets("9_2_2" & i).delete
    '������ ���
    Application.StatusBar = "�������� ������. ����� ��������. " & _
    "���������: " & Int(100 * i / 4) & "%."
  Next i
 
  '����������� ������
  ThisWorkbook.Sheets(SheetName).Activate
  Set list = ActiveSheet
'kolvo = InputBox("������� ����������� ���������� ������")

'If kolvo = "" Then Exit Sub
'If IsNumeric(kolvo) Then
'    kolvo = Fix(kolvo)
    For i = 1 To kolvo
        list.Copy after:=ActiveSheet
        ActiveSheet.Name = list.Name & i
        '������ ���
        Application.StatusBar = "����������� ������. " & _
        "���������: " & Int(100 * i / kolvo) & "%."
    Next
'Else
'    MsgBox "������������ �����"
'End If

'�������������� ������
On Error Resume Next
For Each Sht In Worksheets
    '������ ���
    Application.StatusBar = "�������������� ������."
    If Sht.Name = "91" Then Sht.Name = "9_21"
    If Sht.Name = "92" Then Sht.Name = "9_22"
    If Sht.Name = "93" Then Sht.Name = "9_23"
    If Sht.Name = "94" Then Sht.Name = "9_24"
    If Sht.Name = "95" Then Sht.Name = "9_1"
    If Sht.Name = "96" Then Sht.Name = "9_1_21"
    If Sht.Name = "97" Then Sht.Name = "9_1_22"
    If Sht.Name = "98" Then Sht.Name = "9_1_23"
    If Sht.Name = "99" Then Sht.Name = "9_1_24"
    If Sht.Name = "910" Then Sht.Name = "9_2"
    If Sht.Name = "911" Then Sht.Name = "9_2_21"
    If Sht.Name = "912" Then Sht.Name = "9_2_22"
    If Sht.Name = "913" Then Sht.Name = "9_2_23"
    If Sht.Name = "914" Then Sht.Name = "9_2_24"
Next

'����������� ��������
  On Error Resume Next
  
  For i = 1 To 4
    Sheets("9_2" & i).Activate
        [O2] = "202" & i
        '������ ���
        Application.StatusBar = "����������� �������� ������. ������ ��������. " & _
        "���������: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = 1 To 2
    Sheets("9_" & i).Activate
        [O1] = "���� " & i
        '������ ���
        Application.StatusBar = "����������� �������� ������. ������ ��������. " & _
        "���������: " & Int(100 * i / 2) & "%."
  Next i
  
  For i = 1 To 4
    Sheets("9_1_2" & i).Activate
        [O1] = "���� 1"
        [O2] = "202" & i
        '������ ���
        Application.StatusBar = "����������� �������� ������. ������ ��������. " & _
        "���������: " & Int(100 * i / 4) & "%."
  Next i
  
  For i = 1 To 4
    Sheets("9_2_2" & i).Activate
        [O1] = "���� 2"
        [O2] = "202" & i
        '������ ���
        Application.StatusBar = "����������� �������� ������. �������� ��������. " & _
        "���������: " & Int(100 * i / 4) & "%."
  Next i

ThisWorkbook.Sheets("Preferences").Activate
Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
    
End Sub
