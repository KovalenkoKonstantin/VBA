Attribute VB_Name = "PythonIntegration"

Sub Python()
    Application.StatusBar = "���������� ����� " & ActiveWorkbook.Name
    ActiveWorkbook.Save
    Application.StatusBar = "������� ������ � BackUp"
    If ActiveWorkbook.Name = "���_�����.xlsm" Then
        RunPython ("import �����; �����.FileSaving()")
    ElseIf ActiveWorkbook.Name = "���_45622C075_v.1.0.xlsm" Then
        RunPython ("import C075; C075.FileSaving()")
    ElseIf ActiveWorkbook.Name = "��� ����-23 ������_v1.6.xlsm" Then
        RunPython ("import ����_23; ����_23.FileSaving()")
    ElseIf ActiveWorkbook.Name = "��� 022-7 1 ����_v1.7.xlsm" Then
        RunPython ("import ����������2207; ����������2207.FileSaving()")
    End If
'    Application.StatusBar = "�������� PDF"
'    SaveToPDF
    Application.StatusBar = False
    
End Sub
'���_v.1.0 - �������� �������
'FileSaving() - �������� ������� � �������
