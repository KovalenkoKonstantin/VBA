Attribute VB_Name = "PythonIntegration"

Sub PythonSaving()
    Application.StatusBar = "���������� �������� �����"
    ActiveWorkbook.Save
    Application.StatusBar = "������� ������ � BackUp"
    RunPython ("import prog; prog.FileSaving()")
    Application.StatusBar = "�������� PDF"
    SaveToPDF
    Application.StatusBar = False
End Sub
'���_v.1.0 - �������� �������
'FileSaving() - �������� ������� � �������

Sub Python()
    Application.StatusBar = "���������� ����� " & ActiveWorkbook.Name
    ActiveWorkbook.Save
    Application.StatusBar = "������� ������ � BackUp"
    If ActiveWorkbook.Name = "���_�����.xlsm" Then
        RunPython ("import �����; �����.FileSaving()")
    ElseIf ActiveWorkbook.Name = "���_45622?075_v.1.0.xlsm" Then
        RunPython ("import �����; �����.FileSaving()")
    End If
'    Application.StatusBar = "�������� PDF"
'    SaveToPDF
    Application.StatusBar = False
End Sub
