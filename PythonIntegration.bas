Attribute VB_Name = "PythonIntegration"

Sub PythonSaving����23()
    Application.StatusBar = "���������� �������� �����"
    ActiveWorkbook.Save
    Application.StatusBar = "������� ������ � BackUp"
    RunPython ("import ����_23; ����_23.FileSaving()")
'    Application.StatusBar = "�������� PDF"
'    SaveToPDF
    Application.StatusBar = False
End Sub
'���_v.1.0 - �������� �������
'FileSaving() - �������� ������� � �������
