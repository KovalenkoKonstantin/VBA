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

Sub Python�����()
    Application.StatusBar = "���������� ���� �����"
    ActiveWorkbook.Save
    Application.StatusBar = "������� ������ � BackUp"
    RunPython ("import �����; �����.FileSaving()")
'    Application.StatusBar = "�������� PDF"
'    SaveToPDF
    Application.StatusBar = False
End Sub
