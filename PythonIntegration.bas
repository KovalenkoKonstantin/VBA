Attribute VB_Name = "PythonIntegration"

Sub Python()
    Application.StatusBar = "���������� ����� " & ActiveWorkbook.Name
    ActiveWorkbook.Save
    Application.StatusBar = "������� ������ � BackUp"
    If ActiveWorkbook.Name = "���_�����_v.1.0.xlsm" Then
        RunPython ("import �����; �����.FileSaving()")
    ElseIf ActiveWorkbook.Name = "���_45622C075_v.1.0.xlsm" Then
        RunPython ("import C075; C075.FileSaving()")
    ElseIf ActiveWorkbook.Name = "��� ����-23 ������_v1.7.xlsm" Then
        RunPython ("import ����_23; ����_23.FileSaving()")
    ElseIf ActiveWorkbook.Name = "��� 022-7 1 ����_v1.8.xlsm" Then
        RunPython ("import ����������2207; ����������2207.FileSaving()")
    ElseIf ActiveWorkbook.Name = "���_����-�����-��_v.1.1.xlsm" Then
        RunPython ("import ����_�����_��; ����_�����_��.FileSaving()")
    ElseIf ActiveWorkbook.Name = "���_���-���-23_v.1.0.xlsm" Then
        RunPython ("import ���_���_23; ���_���_23.FileSaving()")
    ElseIf ActiveWorkbook.Name = "���_HW 2000_v.1.0.xlsm" Then
        RunPython ("import HW_2000; HW_2000.FileSaving()")
    ElseIf ActiveWorkbook.Name = "���_HW 2000 & HW 100_v.1.0.xlsm" Then
        RunPython ("import HW_2000_HW_100; HW_2000_HW_100.FileSaving()")
    ElseIf ActiveWorkbook.Name = "��� ����-23_v1.0.xlsm" Then
        RunPython ("import ���_����23; ���_����23.FileSaving()")
    ElseIf ActiveWorkbook.Name = "���_���-���-24_v.1.1.xlsm" Then
        RunPython ("import ���_���_24; ���_���_24.FileSaving()")
    ElseIf ActiveWorkbook.Name = "���_HW50_v.1.0.xlsm" Then
        RunPython ("import HW50; HW50.FileSaving()")
    ElseIf ActiveWorkbook.Name = "���_HW100_C+_unlim_v.1.0.xlsm" Then
        RunPython ("import HW_100_C_unlim; HW_100_C_unlim.FileSaving()")
    ElseIf ActiveWorkbook.Name = "���_HW100_C+_wifi+_unlim_v.1.0.xlsm" Then
        RunPython ("import HW_100_C_wifi_unlim; HW_100_C_wifi_unlim.FileSaving()")
    ElseIf ActiveWorkbook.Name = "���_HW1000_v.1.0.xlsm" Then
        RunPython ("import HW_1000; HW_1000.FileSaving()")
    ElseIf ActiveWorkbook.Name = "���_HW1000_C_v.1.0.xlsm" Then
        RunPython ("import HW_1000_C; HW_1000_C.FileSaving()")
    ElseIf ActiveWorkbook.Name = "���_HW1000_D_v.1.0.xlsm" Then
        RunPython ("import HW_1000_D; HW_1000_D.FileSaving()")
    ElseIf ActiveWorkbook.Name = "���_HW 2000_add_software_v.1.0.xlsm" Then
        RunPython ("import HW_2000_add_software; HW_2000_add_software.FileSaving()")
    ElseIf ActiveWorkbook.Name = "���_HW 2000_v.1.1.xlsm" Then
        RunPython ("import HW_2000_; HW_2000_.FileSaving()")
    ElseIf ActiveWorkbook.Name = "���_�����_v.1.0.xlsm" Then
        RunPython ("import �����; �����.FileSaving()")
    ElseIf ActiveWorkbook.Name Like "������_v.1.*" Then
        RunPython ("import Sample; Sample.FileSaving()")
    End If
'    Application.StatusBar = "�������� PDF"
'    SaveToPDF
    Application.StatusBar = False
    
End Sub
'���_v.1.0 - �������� �������
'FileSaving() - �������� ������� � �������

