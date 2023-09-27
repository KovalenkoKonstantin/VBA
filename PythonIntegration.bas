Attribute VB_Name = "PythonIntegration"

Sub Python()
    Application.StatusBar = "Сохранение книги " & ActiveWorkbook.Name
    ActiveWorkbook.Save
    Application.StatusBar = "Перенос данных в BackUp"
    If ActiveWorkbook.Name = "РКМ_Поиск_v.1.0.xlsm" Then
        RunPython ("import Поиск; Поиск.FileSaving()")
    ElseIf ActiveWorkbook.Name = "РКМ_45622C075_v.1.0.xlsm" Then
        RunPython ("import C075; C075.FileSaving()")
    ElseIf ActiveWorkbook.Name = "ОРЦ Улей-23 работа_v1.7.xlsm" Then
        RunPython ("import Улей_23; Улей_23.FileSaving()")
    ElseIf ActiveWorkbook.Name = "ТФЦ 022-7 1 этап_v1.8.xlsm" Then
        RunPython ("import Профитроль2207; Профитроль2207.FileSaving()")
    ElseIf ActiveWorkbook.Name = "РКМ_Улей-Режим-ПЗ_v.1.1.xlsm" Then
        RunPython ("import Улей_Режим_ПЗ; Улей_Режим_ПЗ.FileSaving()")
    ElseIf ActiveWorkbook.Name = "РКМ_ОБД-СНГ-23_v.1.0.xlsm" Then
        RunPython ("import ОБД_СНГ_23; ОБД_СНГ_23.FileSaving()")
    ElseIf ActiveWorkbook.Name = "РКМ_HW 2000_v.1.0.xlsm" Then
        RunPython ("import HW_2000; HW_2000.FileSaving()")
    ElseIf ActiveWorkbook.Name = "РКМ_HW 2000 & HW 100_v.1.0.xlsm" Then
        RunPython ("import HW_2000_HW_100; HW_2000_HW_100.FileSaving()")
    ElseIf ActiveWorkbook.Name = "ТФЦ Улей-23_v1.0.xlsm" Then
        RunPython ("import ТФЦ_Улей23; ТФЦ_Улей23.FileSaving()")
    ElseIf ActiveWorkbook.Name = "РКМ_ОБД-СНГ-24_v.1.1.xlsm" Then
        RunPython ("import ОБД_СНГ_24; ОБД_СНГ_24.FileSaving()")
    ElseIf ActiveWorkbook.Name = "РКМ_HW50_v.1.0.xlsm" Then
        RunPython ("import HW50; HW50.FileSaving()")
    ElseIf ActiveWorkbook.Name = "РКМ_HW100_C+_unlim_v.1.0.xlsm" Then
        RunPython ("import HW_100_C_unlim; HW_100_C_unlim.FileSaving()")
    ElseIf ActiveWorkbook.Name = "РКМ_HW100_C+_wifi+_unlim_v.1.0.xlsm" Then
        RunPython ("import HW_100_C_wifi_unlim; HW_100_C_wifi_unlim.FileSaving()")
    ElseIf ActiveWorkbook.Name = "РКМ_HW1000_v.1.0.xlsm" Then
        RunPython ("import HW_1000; HW_1000.FileSaving()")
    ElseIf ActiveWorkbook.Name = "РКМ_HW1000_C_v.1.0.xlsm" Then
        RunPython ("import HW_1000_C; HW_1000_C.FileSaving()")
    ElseIf ActiveWorkbook.Name = "РКМ_HW1000_D_v.1.0.xlsm" Then
        RunPython ("import HW_1000_D; HW_1000_D.FileSaving()")
    ElseIf ActiveWorkbook.Name = "РКМ_HW 2000_add_software_v.1.0.xlsm" Then
        RunPython ("import HW_2000_add_software; HW_2000_add_software.FileSaving()")
    ElseIf ActiveWorkbook.Name = "РКМ_HW 2000_v.1.1.xlsm" Then
        RunPython ("import HW_2000_; HW_2000_.FileSaving()")
    ElseIf ActiveWorkbook.Name = "РКМ_ТСИСЗ_v.1.0.xlsm" Then
        RunPython ("import ТСИСЗ; ТСИСЗ.FileSaving()")
    ElseIf ActiveWorkbook.Name Like "Шаблон_v.1.*" Then
        RunPython ("import Sample; Sample.FileSaving()")
    End If
'    Application.StatusBar = "Создание PDF"
'    SaveToPDF
    Application.StatusBar = False
    
End Sub
'РКМ_v.1.0 - название скрипта
'FileSaving() - название функции в скрипте

