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
    ElseIf ActiveWorkbook.Name = "РКМ_HW100_C+wifi_+_unlim_v.1.0.xlsm" Then
        RunPython ("import HW_100_C_wifi_unlim; HW_100_C_wifi_unlim.FileSaving()")
    End If
'    Application.StatusBar = "Создание PDF"
'    SaveToPDF
    Application.StatusBar = False
    
End Sub
'РКМ_v.1.0 - название скрипта
'FileSaving() - название функции в скрипте

