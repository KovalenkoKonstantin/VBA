Attribute VB_Name = "PythonIntegration"

Sub Python()
    Application.StatusBar = "Сохранение книги " & ActiveWorkbook.Name
    ActiveWorkbook.Save
    Application.StatusBar = "Перенос данных в BackUp"
    If ActiveWorkbook.Name = "РКМ_Поиск.xlsm" Then
        RunPython ("import Поиск; Поиск.FileSaving()")
    ElseIf ActiveWorkbook.Name = "РКМ_45622C075_v.1.0.xlsm" Then
        RunPython ("import C075; C075.FileSaving()")
    ElseIf ActiveWorkbook.Name = "ОРЦ Улей-23 работа_v1.6.xlsm" Then
        RunPython ("import Улей_23; Улей_23.FileSaving()")
    ElseIf ActiveWorkbook.Name = "ТФЦ 022-7 1 этап_v1.7.xlsm" Then
        RunPython ("import Профитроль2207; Профитроль2207.FileSaving()")
    End If
'    Application.StatusBar = "Создание PDF"
'    SaveToPDF
    Application.StatusBar = False
    
End Sub
'РКМ_v.1.0 - название скрипта
'FileSaving() - название функции в скрипте
