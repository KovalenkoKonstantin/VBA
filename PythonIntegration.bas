Attribute VB_Name = "PythonIntegration"

Sub PythonSaving()
    Application.StatusBar = "Сохранение активной книги"
    ActiveWorkbook.Save
    Application.StatusBar = "Перенос данных в BackUp"
    RunPython ("import prog; prog.FileSaving()")
    Application.StatusBar = "Создание PDF"
    SaveToPDF
    Application.StatusBar = False
End Sub
'РКМ_v.1.0 - название скрипта
'FileSaving() - название функции в скрипте

Sub Python()
    Application.StatusBar = "Сохранение книги " & ActiveWorkbook.Name
    ActiveWorkbook.Save
    Application.StatusBar = "Перенос данных в BackUp"
    If ActiveWorkbook.Name = "РКМ_Поиск.xlsm" Then
        RunPython ("import Поиск; Поиск.FileSaving()")
    ElseIf ActiveWorkbook.Name = "РКМ_45622?075_v.1.0.xlsm" Then
        RunPython ("import Поиск; Поиск.FileSaving()")
    End If
'    Application.StatusBar = "Создание PDF"
'    SaveToPDF
    Application.StatusBar = False
End Sub
