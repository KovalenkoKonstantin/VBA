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

Sub PythonПоиск()
    Application.StatusBar = "Сохранение этой книги"
    ActiveWorkbook.Save
    Application.StatusBar = "Перенос данных в BackUp"
    RunPython ("import Поиск; Поиск.FileSaving()")
'    Application.StatusBar = "Создание PDF"
'    SaveToPDF
    Application.StatusBar = False
End Sub
