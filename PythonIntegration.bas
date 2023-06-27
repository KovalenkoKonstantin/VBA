Attribute VB_Name = "PythonIntegration"

Sub PythonSavingУлей23()
    Application.StatusBar = "Сохранение активной книги"
    ActiveWorkbook.Save
    Application.StatusBar = "Перенос данных в BackUp"
    RunPython ("import Улей_23; Улей_23.FileSaving()")
'    Application.StatusBar = "Создание PDF"
'    SaveToPDF
    Application.StatusBar = False
End Sub
'РКМ_v.1.0 - название скрипта
'FileSaving() - название функции в скрипте
