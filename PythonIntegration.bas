Attribute VB_Name = "PythonIntegration"

Sub Python()
Dim ThisWorkbook, Command As String
Filename = ActiveWorkbook.Name

'Python не умеет в \
src = Replace(ActiveWorkbook.Path, "\", "/") + "/"

'Debug.Print Filename
'Debug.Print src

    Application.StatusBar = "Сохранение книги " & ActiveWorkbook.Name
    ActiveWorkbook.Save
    Application.StatusBar = "Перенос данных в BackUp"
    
'команда для xlwings
Command = "import Save; Save.FileSaving('" + Filename + "', '" + src + "')"

'Debug.Print Command

'имя функции в xlwings
RunPython (Command)

Application.StatusBar = False
    
End Sub
