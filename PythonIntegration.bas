Attribute VB_Name = "PythonIntegration"

Sub Python()
Dim ThisWorkbook, Command As String
Filename = ActiveWorkbook.Name

'Python �� ����� � \
src = Replace(ActiveWorkbook.Path, "\", "/") + "/"

'Debug.Print Filename
'Debug.Print src

    Application.StatusBar = "���������� ����� " & ActiveWorkbook.Name
    ActiveWorkbook.Save
    Application.StatusBar = "������� ������ � BackUp"
    
'������� ��� xlwings
Command = "import Save; Save.FileSaving('" + Filename + "', '" + src + "')"

'Debug.Print Command

'��� ������� � xlwings
RunPython (Command)

Application.StatusBar = False
    
End Sub
