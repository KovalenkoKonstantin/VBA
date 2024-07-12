Attribute VB_Name = "PythonIntegration"
Sub Python()
    ' Объявляем переменные для хранения имени книги и команды
    Dim ThisWorkbook, Command As String
    ' Получаем имя активной книги
    Filename = ActiveWorkbook.Name

    ' Заменяем обратные слеши на прямые в пути к активной книге
    ' Python не умеет работать с обратными слешами
    src = Replace(ActiveWorkbook.Path, "\", "/") + "/"

    ' Выводим отладочную информацию
    ' Debug.Print Filename
    ' Debug.Print src

    ' Устанавливаем статус-бар для отображения процесса сохранения книги
    Application.StatusBar = "Сохранение книги " & ActiveWorkbook.Name
    ' Сохраняем активную книгу
    ActiveWorkbook.Save
    ' Обновляем статус-бар для отображения процесса переноса данных
    Application.StatusBar = "Перенос данных в BackUp"
    
    ' Формируем команду для выполнения через xlwings
    Command = "import Save; Save.FileSaving('" + Filename + "', '" + src + "')"

    ' Выводим отладочную информацию
    ' Debug.Print Command

    ' Выполняем команду Python через xlwings
    ' Имя функции в xlwings
    RunPython (Command)

    ' Сбрасываем статус-бар
    Application.StatusBar = False
End Sub
