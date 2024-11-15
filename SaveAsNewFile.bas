Attribute VB_Name = "SaveAsNewFile"
Sub SaveFinalTableAsNewFile_v_1_0()
    Dim wsSource As Worksheet
    Dim tblFinal As ListObject
    Dim newFilePath As String
    Dim currentDate As String
    Dim currentTime As String
    Dim newWorkbook As Workbook
    Dim newWorksheet As Worksheet
    
    ' Устанавливаем ссылку на лист "Для загрузки"
    Set wsSource = ThisWorkbook.Sheets("Для загрузки")
    
    ' Находим таблицу "Final" на этом листе
    Set tblFinal = wsSource.ListObjects("Final")
    
    ' Получаем текущую дату и время в нужном формате
    currentDate = Format(Date, "DD.MM.YYYY")
    currentTime = Format(Time, "HH-MM")
    
    ' Формируем имя нового файла
    newFilePath = ThisWorkbook.Path & "\Для загрузки ДМС (" & currentDate & " " & currentTime & ").xlsx"
    
    ' Создаём новый Workbook
    Set newWorkbook = Workbooks.Add
    ' Добавляем новый лист в новый файл
    Set newWorksheet = newWorkbook.Sheets(1)
    
    ' Копируем данные из таблицы "Final" в новый лист
    tblFinal.Range.Copy Destination:=newWorksheet.Cells(1, 1)
    
    ' Убираем форматирование ссылок (если оно есть), только значения
    newWorksheet.Cells.Copy
    newWorksheet.Cells.PasteSpecial Paste:=xlPasteValues
    
    ' Сохраняем новый файл в ту же директорию
    newWorkbook.SaveAs Filename:=newFilePath, FileFormat:=xlOpenXMLWorkbook
    
    ' Закрываем новый файл
    newWorkbook.Close SaveChanges:=False
    
'    ' Выводим сообщение о сохранении
'    MsgBox "Файл сохранён как: " & newFilePath, vbInformation
End Sub


Sub SaveFinalTableAsNewFile_v_1_1()
    Dim wsSource As Worksheet
    Dim tblSelected As ListObject
    Dim newFilePath As String
    Dim currentDate As String
    Dim currentTime As String
    Dim newWorkbook As Workbook
    Dim newWorksheet As Worksheet
    Dim userChoice As String
    Dim folderPath As String
    
    ' Путь к папке Загрузка
    folderPath = ThisWorkbook.Path & "\Загрузка\"
    
    ' Проверяем, существует ли папка "Загрузка"
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Папка 'Загрузка' не найдена в текущей директории.", vbExclamation
        Exit Sub
    End If
    
    ' Показать диалоговое окно с двумя кнопками: ДМС и НС
    userChoice = MsgBox("Выберите тип: ДМС или НС." & vbCrLf & _
                        "Нажмите 'Да' для ДМС, 'Нет' для НС", vbYesNo + vbQuestion, "Выбор типа")
    
    ' Проверка выбора пользователя
    If userChoice = vbYes Then
        ' Выбран ДМС
        userChoice = "ДМС"
    ElseIf userChoice = vbNo Then
        ' Выбран НС
        userChoice = "НС"
    Else
        MsgBox "Операция отменена.", vbInformation
        Exit Sub
    End If
    
    ' Устанавливаем ссылку на лист "Для загрузки"
    Set wsSource = ThisWorkbook.Sheets("Для загрузки")
    
    ' Определяем, какую таблицу выбирать в зависимости от выбора пользователя
    If userChoice = "ДМС" Then
        ' Таблица "Final_ДМС" для ДМС
        On Error Resume Next
        Set tblSelected = wsSource.ListObjects("Final_ДМС")
        On Error GoTo 0
        If tblSelected Is Nothing Then
            MsgBox "Таблица 'Final_ДМС' не найдена!", vbExclamation
            Exit Sub
        End If
    ElseIf userChoice = "НС" Then
        ' Таблица "Final_НС" для НС
        On Error Resume Next
        Set tblSelected = wsSource.ListObjects("Final_НС")
        On Error GoTo 0
        If tblSelected Is Nothing Then
            MsgBox "Таблица 'Final_НС' не найдена!", vbExclamation
            Exit Sub
        End If
    End If
    
    ' Получаем текущую дату и время в нужном формате
    currentDate = Format(Date, "DD.MM.YYYY")
    currentTime = Format(Time, "HH-MM")
    
    ' Формируем имя нового файла в зависимости от выбора
    newFilePath = folderPath & "Для загрузки " & userChoice & " (" & currentDate & " " & currentTime & ").xlsx"
    
    ' Проверка на существование файла и его удаление, если файл уже существует
    If Dir(newFilePath) <> "" Then
        ' Файл существует, удаляем его
        Kill newFilePath
    End If
    
    ' Создаём новый Workbook
    Set newWorkbook = Workbooks.Add
    ' Добавляем новый лист в новый файл
    Set newWorksheet = newWorkbook.Sheets(1)
    
    ' Копируем данные из выбранной таблицы в новый лист
    tblSelected.Range.Copy Destination:=newWorksheet.Cells(1, 1)
    
    ' Убираем форматирование ссылок (если оно есть), только значения
    newWorksheet.Cells.Copy
    newWorksheet.Cells.PasteSpecial Paste:=xlPasteValues
    
    ' Сохраняем новый файл в ту же директорию (папка Загрузка)
    newWorkbook.SaveAs Filename:=newFilePath, FileFormat:=xlOpenXMLWorkbook
    
    ' Закрываем новый файл
    newWorkbook.Close SaveChanges:=False
    
    ' Выводим сообщение о сохранении
'    MsgBox "Файл сохранён как: " & newFilePath, vbInformation
End Sub

Sub SaveFinalTableAsNewFile()
    Dim wsSource As Worksheet
    Dim tblSelected As ListObject
    Dim newFilePath As String
    Dim currentDate As String
    Dim currentTime As String
    Dim newWorkbook As Workbook
    Dim newWorksheet As Worksheet
    Dim folderPath As String
    
    ' Показываем форму для выбора типа
    Опции.Show
'    Debug.Print (Опции.userChoice)
    ' Если форма была закрыта, пользователь не сделал выбора
    If Опции.userChoice = "" Then
        MsgBox "Операция отменена.", vbInformation
        Exit Sub
    End If
    
    ' Путь к папке "Загрузка"
    folderPath = ThisWorkbook.Path & "\Загрузка\"
    
    ' Проверяем, существует ли папка "Загрузка"
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Папка 'Загрузка' не найдена в текущей директории.", vbExclamation
        Exit Sub
    End If
    
    ' Устанавливаем ссылку на лист "Для загрузки"
    Set wsSource = ThisWorkbook.Sheets("Для загрузки")
    
    ' Определяем, какую таблицу выбирать в зависимости от выбора пользователя
    If Опции.userChoice = "ДМС" Then
        ' Таблица "Final_ДМС" для ДМС
        On Error Resume Next
        Set tblSelected = wsSource.ListObjects("Final_ДМС")
        On Error GoTo 0
        If tblSelected Is Nothing Then
            MsgBox "Таблица 'Final_ДМС' не найдена!", vbExclamation
            Exit Sub
        End If
    ElseIf Опции.userChoice = "НС" Then
        ' Таблица "Final_НС" для НС
        On Error Resume Next
        Set tblSelected = wsSource.ListObjects("Final_НС")
        On Error GoTo 0
        If tblSelected Is Nothing Then
            MsgBox "Таблица 'Final_НС' не найдена!", vbExclamation
            Exit Sub
        End If
    End If
    
    ' Получаем текущую дату и время в нужном формате
    currentDate = Format(Date, "DD.MM.YYYY")
    currentTime = Format(Time, "HH-MM")
    
    ' Формируем имя нового файла в зависимости от выбора
    newFilePath = folderPath & "Для загрузки " & Опции.userChoice & " (" & currentDate & " " & currentTime & ").xlsx"
    
    ' Проверка на существование файла и его удаление, если файл уже существует
    If Dir(newFilePath) <> "" Then
        ' Файл существует, удаляем его
        Kill newFilePath
    End If
    
    ' Создаём новый Workbook
    Set newWorkbook = Workbooks.Add
    ' Добавляем новый лист в новый файл
    Set newWorksheet = newWorkbook.Sheets(1)
    
    ' Копируем данные из выбранной таблицы в новый лист
    tblSelected.Range.Copy Destination:=newWorksheet.Cells(1, 1)
    
    ' Убираем форматирование ссылок (если оно есть), только значения
    newWorksheet.Cells.Copy
    newWorksheet.Cells.PasteSpecial Paste:=xlPasteValues
    
    ' Сохраняем новый файл в ту же директорию (папка Загрузка)
    newWorkbook.SaveAs Filename:=newFilePath, FileFormat:=xlOpenXMLWorkbook
    
    ' Закрываем новый файл
    newWorkbook.Close SaveChanges:=False
    
    ' Выводим сообщение о сохранении
    MsgBox "Файл сохранён как: " & newFilePath, vbInformation
End Sub

