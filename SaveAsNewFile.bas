Attribute VB_Name = "SaveAsNewFile"
Sub SaveFinalTableAsNewFile()
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

