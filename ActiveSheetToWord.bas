Attribute VB_Name = "ActiveSheetToWord"
Sub ExportActiveSheetToWord()
    Dim ws As Worksheet
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim printArea As Range
    Dim filePath As String
    Dim fileName As String

    ' Устанавливаем ссылку на активный лист
    Set ws = ThisWorkbook.ActiveSheet

    ' Определяем область печати
    On Error Resume Next
    Set printArea = ws.Range(ws.PageSetup.printArea)
    If printArea Is Nothing Then
        Set printArea = ws.UsedRange ' если область печати не задана, используем все занятые ячейки
    End If
    On Error GoTo 0
    
    ' Создаем приложение Word
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True ' Делает Word видимым

    ' Создаем новый документ
    Set wordDoc = wordApp.Documents.Add

    ' Копируем область печати из Excel
    printArea.Copy

    ' Вставляем в документ Word
    wordDoc.Content.Paste

    ' Определяем путь сохранения
    filePath = ThisWorkbook.Path
    fileName = ws.Name & ".docx" ' Имя файла по названию листа

    ' Сохраняем документ
    wordDoc.SaveAs filePath & "\" & fileName
    wordDoc.Close
    wordApp.Quit

    ' Освобождаем память
    Set wordDoc = Nothing
    Set wordApp = Nothing
    
    ' Уведомление пользователя
'    MsgBox "Документ сохранен: " & filePath & "\" & fileName
End Sub
