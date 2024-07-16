Attribute VB_Name = "ToPDF"
Sub SaveToPDF()
    Dim Start As Double
    Dim Finish As Double
    Dim SaveName As String
    Dim Path As String
    Dim ThisWorkbook As Workbook
    
    ' Запоминаем время начала выполнения процедуры
    Start = Now()
    
    ' Устанавливаем ссылку на активную книгу
    Set ThisWorkbook = ActiveWorkbook
    
    ' Обработка ошибок: переход к ExitHandler в случае ошибки
    On Error GoTo ExitHandler
    
    ' Активируем лист "Preferences" и получаем имя для сохранения из ячейки H30
    ThisWorkbook.Sheets("Preferences").Activate
    SaveName = ActiveSheet.Range("H30").Text
    
    ' Получаем путь к папке, где сохранена активная книга
    Path = ThisWorkbook.Path
    
    ' Отключаем обновление экрана, события, разрывы страниц, статусную строку и предупреждения
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
    
    ' Вызываем процедуру для выбора листов с бледно-желтым цветом вкладки
    SelectPaleYellowSheets
    
    ' Экспортируем активный лист в PDF
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=Path & "\" & SaveName & ".pdf", _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
    ' Возвращаемся на лист "Preferences"
    Sheets("Preferences").Select
    
    ' Вычисляем время выполнения процедуры в секундах
    Finish = (Now() - Start) * 24 * 60 * 60
    
ExitHandler:
    ' Проверяем, было ли выполнение процедуры слишком быстрым (менее 0.1 секунды)
    If Finish < 0.1 Then
        MsgBox "Неправильные диапазоны. Файл открыт."
    Else
'        MsgBox "Файл сохранен в формате PDF в корневой папке", vbInformation, "Done"
    End If
    
    ' Восстанавливаем настройки Excel
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    
    ' Возвращаемся на лист "Preferences"
    ThisWorkbook.Sheets("Preferences").Activate
End Sub

Sub SelectPaleYellowSheets()
    Dim ws As Worksheet
    Dim paleYellowSheets As Collection
    Dim i As Integer
    Dim paleYellowColor As Long
    
    ' Устанавливаем цвет бледно-желтого
    paleYellowColor = 13434879
    
    ' Создаем коллекцию для хранения имен листов, которые нужно выделить
    Set paleYellowSheets = New Collection
    
    ' Проходим по всем листам в книге
    For Each ws In ThisWorkbook.Sheets
        ' Проверяем, залит ли лист бледно-желтым цветом
        If ws.Tab.Color = paleYellowColor Then
            paleYellowSheets.Add ws.Name, ws.Name
        End If
    Next ws
    
' Проверяем, есть ли хотя бы один лист в коллекции
    If paleYellowSheets.Count > 0 Then
        ' Активируем первый лист из коллекции
        ThisWorkbook.Sheets(paleYellowSheets(1)).Activate
        
        ' Выделяем только те листы, которые находятся в коллекции
        For i = 1 To paleYellowSheets.Count
            ThisWorkbook.Sheets(paleYellowSheets(i)).Select Replace:=False
        Next i
    Else
        MsgBox "Нет листов с бледно-желтым цветом вкладки."
    End If
End Sub

