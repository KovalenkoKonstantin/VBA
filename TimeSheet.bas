Attribute VB_Name = "TimeSheet"
Sub TimeSheet()
    ' Начало процедуры
    Start = Now() ' Запоминаем текущее время для измерения продолжительности выполнения

    Dim FilesToOpen ' Переменная для хранения пути к открываемым файлам
    Dim ThisWorkbook As Workbook, importWB As Workbook ' Переменные для рабочей книги
    Dim SheetName As String ' Переменная для хранения имени листа
    Dim ws As Worksheet ' Переменная для работы с листами

    Set ThisWorkbook = ActiveWorkbook ' Устанавливаем объект ThisWorkbook на текущую активную книгу
    On Error GoTo ExitHandler ' Обрабатываем ошибки, переводим поток выполнения в ExitHandler при возникновении ошибки
    SheetName = "Табель" ' Имя листа, с которым будем работать
    awLastCol = 63 ' Последний столбец для операций

    ' Отключаем обновление экрана и другие функции для ускорения выполнения
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False

    ' Делаем лист видимым и активируем его
    ThisWorkbook.Sheets(SheetName).Visible = True
    ThisWorkbook.Sheets(SheetName).Activate

    ' Открываем диалог выбора файлов
    FilesToOpen = Application.GetOpenFilename _
        (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
        MultiSelect:=True, Title:="Выберите файл с таблицей для редактирования")

    ' Если ничего не выбрано, выходим из процедуры
    If TypeName(FilesToOpen) = "Boolean" Then
        GoTo ExitHandler
    End If

    ThisWorkbook.Sheets(SheetName).Activate ' Активируем лист для обновления данных
    On Error Resume Next ' Игнорируем ошибки на следующих строках
    ActiveSheet.ShowAllData ' Показываем все данные, если имеются фильтры

    ' Импорт данных из первого выбранного файла
    Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

    On Error Resume Next ' Снова игнорируем ошибки

    importWB.Sheets(1).Activate ' Активируем первый лист импортируемой книги

    ' Удаляем предыдущие данные на листе
    ThisWorkbook.Sheets(SheetName).Activate
    awLastRow = Cells(Rows.Count, "AC").End(xlUp).row ' Находим номер последней заполненной строки в столбце AC
    Range(Cells(1, 1), Cells(awLastRow, awLastCol)).Select ' Выбираем диапазон для очистки
    With Selection
        .Clear ' Очищаем выбранный диапазон
    End With

    ' Импорт новых данных из всех листов импортированной книги
    For Each ws In importWB.Sheets ' Для каждого листа в импортируемой книге
        ws.Activate ' Активируем текущий лист
        iwLastRow = Cells(Rows.Count, "AC").End(xlUp).row ' Находим номер последней заполненной строки в столбце AC

        importWB.Activate
        Range(Cells(1, 1), Cells(iwLastRow, awLastCol)).Copy ' Копируем данные с текущего листа

        ThisWorkbook.Sheets(SheetName).Activate ' Возвращаемся к основному листу
        awFirstRow = Cells(Rows.Count, "AC").End(xlUp).row ' Находим номер последней заполненной строки на текущем листе
        awFirstCol = 1 ' Первая колонка, начиная с 1

        ' Вставляем скопированные данные в первый пустой ряд
        Range(Cells(awFirstRow + 1, awFirstCol), Cells(iwLastRow + awFirstRow, awLastCol)).Select
        With Selection
            .PasteSpecial Paste:=xlPasteAll ' Вставляем скопированные данные
        End With
    Next ws

    ' Закрываем файл, из которого импортировали данные
    importWB.Close
    ThisWorkbook.Sheets(SheetName).Activate ' Вновь активируем основной лист

    ' Показываем сообщение после завершения
    ' MsgBoxEx "Таблица успешно обновлена", 0, "Уведомление", 15

ExitHandler: ' Обработчик выхода из процедуры
    ' Включаем обратно параметры приложения, которые были отключены
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ' Этот код может скрыть лист, но закомментирован
    ' ThisWorkbook.Sheets(SheetName).Visible = False
    ThisWorkbook.Sheets("Preferences").Activate ' Активируем лист "Preferences"
    
    ' Рассчитываем и показываем время выполнения (закомментировано)
    ' Finish = (Now() - Start) * 24 * 60 * 60
    ' MsgBox (Finish)

    Exit Sub ' Завершаем процедуру
    
ErrHandler: ' Обработчик ошибок
    MsgBox Err.Description ' Если возникает ошибка, показываем сообщение с описанием
    Resume ExitHandler ' Возвращаемся в ExitHandler для очистки
End Sub
