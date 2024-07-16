Attribute VB_Name = "Layout"
Sub LayoutOn()
    Dim ThisWorkbook As Workbook
    Dim ws As Worksheet
    Dim CompanyName As String
    Dim aws As Integer
    Dim tottal As Integer

    ' Скрываем системные листы (предполагается, что у вас есть процедура HideSys)
    HideSys

    ' Устанавливаем ссылку на активную книгу
    Set ThisWorkbook = ActiveWorkbook

    ' Получаем имя компании из ячейки C7 на листе "Preferences"
    CompanyName = ThisWorkbook.Sheets("Preferences").Range("C7").Value2

    ' Обработка ошибок: переход к ExitHandler в случае ошибки
    On Error GoTo ExitHandler

    ' Отключаем обновление экрана, события, разрывы страниц и предупреждения
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayAlerts = False

    ' Подсчитываем количество видимых листов
    tottal = 0
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            tottal = tottal + 1
        End If
    Next ws

    ' Проходим по всем листам в книге
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Activate
            aws = Application.ActiveSheet.Index

            ' Обновляем статусную строку
            Application.StatusBar = "Обрабатывается " & aws & " лист из " & tottal & " листов. Выполнено: " & _
                                    Int(aws / tottal * 100) & " %. Расчётное время до конца выполнения программы: " & _
                                    Int((tottal - aws) * 3) & " секунд(ы)."

            ' Настраиваем колонтитулы
            With ActiveSheet.PageSetup
                .CenterHeader = "&""Times New Roman,обычный""&KFF0000Данный документ не согласован."
                .RightFooter = "&""Times New Roman,обычный""&KFF0000Настоящий документ и любые приложения к нему содержат информацию, относящуюся к коммерческой тайне " & CompanyName
            End With
        End If
    Next ws

ExitHandler:
    ' Восстанавливаем настройки Excel
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True

    ' Возвращаемся на лист "Preferences"
    ThisWorkbook.Sheets("Preferences").Activate
    ActiveWindow.View = xlNormalView

    Exit Sub

ErrHandler:
    ' Выводим сообщение об ошибке и возвращаемся к ExitHandler
    MsgBox Err.Description
    Resume ExitHandler
End Sub
Sub LayoutInfotecs()
    Dim ThisWorkbook As Workbook
    Dim ws As Worksheet
    Dim CompanyName As String
    Dim aws As Integer
    Dim tottal As Integer
    
    ' Скрываем системные листы (предполагается, что у вас есть процедура HideSys)
    HideSys
    
    ' Устанавливаем ссылку на активную книгу
    Set ThisWorkbook = ActiveWorkbook
    
    ' Получаем имя компании из ячейки C7 на листе "Preferences"
    CompanyName = ThisWorkbook.Sheets("Preferences").Range("C7").Value2
    
    ' Обработка ошибок: переход к ExitHandler в случае ошибки
    On Error GoTo ExitHandler
    
    ' Отключаем обновление экрана, события, разрывы страниц и предупреждения
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayAlerts = False
    
    ' Подсчитываем количество видимых листов
    tottal = 0
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            tottal = tottal + 1
        End If
    Next ws
    
    ' Проходим по всем листам в книге
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Activate
            aws = Application.ActiveSheet.Index
            
            ' Обновляем статусную строку
            Application.StatusBar = "Обрабатывается " & aws & " лист из " & tottal & " листов. Выполнено: " & _
                                    Int(aws / tottal * 100) & " %. Расчётное время до конца выполнения программы: " & _
                                    Int((tottal - aws) * 3) & " секунд(ы)."
            
            ' Настраиваем нижний колонтитул
            With ActiveSheet.PageSetup
                .RightFooter = "&""Times New Roman,обычный""&KFF0000Экземпляр " & CompanyName
            End With
        End If
    Next ws
    
ExitHandler:
    ' Восстанавливаем настройки Excel
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    
    ' Возвращаемся на лист "Preferences"
    ThisWorkbook.Sheets("Preferences").Activate
    ActiveWindow.View = xlNormalView
    
    Exit Sub
    
ErrHandler:
    ' Выводим сообщение об ошибке и возвращаемся к ExitHandler
    MsgBox Err.Description
    Resume ExitHandler
End Sub

Sub LayoutOff()
    Dim ThisWorkbook As Workbook
    Dim ws As Worksheet
    Dim aws As Integer
    Dim tottal As Integer

    ' Скрываем системные листы (предполагается, что у вас есть процедура HideSys)
    HideSys

    ' Устанавливаем ссылку на активную книгу
    Set ThisWorkbook = ActiveWorkbook

    ' Обработка ошибок: переход к ExitHandler в случае ошибки
    On Error GoTo ExitHandler

    ' Отключаем обновление экрана, события, разрывы страниц и предупреждения
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayAlerts = False

    ' Подсчитываем количество видимых листов
    tottal = 0
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            tottal = tottal + 1
        End If
    Next ws

    ' Проходим по всем листам в книге
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Activate
            aws = Application.ActiveSheet.Index

            ' Обновляем статусную строку
            Application.StatusBar = "Обрабатывается " & aws & " лист из " & tottal & " листов. Выполнено: " & _
                                    Int(aws / tottal * 100) & " %. Расчётное время до конца выполнения программы: " & _
                                    Int((tottal - aws) * 3) & " секунд."

            ' Очищаем настройки колонтитулов
            ActiveSheet.PageSetup.CenterHeaderPicture.Filename = ""
            With ActiveSheet.PageSetup
                .CenterHeader = ""
                .RightHeader = ""
                .RightFooter = ""
            End With
        End If
    Next ws

ExitHandler:
    ' Восстанавливаем настройки Excel
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True

    ' Возвращаемся на лист "Preferences"
    ThisWorkbook.Sheets("Preferences").Activate
    ActiveWindow.View = xlNormalView

    Exit Sub

ErrHandler:
    ' Выводим сообщение об ошибке и возвращаемся к ExitHandler
    MsgBox Err.Description
    Resume ExitHandler
End Sub


