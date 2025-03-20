Attribute VB_Name = "Support"
Sub ShowTabs()
 Dim tb
 For Each tb In Worksheets
 tb.Visible = True
 Next
 ActiveWorkbook.Sheets("Preferences").Activate
End Sub

Sub HideSys()
Application.Calculation = xlManual
 Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Range("A1").Value2 = "sys" _
        Or ws.Range("A1").Value2 = "Трудоёмкость" _
        Or ws.Range("A1").Value2 = "Статья затрат" Or ws.Range("A1").Value2 = "Имя" _
        Or ws.Range("A1").Value2 = "company_name" _
        Or ws.Range("A1").Value2 = "Наименование статей в 1С" _
        Or ws.Range("H2").Value2 = "Отчет о финансовых результатах" _
        Or ws.Range("A1").Value2 = "organization_id" _
        Or ws.Range("J1").Value2 = "Сумма" Then
            ws.Visible = False
        End If
    Next ws
Application.Calculation = xlAutomatic
 ActiveWorkbook.Sheets("Preferences").Activate
End Sub

Sub UnhideSys()
 Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Range("A1").Value2 = "sys" _
        Or ws.Range("A1").Value2 = "Трудоёмкость" _
        Or ws.Range("A1").Value2 = "Статья затрат" Or ws.Range("A1").Value2 = "Имя" _
        Or ws.Range("A1").Value2 = "company_name" _
        Or ws.Range("A1").Value2 = "Наименование статей в 1С" _
        Or ws.Range("H2").Value2 = "Отчет о финансовых результатах" _
        Or ws.Range("A1").Value2 = "organization_id" _
        Or ws.Range("J1").Value2 = "Сумма" Then
            ws.Visible = True
        End If
    Next ws
 ActiveWorkbook.Sheets("Preferences").Activate
End Sub

Sub HideEmpty()
 Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Range("A1").Value2 = "1" Then
            ws.Visible = True
            ws.Select
            With ws.Tab
                .ColorIndex = xlNone
                .TintAndShade = 0
            End With
            ws.Visible = False
        End If
    Next ws
 ActiveWorkbook.Sheets("Preferences").Activate
End Sub


Sub RefreshAllTables_new()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim info As String
    Dim pt As PivotTable
    Dim totalTables As Long
    Dim processedTables As Long
    Dim percentage As Double
    Dim msg As String

    ' Счетчик для общего количества таблиц
    totalTables = 0
    processedTables = 0

    ' Подсчет всех таблиц
    For Each ws In ThisWorkbook.Worksheets
        totalTables = totalTables + ws.ListObjects.Count
    Next ws

    ' Запретить использование Excel во время выполнения
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' Обновление таблиц
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            On Error Resume Next
            
            ' Обновление таблицы
            lo.QueryTable.Refresh BackgroundQuery:=False
            lo.Refresh
            
            ' Обновление прогресса
            processedTables = processedTables + 1
            percentage = (processedTables / totalTables) * 100
            
            ' Обновление сообщения
            msg = "Обновление таблиц... " & Round(percentage, 2) & "% завершено."
            Application.StatusBar = msg
            
            ' Динамическое обновление интерфейса
            DoEvents
        Next lo
    Next ws

    ' Обновление сводных таблиц
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.Refresh
        Next pt
    Next ws

    ' Завершение
    Application.StatusBar = "Обновление завершено на 100%."
'    MsgBox "Обновление таблиц завершено успешно!", vbInformation, "Завершено"

    ' Восстановление параметров Excel
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
End Sub

Sub Protect()
 Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect Password:="123$"
    Next ws
    ThisWorkbook.Protect Password:="123$"
 ActiveWorkbook.Sheets("Preferences").Activate
End Sub

Sub UnProtect()
 Dim ws As Worksheet
 On Error GoTo errorhandler
    For Each ws In ThisWorkbook.Worksheets
        ws.UnProtect Password:="123$"
    Next ws
    ThisWorkbook.UnProtect Password:="123$"

errorhandler:
' MsgBox ("Все листы разблокированы")
ActiveWorkbook.Sheets("Preferences").Activate

End Sub
