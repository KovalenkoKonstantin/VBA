Attribute VB_Name = "PageCount"
Public Sub CountPages()
Dim ThisWorkbook As Workbook
Set ThisWorkbook = ActiveWorkbook

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

On Error Resume Next

X = 0
ThisWorkbook.Sheets("Preferences").Activate
DistinctColumn = Rows(1).Find("Содержание", LookIn:=xlValues).column 'колонка содержания
ThisWorkbook.Sheets("Опись").Activate
NameColumn = 3 'Rows(5).Find("Наименование документа", LookIn:=xlValues).column 'колонка наименования

I = "1"
j = "Протокол (согласование) цены (прогнозной*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
    X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
'[d7] = X
X = 0

I = "2"
j = "Плановая калькуляция затрат (ф. № 2*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
I = "2_22"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = X + ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
I = "2_23"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = X + ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
I = "2_24"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = X + ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
I = "2_1"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = X + ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
I = "2_1_22"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = X + ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
I = "2_1_23"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = X + ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
I = "2_2"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = X + ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
I = "2_2_23"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = X + ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
I = "2_2_24"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = X + ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
'[d8] = X
X = 0

'[d9] = 0
'[d10] = 0
'[d11] = 0
'[d12] = 0

I = "6"
j = "Расшифровка затрат на приобретение комплектующих изделий (ф. № 6 (6д)"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
    X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
'[d13] = X
X = 0

'[d14] = 0
'[d15] = 0

I = "7_1"
j = "Расшифровка затрат по работам (услугам), выполняемым (оказываемым) сторонними организациями (ф. № 7 (7д) НИР (ОКР)"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
    X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
'[d16] = X
X = 0


'[d17] = 0
'[d18] = 0
'[d19] = 0

I = "9_1э"
j = "Расшифровка основной заработной платы (ф. № 9 (9д)*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
I = "9_2э"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
'[d20] = X
X = 0

I = "10"
j = "Расчет-обоснование уровня (*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
'[d21] = X
X = 0

'[d22] = 0

I = "12"
j = "Смета и расчет общехозяйственных затрат/административно-управленческих расходов*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
'[d23] = X
X = 0
'
'[d24] = 0
'[d25] = 0
'[d26] = 0
'[d27] = 0
'[d28] = 0
'[d29] = 0
'[d30] = 0
'[d31] = 0
'[d32] = 0

I = "18"
j = "Расшифровка прочих прямых затрат (ф. № 18 (18д*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
'[d33] = 0
X = 0

'[d34] = 0

I = "20"
j = "Расчет и обоснование прибыли (ф. № 20 (20д*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = ThisWorkbook.Sheets("20").PageSetup.Pages.Count
End If
I = "20_1"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
    X = X + ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
I = "20_2"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
    X = X + ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
'[d35] = X
X = 0

I = "21ф"
j = "Сведения об объемах поставки продукции, в т. ч. по ГОЗ, включая экспортные поставки (ф. № 21*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
    X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
'[d36] = X
X = 0

I = "22ф"
j = "Сведения о нормативах и экономических показателях организации*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
    X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
I = "22ф-23"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
    X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
I = "22ф-24"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
    X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
X = 0

I = "23ф"
j = "Расчет (обоснование) трудоемкости (ф. № 23)"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
'[d38] = X
X = 0

I = "Приказ"
j = "Приказы (Перечень специалистов задействованных*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
'[d39] = X
X = 0

I = "П5"
j = "Сводная ведомость с указанием ФИО*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
'[d40] = X
X = 0

I = "НЧ"
j = "Перечень с указанием размера налоговых*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
'[d41] = X
X = 0

I = "П6"
j = "Справка по отчислениям в фонды*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
'[d42] = X
X = 0

I = "П7"
j = "Справка по статье *"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
'[d43] = X
X = 0

'[d44] = 0

j = "Учётная политика"
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = 40

'[d46] = 0

I = "П8"
j = "Расшифровка общехозяйственных затрат ст. *"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
'[d47] = X
X = 0

I = "Табель"
j = "Табеля учёта*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
'[d48] = X
X = 0

I = "ПЗ"
j = "Пояснительная*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
'[d49] = X
X = 0

I = "КУЗ"
j = "Карточка учёта*"
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
'[d50] = X
X = 0

I = "ШР1"
j = "Штатное расписание от 01.06.2022 г."
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
    X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
X = 0

I = "ШР2"
j = "Штатное расписание от 01.03.2023 г."
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
    X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
X = 0

I = "ШР3"
j = "Штатное расписание от 01.04.2023 г."
ThisWorkbook.Sheets("Preferences").Activate
RowData = Columns(DistinctColumn).Find(I, LookIn:=xlValues).row 'ряд значения
If ThisWorkbook.Sheets("Preferences").Range("W" & RowData).Value2 = True Then
    X = ThisWorkbook.Sheets(I).PageSetup.Pages.Count
End If
ThisWorkbook.Sheets("Опись").Activate
RowKey = Columns(NameColumn).Find(j, LookIn:=xlValues).row 'ряд ключа
ThisWorkbook.Sheets("Опись").Cells(RowKey, NameColumn + 1).Value2 = X
X = 0

'[d52] = 7
'[d53] = 6
'[d54] = 1
'[d55] = 0
'[d56] = 7
'[d57] = 2
'[d58] = 2
'[d59] = 17

Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True

End Sub

