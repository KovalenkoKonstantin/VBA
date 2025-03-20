Attribute VB_Name = "InsertionCUR"
Sub Data_Insertion_CUR()
 
 Start = Now()
 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim ws, this As Worksheet
 Dim MyRange, MyCell As Range
 Dim SheetName, CompanyName As String
 Dim dict As Object
 Dim columnsToFormat As Variant
 Dim bound As Integer
 Dim Limit As Integer
 Set dict = CreateObject("Scripting.Dictionary")
 ps = "123$"
  
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 SheetName = "ProcessingCUR"
 DistinctYear = Year(Date)
 Limit = 153 'последняя колонка базы
 begin = 12 'первый ряд вставки
 CompanyName = ThisWorkbook.Sheets("Preferences").Range("C7").Value2 'имя проекта
 
' Dim aw(1 To 147) As Variant
' Dim iw(1 To 147) As Variant

' Объявляем массивы как динамические
Dim aw() As Variant
Dim iw() As Variant

'Устанавливаем размер массивов
ReDim aw(1 To Limit)
ReDim iw(1 To Limit)

 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual

FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="Выберите расчётную ведомость по компании " _
 & CompanyName & " за " & Year(Date) & " год")
 
 'статус бар
Application.StatusBar = "Анализ данных..."

 If TypeName(FilesToOpen) = "Boolean" Then ',была нажата кнопка отмены выход из процедуры
 GoTo ExitHandler
 End If

'вставка листов
Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

On Error Resume Next

 'проверка правильности выбора данных
 importWB.Sheets(1).Activate
 Range("G2").Select
 ActiveCell.FormulaR1C1 = "=YEAR(MID(RC[-4],SEARCH("" "",RC[-4],1)+1,10))"
' If Range("G2").Value2 <> DistinctYear Or Range("A11").Value2 <> CompanyName Then
 If Range("A11").Value2 <> CompanyName Then
    Range("G2").Select
    With Selection:
        .Clear
    End With
    MsgBoxEx "Выбрана неправильная расчётная ведомость. Наименование компании не совпадает." _
    & vbCr & "Процесс прерван.", vbCritical, "Bad Day", 20
    GoTo ExitHandler
 ElseIf Range("G2").Value2 = DistinctYear Then
 Range("G2").Select
    With Selection:
        .Clear
    End With
 End If

ThisWorkbook.UnProtect Password:=ps
ThisWorkbook.Sheets(SheetName).Visible = True
ThisWorkbook.Sheets(SheetName).UnProtect Password:=ps
ThisWorkbook.Sheets(SheetName).Activate
On Error Resume Next
ActiveSheet.ShowAllData

ThisWorkbook.Activate

'определение колонок рабочей книги
On Error Resume Next
For i = 1 To 20
    If Worksheets(SheetName).Cells(i, 1) = "Сотрудник" Then
        DataRow = i
    End If
Next i

' Заполнение словаря с ключами и соответствующими индексами
    dict.Add "Сотрудник", 1
    dict.Add "Месяц", 2
    dict.Add "расчётная норма часов", 3
    dict.Add "анализ часов", 4
    dict.Add "расчётная норма дней", 5
    dict.Add "анализ дней", 6
    dict.Add "Исключение всех кроме 20,26,44 счёта", 7
    dict.Add "Имя Отчество", 8
    dict.Add "Анализ изменения фамилии", 9
    dict.Add "Приказ об увольнении.Статья ТК РФ", 10
    dict.Add "Подразделение история", 11
    dict.Add "Должность", 12
    dict.Add "Вид занятости", 13
    dict.Add "Дата рождения", 14
    dict.Add "Способ отражения зарплаты в бух учете", 15
    dict.Add "График работы", 16
    dict.Add "Норма дней", 17
    dict.Add "Норма часов", 18
    dict.Add "Дней", 19
    dict.Add "Часов", 20
    dict.Add "Начислено", 21
    dict.Add "Оплата больничных листов за счет работодателя", 22
    dict.Add "Оплата больничных листов", 23
    dict.Add "Оклад по дням", 24
    dict.Add "Отсутствие по болезни (больничный еще не закрыт)", 25
    dict.Add "Надбавка за сложность и напряженность", 26
    dict.Add "Оплата отпуска по календарным дням", 27
    dict.Add "Премия за объем продаж", 28
    dict.Add "Надбавка за профессионализм и качество выполняемых работ", 29
    dict.Add "Премия по итогам года", 30
    dict.Add "Отпуск за свой счет", 31
    dict.Add "Компенсация отпуска (Отпуск основной)", 32
    dict.Add "Квартальная премия", 33
    dict.Add "Премия разовая", 34
    dict.Add "Выходное пособие при увольнении", 35
    dict.Add "Пособие по уходу за ребенком до полутора лет", 36
    dict.Add "Оклад по дням (пропорционально отработанным дням)", 37
    dict.Add "Надбавка за профессионализм и качество выполняемых работ (пропорционально отработанным дням)", 38
    dict.Add "Надбавка за работу со сведениями, составляющими государственную тайну", 39
    dict.Add "Оклад по дням 26 Нерезиденты", 40
    dict.Add "Премия не учитываемая", 41
    dict.Add "Северная надбавка", 42
    dict.Add "Районный коэффициент", 43
    dict.Add "Премия с учетом районного коэффициента(квартальная)", 44
    dict.Add "Ежегодный дополнительный оплачиваемый отпуск", 45
    dict.Add "Премия по итогам года с учетом РК", 46
    dict.Add "Вознаграждение за изобретение", 47
    dict.Add "Договор (работы, услуги)", 48
    dict.Add "Компенсация за фитнес", 49
    dict.Add "Премия с учетом районного коэффициента (месячная)", 50
    dict.Add "Прогул", 51
    dict.Add "Доплата за совмещение должностей, исполнение обязанностей", 52
    dict.Add "Месячная премия", 53
    dict.Add "Военные сборы", 54
    dict.Add "Дополнительный учебный отпуск (оплачиваемый)", 55
    dict.Add "Оплата работы в праздничные и выходные дни", 56
    dict.Add "Вознаграждение членам Совета Директоров", 57
    dict.Add "Премия Германия У.Е", 58
    dict.Add "Отсутствие по невыясненной причине", 59
    dict.Add "Отпуск без оплаты согласно ТК РФ", 60
    dict.Add "Оплата дней ухода за детьми-инвалидами", 61
    dict.Add "Пособие по уходу за ребёнком до 3 лет без оплаты", 62
    dict.Add "Оклад по часам (пропорциально отработанному времени)", 63
    dict.Add "Надбавка за сложность и напряженность (по часам пропорционально отработаннму времени)", 64
    dict.Add "Надбавка за профессионализм и качество выполняемых работ (по часам пропорц. отраб. времени)", 65
    dict.Add "Отсутствие по болезни", 66
    dict.Add "Доплата за работу в ночное время", 67
    dict.Add "По итогам работы за год", 68
    dict.Add "Премия директорам (Чефранова, Басов, Набережный, Таранов)", 69
    dict.Add "Надбавка за соблюдение конфиденциальности в отношении с", 70
    dict.Add "Отпуск по беременности и родам", 71
    dict.Add "Доплата за работу в праздничные дни (дневное время)", 72
    dict.Add "Доплата за работу в праздничные дни (ночное время)", 73
    dict.Add "Доплата за переработки при суммированном учете рабочего времени", 74
    dict.Add "Неоплачиваемые дни отпуска по беременности и родам", 75
    dict.Add "Доплата за ненормированный рабочий день", 76
    dict.Add "Материальная помощь", 77
    dict.Add "Сумма по Доход в натуральной форме", 78
    dict.Add "Питание за счет средств предприятия", 79
    dict.Add "Доход в натуральной форме (обл. НДФЛ)", 80
    dict.Add "Проездные", 81
    dict.Add "НДФЛ к зачету в счет будущих платежей", 82
    dict.Add "Питание договорников", 83
    dict.Add "Подарок", 84
    dict.Add "Проездные Германия", 85
    dict.Add "Зачтено излишне удержанного НДФЛ", 86
    dict.Add "Сотовая связь", 87
    dict.Add "Удержано", 88
    dict.Add "НДФЛ", 89
    dict.Add "НДФЛ с превышения", 90
    dict.Add "Удержано из зарплаты за оплату медстраховки", 91
    dict.Add "Удержание из з/п командировочных расходов", 92
    dict.Add "Удержание по исп. листу процентом", 93
    dict.Add "Удержание по заявлению", 94
    dict.Add "Выплата родственникам умершего сотрудника", 95
    dict.Add "Удержано из зарплаты за оплату физкультурно-оздоровительных услуг", 96
    dict.Add "Удержание за обучение из зарплаты", 97
    dict.Add "Удержание из зарплаты 73,03", 98
    dict.Add "Удержание по исп. листу фикс. суммой", 99
    dict.Add "Удержание по исполнительному документу", 100
    dict.Add "ВЗНОСЫ", 101
    dict.Add "ПФР до превышения", 102
    dict.Add "ПФР с превышения", 103
    dict.Add "ФФОМС", 104
    dict.Add "ФСС", 105
    dict.Add "ФСС НС", 106
    dict.Add "% Страховых взносов", 107
    dict.Add "База взносов", 108
    dict.Add "Надбавка за стаж работы по защите государственной тайны", 109
    dict.Add "Оплата за дополнительный день (дни) отдыха донору", 110
    dict.Add "Командировка", 111
    dict.Add "Премия квартальная объем продаж (ПМ)", 112
    dict.Add "Премия Германия", 113
    dict.Add "Оплата по окладу (по часам)", 114
    dict.Add "Оклад по часам", 115
    dict.Add "Надбавка за сложность и напряженность (по часам)", 116
    dict.Add "Надбавка за сложность и напряженность (пропорционально отработанным дням)", 117
    dict.Add "Медицинский осмотр", 118
    dict.Add "Компенсация расходов по договорам подряда", 119
    dict.Add "Премия месячная (с учетом РК)", 120
    dict.Add "Премия квартальная", 121
    dict.Add "Премия квартальная (с учетом РК)", 122
    dict.Add "Премия по итогам года (с учетом РК)", 123
    dict.Add "Надбавка за сложность и напряженность (по часам пропорц. отработанному времени)", 124
    dict.Add "Премия разовая (с учетом РК)", 125
    dict.Add "Дата увольнения", 126
    dict.Add "Премия месячная", 127
    dict.Add "Год", 128
    dict.Add "ФОТ ИТОГО", 133
    dict.Add "Доплата за работу в ночное время (праздничные и выходные дни)", 134
    dict.Add "Премия полугодовая (с учетом РК)", 135
    dict.Add "Премия полугодовая", 136
    dict.Add "Компенсация отпуска (Отпуск лицам, работающим в районах Крайнего Севера)", 137
    dict.Add "База взносов ДГПХ", 138
    dict.Add "% Страховых взносов ДГПХ", 139
    dict.Add "Доход в натуральной форме (ГПД)", 140
    dict.Add "Оплата отпуска по календарным дням (до 2025)", 141
    dict.Add "Оплата больничных листов за счет работодателя (до 2025)", 142
    dict.Add "Компенсация отпуска (Отпуск основной) (до 2025)", 143
    dict.Add "Дополнительный учебный отпуск (оплачиваемый) (до 2025)", 144
    dict.Add "Оплата за дополнительный день (дни) отдыха донору (до 2025)", 145
    dict.Add "Компенсация отпуска (Отпуск лицам, работающим в районах Крайнего Севера) (до 2025)", 146
    dict.Add "Оплата отпуска по календарным дням (доля РК)", 147
    dict.Add "Компенсация отпуска при увольнении по календарным дням", 148
    dict.Add "Компенсация отпуска при увольнении по календарным дням (доля РК)", 149
    dict.Add "Компенсация отпуска (Отпуск лицам, работающим в районах Крайнего Севера) (доля РК)", 150
    dict.Add "Компенсация отпуска (Отпуск лицам, работающим в районах Крайнего Севера) (доля СН)", 151
    dict.Add "Компенсация отпуска при увольнении по календарным дням (доля СН)", 152
    dict.Add "Оплата отпуска по календарным дням (доля СН)", 153


' Перебор значений в словаре и заполнение массива
    For Each Key In dict.Keys
        For i = 1 To Limit
            Item = dict(Key)
            If Worksheets(SheetName).Cells(DataRow, i) = Key Then
                aw(Item) = i
            End If
        Next i
    Next Key

 
 importWB.Sheets(1).Activate

'определение колонок импортируемой книги
For i = 1 To 20
    If importWB.Sheets(1).Cells(i, 1) = "Организация" Then
        ImportFirstDataRow = i
    End If
Next i
For i = 1 To 20
    If importWB.Sheets(1).Cells(i, 1) = "Сотрудник" Then
        ImportSecondDataRow = i
    End If
Next i

' Перебор значений в словаре и заполнение массива
    For Each Key In dict.Keys
        For i = 1 To Limit
            Item = dict(Key)
            If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = Key Then
                iw(Item) = i
            End If
            If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = Key Then
                iw(Item) = i
            End If
        Next i
    Next Key


'удаление предыдущих данных
ThisWorkbook.Sheets(SheetName).Activate
awLastRow = Cells(Rows.Count, "A").End(xlUp).row
awLastCol = Cells(begin, Columns.Count).End(xlUp).column
Range(Cells(begin, 1), Cells(awLastRow, Limit)).Select
 With Selection
        .Clear
 End With

 'статус бар
Application.StatusBar = "Вставка данных"

 'вставка данных
 importWB.Sheets(1).Activate
 iwLastRow = Cells(Rows.Count, "A").End(xlUp).row

For i = 1 To Limit
'статус бар
Application.StatusBar = "Промежуточный цикл. Выполнено: " & Int(100 * i / Limit) & "%." & _
" Общий прогресс: " & Int(87 * i / Limit) & "%" & _
" Расчётное время до конца выполнения программы: " & _
Int((100 - Int(87 * i / Limit)) * (((Now() - Start) * 24 * 60 * 60) / (Int(87 * i / Limit)))) & " секунд"
 importWB.Activate
 Range(Cells(begin - 1, iw(i)), Cells(iwLastRow, iw(i))).Copy

 ThisWorkbook.Sheets(SheetName).Activate
 Range(Cells(begin, aw(i)), Cells(iwLastRow, aw(i))).Select
    With Selection
           .PasteSpecial Paste:=xlPasteAll
           .UnMerge
           .Font.Name = "Times New Roman"
           .WrapText = False
           .MergeCells = False
           .Font.Size = 8
    End With

Next i

'статус бар
Application.StatusBar = "Форматирование ячеек. Выполнено: 87 %"

'форматы
ThisWorkbook.Sheets(SheetName).Activate
Columns("Q:DD").Select
    Selection.NumberFormat = _
        "_-* #,##0.00 _?_-;-* #,##0.00 _?_-;_-* ""-""?? _?_-;_-@_-"

'вставка проверочных формул
ThisWorkbook.Sheets(SheetName).Activate
'статус бар
Application.StatusBar = "Добавление проверочных формул месяца. Выполнено: 88 %"
'месяц
Cells(begin, aw(2)).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(IFERROR(SEARCH("" 20"",RC[-1],1)>0,FALSE)," _
        & "TRIM(MID(IF(IFERROR(SEARCH("" 20"",RC[-1],1)>0,FALSE)," _
        & "RC[-1],R[-1]C),1,SEARCH("" "",IF(IFERROR(SEARCH("" 20"",RC[-1],1)>0,FALSE)," _
        & "RC[-1],R[-1]C),1)-1)),R[-1]C)"
    Cells(begin, aw(2)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(2)), Cells(iwLastRow, aw(2)))

'статус бар
Application.StatusBar = "Добавление формул расчётной нормы часов. Выполнено: 89 %"
'расчётная норма часов
K = 3
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = "=IF(OR(CONCATENATE(RC[-1],"" "",RC[5])=RC[-2]," _
    & "VLOOKUP(RC[-2],RC1:RC214,MATCH(R8C4,R9C1:R9C214,0),0)>0," _
    & "VLOOKUP(RC[-2],RC1:RC214,MATCH(R7C4,R9C1:R9C214,0),0)>0," _
    & "VLOOKUP(RC[-2],RC1:RC214,MATCH(R6C4,R9C1:R9C214,0),0)>0," _
    & "RC[4]=TRUE),"""",VLOOKUP(RC[-1],INDIRECT(CONCATENATE(""'"",VALUE(RC[5])," _
    & """ произ. календарь'!$A:$BR"")),HLOOKUP(RC[20],INDIRECT" _
    & "(CONCATENATE(""'"",VALUE(RC" & _
        "[5]),"" произ. календарь'!$2:$3"")),2,0),0))" & _
        ""
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
    
'статус бар
Application.StatusBar = "Добавление формул анализа нормы часов. Выполнено: 90 %"
'анализ часов
K = 4
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
        "=OR(RC[-1]="""",OR(RC[-1]=" _
        & "VALUE(RC[21]),VLOOKUP(RC[-3],RC1:RC114," _
        & "MATCH(R8C4,R9C1:R9C114,0),0)>0),SUM(RC[22]:RC[124])=0," _
        & "NOT(ISNA(MATCH(RC[-3]&RC[-2]&RC[4],Изм.граф!C7,0))))"
    Range("A1").Copy
    Cells(begin, aw(K)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
   
'статус бар
Application.StatusBar = "Добавление формул нормы дней. Выполнено: 91 %"
'расчётная норма дней
K = 5
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(CONCATENATE(RC[-3],"" "",RC[3])=RC[-4]," _
        & "VLOOKUP(RC[-4],RC1:RC114,MATCH(R8C4,R9C1:R9C114,0),0)>0," _
        & "RC[2]=TRUE),"""",VLOOKUP(RC[-3],INDIRECT(CONCATENATE(""'"",VALUE(RC[3])," _
        & """ произ. календарь'!$A$18:$BR$31"")),HLOOKUP(RC[18]," _
        & "INDIRECT(CONCATENATE(""'"",VALUE(RC[3]),"" произ. календарь'!$18:$19"")),2,0),0))"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
    
'статус бар
Application.StatusBar = "Добавление формул анализа нормы дней. Выполнено: 92 %"
'анализ дней
K = 6
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
        "=OR(RC[-3]="""",OR(RC[-1]=VALUE(RC[18]),VLOOKUP(RC[-5]," _
        & "RC1:RC114,MATCH(R8C4,R9C1:R9C114,0),0)>0),SUM(RC[20]:RC[122])=0," _
        & "NOT(ISNA(MATCH(RC[-5]&RC[-4]&RC[2],Изм.граф!C7,0))))"
    Range("A1").Copy
    Cells(begin, aw(K)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))

'статус бар
Application.StatusBar = "Добавление формул исключения из расчёта всех кроме 20,26,44 счетов. Выполнено: 93 %"
'Исключение всех кроме 20,26,44 счёта
K = 7
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
       "=NOT(OR(IFERROR(SEARCH(20,RC[15],1),FALSE)," _
       & "IFERROR(SEARCH(26,RC[15],1),FALSE),IFERROR(SEARCH(44,RC[15],1),FALSE)))"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
    
'статус бар
Application.StatusBar = "Добавление формул выделения Имени Отчества сотрудников. Выполнено: 94 %"
'Имя Отчество
K = 8
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(CONCATENATE(RC[-12],"" "",RC[-6])=RC[-13],""""," _
        & "(MID(RC[-13],SEARCH("" "",RC[-13],1),LEN(RC[-13]))))"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))

'статус бар
Application.StatusBar = "Добавление формул и форматирование анализа изменения фамилии. Выполнено: 95 %"
'Анализ изменения фамилии
K = 9
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-14],RC[-13],RC[5])"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
        
Columns("O:O").Select
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Font
        .Color = -16754788
'        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10284031
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False


'статус бар
Application.StatusBar = "Добавление формул по доходу в натуральной форме. Выполнено: 96 %"
'Сумма по Доход в натуральной форме
K = 78
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[1]:RC[9])"
    Cells(begin, aw(K)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 12517371
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = -13680896
        .TintAndShade = 0
    End With
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
    
'статус бар
Application.StatusBar = "Добавление формул выделения года. Выполнено: 97 %"
'Год
K = 128
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
       "=VALUE(IF(IFERROR(SEARCH("" 20"",RC[-7],1)>0,FALSE)," _
        & "MID(RC[-7],SEARCH("" "",RC[-7],1)+1,4),R[-1]C))"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))


'статус бар
Application.StatusBar = "Форматирование диапазонов. Выполнено: 98 %"

Columns("A:A").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = -1
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

Range("C:C,E:E").Select
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    
'статус бар
Application.StatusBar = "Форматирование диапазонов. Выполнено: 99 %"
' форматирование колонки базы страховых взносов
'Массив с нужными столбцами
columnsToFormat = Array("EC", "DW", "DZ", "ED", "EI", "EH", "EO", "EQ", "EP", _
            "ER", "ES", "EU", "EV", "EW")

' Проходим по каждому столбцу в массиве и задаём формат
For bound = LBound(columnsToFormat) To UBound(columnsToFormat)
    Columns(columnsToFormat(bound) & ":" & columnsToFormat(bound)).NumberFormat = _
        "_-* #,##0.00 _?_-;-* #,##0.00 _?_-;_-* ""-""?? _?_-;_-@_-"
Next bound

'статус бар
Application.StatusBar = "Выполнено: 100 %"

'завершение

'ThisWorkbook.Sheets(SheetName).Protect Password:=ps
'ThisWorkbook.Sheets(SheetName).Visible = False
'MsgBoxEx "Расчётная ведомость " _
'    & "по компании " & vbCr & ThisWorkbook.Sheets("Preferences").Range("C7").Value2 _
'    & vbCr & "за " & ThisWorkbook.Sheets(SheetName).Range("B2").Value2 & " год" _
'    & vbCr & "добавлена успешно", 0, "Выполнено", 25

'ThisWorkbook.Sheets("Calculation23").Activate

ExitHandler:
    On Error Resume Next
    importWB.Close
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    Application.Calculation = xlAutomatic
    ThisWorkbook.Sheets("Preferences").Activate
    ThisWorkbook.Sheets("Preferences").UnProtect Password:=ps
    Rows("82:93").EntireRow.AutoFit
'    ThisWorkbook.Sheets("Preferences").Protect Password:=ps
'    ThisWorkbook.Protect Password:=ps
 Exit Sub
  
errhandler:
 MsgBox Err.Description
 Resume ExitHandler
End Sub








