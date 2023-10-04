Attribute VB_Name = "Expenditures"
Sub Overheads()
 
 Start = Now()
 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim ws, this As Worksheet
 Dim MyRange, MyCell As Range
 Dim SheetName, CompanyName As String
  
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 SheetName = "Expenditures"
' DistinctYear = 2021
 Limit = 138 'последняя колонка базы
 begin = 12 'первый ряд вставки
 CompanyName = ThisWorkbook.Sheets("Preferences").Range("C7").Value2 'имя проекта
 
 Dim aw(1 To 138) As Variant
 Dim iw(1 To 138) As Variant
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual

FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="Выберите расчётную ведомость по компании " & CompanyName)
 
 MsgBoxEx "Расчётная ведомость должна быть с начала года!" _
 & vbCr & "В противном случае, не все проверки буду отработанны корректно." _
    & vbCr & "...", vbCritical, "Pay attention", 20
 
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
'    Range("G2").Select
'    With Selection:
'        .Clear
'    End With
    MsgBoxEx "Выбрана неправильная расчётная ведомость." _
    & vbCr & "Процесс прерван.", vbCritical, "Bad Day", 20
    GoTo ExitHandler
' ElseIf Range("G2").Value2 = DistinctYear Then
' Range("G2").Select
'    With Selection:
'        .Clear
'    End With
 End If
 
  Range("G2").Select
    With Selection:
        .Clear
    End With

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

For i = 1 To Limit
    If Worksheets(SheetName).Cells(DataRow, i) = "Сотрудник" Then
        aw(1) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Месяц" Then
        aw(2) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "расчётная норма часов" Then
        aw(3) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "анализ часов" Then
        aw(4) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "расчётная норма дней" Then
        aw(5) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "анализ дней" Then
        aw(6) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Исключение всех кроме 20,26,44 счёта" Then
        aw(7) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Имя Отчество" Then
        aw(8) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Анализ изменения фамилии" Then
        aw(9) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Приказ об увольнении.Статья ТК РФ" Then
        aw(10) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Подразделение история" Then
        aw(11) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Должность" Then
        aw(12) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Вид занятости" Then
        aw(13) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Дата рождения" Then
        aw(14) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Способ отражения зарплаты в бух учете" Then
        aw(15) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "График работы" Then
        aw(16) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Норма дней" Then
        aw(17) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Норма часов" Then
        aw(18) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Отработано дней" Then
        aw(19) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Отработано часов" Then
        aw(20) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, i) = "Начислено" Then
        aw(21) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Оплата больничных листов за счет работодателя" Then
        aw(22) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Оплата больничных листов" Then
        aw(23) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Оклад по дням" Then
        aw(24) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Отсутствие по болезни (больничный еще не закрыт)" Then
        aw(25) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Надбавка за сложность и напряженность" Then
        aw(26) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Оплата отпуска по календарным дням" Then
        aw(27) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия за объем продаж" Then
        aw(28) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Надбавка за профессионализм и качество выполняемых работ" Then
        aw(29) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия по итогам года" Then
        aw(30) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Отпуск за свой счет" Then
        aw(31) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Компенсация отпуска (Отпуск основной)" Then
        aw(32) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Квартальная премия" Then
        aw(33) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия разовая" Then
        aw(34) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Выходное пособие при увольнении" Then
        aw(35) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Пособие по уходу за ребенком до полутора лет" Then
        aw(36) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Оклад по дням (пропорционально отработанным дням)" Then
        aw(37) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Надбавка за профессионализм и качество выполняемых работ (пропорционально отработанным дням)" Then
        aw(38) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Надбавка за работу со сведениями, составляющими государственную тайну" Then
        aw(39) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Оклад по дням 26 Нерезиденты" Then
        aw(40) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия не учитываемая" Then
        aw(41) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Северная надбавка" Then
        aw(42) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Районный коэффициент" Then
        aw(43) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия с учетом районного коэффициента(квартальная)" Then
        aw(44) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Ежегодный дополнительный оплачиваемый отпуск" Then
        aw(45) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия по итогам года с учетом РК" Then
        aw(46) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Вознаграждение за изобретение" Then
        aw(47) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Договор (работы, услуги)" Then
        aw(48) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Компенсация за фитнес" Then
        aw(49) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия с учетом районного коэффициента (месячная)" Then
        aw(50) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Прогул" Then
        aw(51) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Доплата за совмещение должностей, исполнение обязанностей" Then
        aw(52) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Месячная премия" Then
        aw(53) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Военные сборы" Then
        aw(54) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Дополнительный учебный отпуск (оплачиваемый)" Then
        aw(55) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Оплата работы в праздничные и выходные дни" Then
        aw(56) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "вознаграждение членам Совета Директоров" Then
        aw(57) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия Германия У.Е" Then
        aw(58) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Отсутствие по невыясненной причине" Then
        aw(59) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Отпуск без оплаты согласно ТК РФ" Then
        aw(60) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Оплата дней ухода за детьми-инвалидами" Then
        aw(61) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Пособие по уходу за ребёнком до 3 лет без оплаты" Then
        aw(62) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Оклад по часам (пропорциально отработанному времени)" Then
        aw(63) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Надбавка за сложность и напряженность (по часам пропорционально отработаннму времени)" Then
        aw(64) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Надбавка за профессионализм и качество выполняемых работ (по часам пропорц отраб времени)" Then
        aw(65) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Отсутствие по болезни" Then
        aw(66) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Доплата за работу в ночное время" Then
        aw(67) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "По итогам работы за год" Then
        aw(68) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия директорам (Чефранова, Басов, Набережный, Таранов)" Then
        aw(69) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Надбавка за соблюдение конфиденциальности в отношении с" Then
        aw(70) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Отпуск по беременности и родам" Then
        aw(71) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Доплата за работу в праздничные дни (дневное время)" Then
        aw(72) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Доплата за работу в праздничные дни (ночное время)" Then
        aw(73) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Доплата за переработки при суммированном учете рабочего времени" Then
        aw(74) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Неоплачиваемые дни отпуска по беременности и родам" Then
        aw(75) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Доплата за ненормированный рабочий день" Then
        aw(76) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Материальная помощь" Then
        aw(77) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, i) = "Сумма по Доход в натуральной форме" Then
        aw(78) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Питание за счет средств предприятия" Then
        aw(79) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Доход в натуральной форме (обл. НДФЛ)" Then
        aw(80) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Проездные" Then
        aw(81) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "НДФЛ к зачету в счет будущих платежей" Then
        aw(82) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Питание договорников" Then
        aw(83) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Подарок" Then
        aw(84) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Проездные Германия" Then
        aw(85) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Зачтено излишне удержанного НДФЛ" Then
        aw(86) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "сотовая связь" Then
        aw(87) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, i) = "Удержано" Then
        aw(88) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "НДФЛ" Then
        aw(89) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "НДФЛ с превышения" Then
        aw(90) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Удержано из зарплаты за оплату медстраховки" Then
        aw(91) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Удержание из з/п командировочных расходов" Then
        aw(92) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Удержание по исп. листу процентом" Then
        aw(93) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Удержание по заявлению" Then
        aw(94) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Выплата родственникам умершего сотрудника" Then
        aw(95) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Удержано из зарплаты за оплату физкультурно-оздоровительных услуг" Then
        aw(96) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Удержание за обучение из зарплаты" Then
        aw(97) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Удержание из зарплаты 73,03" Then
        aw(98) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Удержание по исп. листу фикс. суммой" Then
        aw(99) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Удержание по исполнительному документу" Then
        aw(100) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, i) = "ВЗНОСЫ" Then
        aw(101) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "ПФР до превышения" Then
        aw(102) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "ПФР с превышения" Then
        aw(103) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "ФФОМС" Then
        aw(104) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "ФСС" Then
        aw(105) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "ФСС НС" Then
        aw(106) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, i) = "% Страховых взносов" Then
        aw(107) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "База взносов" Then
        aw(108) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, i) = "Надбавка за стаж работы по защите государственной тайны" Then
        aw(109) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Оплата за дополнительный день (дни) отдыха донору" Then
        aw(110) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Командировка" Then
        aw(111) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия квартальная объем продаж (ПМ)" Then
        aw(112) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия Германия" Then
        aw(113) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Оплата по окладу (по часам)" Then
        aw(114) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Оклад по часам" Then
        aw(115) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Надбавка за сложность и напряженность (по часам)" Then
        aw(116) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Надбавка за сложность и напряженность (пропорционально отработанным дням)" Then
        aw(117) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Медицинский осмотр" Then
        aw(118) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Компенсация расходов по договорам подряда" Then
        aw(119) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия месячная (с учетом РК)" Then
        aw(120) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия квартальная" Then
        aw(121) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия квартальная (с учетом РК)" Then
        aw(122) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия по итогам года (с учетом РК)" Then
        aw(123) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Надбавка за сложность и напряженность (по часам пропорц. отработанному времени)" Then
        aw(124) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия разовая (с учетом РК)" Then
        aw(125) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Дата увольнения" Then
        aw(126) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия месячная" Then
        aw(127) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Год" Then
        aw(128) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Доплата за работу в ночное время (праздничные и выходные дни)" Then
        aw(134) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия полугодовая (с учетом РК)" Then
        aw(135) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия полугодовая" Then
        aw(136) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Надбавка за профессионализм и качество выполняемых работ (по часам пропорц. отраб. времени)" Then
        aw(137) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Сотовая связь" Then
        aw(138) = i
    End If
    
Next i
 
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

For i = 1 To Limit
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Сотрудник" Then '-
        iw(1) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Месяц" Then
        iw(2) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "расчётная норма часов" Then
        iw(2) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "анализ часов" Then
        iw(4) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "расчётная норма дней" Then
        iw(5) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "анализ дней" Then
        iw(6) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Исключение всех кроме 20,26,44 счёта" Then
        iw(7) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Имя Отчество" Then
        iw(8) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Анализ изменения фамилии" Then
        iw(9) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Приказ об увольнении.Статья ТК РФ" Then
        iw(10) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Подразделение история" Then '-
        iw(11) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Должность" Then '-
        iw(12) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Вид занятости" Then '-
        iw(13) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Дата рождения" Then '-
        iw(14) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Способ отражения зарплаты в бух учете" Then '-
        iw(15) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "График работы" Then '-
        iw(16) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Норма дней" Then '-
        iw(17) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Норма часов" Then '-
        iw(18) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Дней" Then '-
        iw(19) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Часов" Then '-
        iw(20) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Начислено" Then
        iw(21) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Оплата больничных листов за счет работодателя" Then
        iw(22) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Оплата больничных листов" Then
        iw(23) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Оклад по дням" Then
        iw(24) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Отсутствие по болезни (больничный еще не закрыт)" Then
        iw(25) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Надбавка за сложность и напряженность" Then
        iw(26) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Оплата отпуска по календарным дням" Then
        iw(27) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия за объем продаж" Then
        iw(28) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Надбавка за профессионализм и качество выполняемых работ" Then
        iw(29) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия по итогам года" Then
        iw(30) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Отпуск за свой счет" Then
        iw(31) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Компенсация отпуска (Отпуск основной)" Then
        iw(32) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Квартальная премия" Then
        iw(33) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия разовая" Then
        iw(34) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Выходное пособие при увольнении" Then
        iw(35) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Пособие по уходу за ребенком до полутора лет" Then
        iw(36) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Оклад по дням (пропорционально отработанным дням)" Then
        iw(37) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Надбавка за профессионализм и качество выполняемых работ (пропорционально отработанным дням)" Then
        iw(38) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Надбавка за работу со сведениями, составляющими государственную тайну" Then
        iw(39) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Оклад по дням 26 Нерезиденты" Then
        iw(40) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия не учитываемая" Then
        iw(41) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Северная надбавка" Then
        iw(42) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Районный коэффициент" Then
        iw(43) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия с учетом районного коэффициента(квартальная)" Then
        iw(44) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Ежегодный дополнительный оплачиваемый отпуск" Then
        iw(45) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия по итогам года с учетом РК" Then
        iw(46) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Вознаграждение за изобретение" Then
        iw(47) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Договор (работы, услуги)" Then
        iw(48) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Компенсация за фитнес" Then
        iw(49) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия с учетом районного коэффициента (месячная)" Then
        iw(50) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Прогул" Then
        iw(51) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Доплата за совмещение должностей, исполнение обязанностей" Then
        iw(52) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Месячная премия" Then
        iw(53) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Военные сборы" Then
        iw(54) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Дополнительный учебный отпуск (оплачиваемый)" Then
        iw(55) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Оплата работы в праздничные и выходные дни" Then
        iw(56) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "вознаграждение членам Совета Директоров" Then
        iw(57) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия Германия У.Е" Then
        iw(58) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Отсутствие по невыясненной причине" Then
        iw(59) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Отпуск без оплаты согласно ТК РФ" Then
        iw(60) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Оплата дней ухода за детьми-инвалидами" Then
        iw(61) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Пособие по уходу за ребёнком до 3 лет без оплаты" Then
        iw(62) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Оклад по часам (пропорциально отработанному времени)" Then
        iw(63) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Надбавка за сложность и напряженность (по часам пропорционально отработаннму времени)" Then
        iw(64) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Надбавка за профессионализм и качество выполняемых работ (по часам пропорц отраб времени)" Then
        iw(65) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Отсутствие по болезни" Then
        iw(66) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Доплата за работу в ночное время" Then
        iw(67) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "По итогам работы за год" Then
        iw(68) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия директорам (Чефранова, Басов, Набережный, Таранов)" Then
        iw(69) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Надбавка за соблюдение конфиденциальности в отношении с" Then
        iw(70) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Отпуск по беременности и родам" Then
        iw(71) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Доплата за работу в праздничные дни (дневное время)" Then
        iw(72) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Доплата за работу в праздничные дни (ночное время)" Then
        iw(73) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Доплата за переработки при суммированном учете рабочего времени" Then
        iw(74) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Неоплачиваемые дни отпуска по беременности и родам" Then
        iw(75) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Доплата за ненормированный рабочий день" Then
        iw(76) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Материальная помощь" Then
        iw(77) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Сумма по Доход в натуральной форме" Then
        iw(78) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Питание за счет средств предприятия" Then
        iw(79) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Доход в натуральной форме (обл. НДФЛ)" Then
        iw(80) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Проездные" Then
        iw(81) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "НДФЛ к зачету в счет будущих платежей" Then
        iw(82) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Питание договорников" Then
        iw(83) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Подарок" Then
        iw(84) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Проездные Германия" Then
        iw(85) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Зачтено излишне удержанного НДФЛ" Then
        iw(86) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "сотовая связь" Then
        iw(87) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Удержано" Then
        iw(88) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "НДФЛ" Then
        iw(89) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "НДФЛ с превышения" Then
        iw(90) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Удержано из зарплаты за оплату медстраховки" Then
        iw(91) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Удержание из з/п командировочных расходов" Then
        iw(92) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Удержание по исп. листу процентом" Then
        iw(93) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Удержание по заявлению" Then
        iw(94) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Выплата родственникам умершего сотрудника" Then
        iw(95) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Удержано из зарплаты за оплату физкультурно-оздоровительных услуг" Then
        iw(96) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Удержание за обучение из зарплаты" Then
        iw(97) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Удержание из зарплаты 73,03" Then
        iw(98) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Удержание по исп. листу фикс. суммой" Then
        iw(99) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Удержание по исполнительному документу" Then
        iw(100) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "ВЗНОСЫ" Then
        iw(101) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "ПФР до превышения" Then
        iw(102) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "ПФР с превышения" Then
        iw(103) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "ФФОМС" Then
        iw(104) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "ФСС" Then
        iw(105) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "ФСС НС" Then
        iw(106) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "% Страховых взносов" Then
        iw(107) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "База взносов" Then '-
        iw(108) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Надбавка за стаж работы по защите государственной тайны" Then '-
        iw(109) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Оплата за дополнительный день (дни) отдыха донору" Then '-
        iw(110) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Командировка" Then '-
        iw(111) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия квартальная объем продаж (ПМ)" Then '-
        iw(112) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия Германия" Then '-
        iw(113) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Оплата по окладу (по часам)" Then '-
        iw(114) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Оклад по часам" Then '-
        iw(115) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Надбавка за сложность и напряженность (по часам)" Then '-
        iw(116) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Надбавка за сложность и напряженность (пропорционально отработанным дням)" Then '-
        iw(117) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Медицинский осмотр" Then '-
        iw(118) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Компенсация расходов по договорам подряда" Then '-
        iw(119) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия месячная (с учетом РК)" Then '-
        iw(120) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия квартальная" Then '-
        iw(121) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия квартальная (с учетом РК)" Then '-
        iw(122) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия по итогам года (с учетом РК)" Then '-
        iw(123) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Надбавка за сложность и напряженность (по часам пропорц. отработанному времени)" Then '-
        iw(124) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия разовая (с учетом РК)" Then '-
        iw(125) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Дата увольнения" Then '-
        iw(126) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия месячная" Then '-
        iw(127) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Доплата за работу в ночное время (праздничные и выходные дни)" Then '-
        iw(134) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия полугодовая (с учетом РК)" Then '-
        iw(135) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия полугодовая" Then '-
        iw(136) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Надбавка за профессионализм и качество выполняемых работ (по часам пропорц. отраб. времени)" Then '-
        iw(137) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Сотовая связь" Then '-
        iw(138) = i
    End If
    

Next i

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
" Расчётное время до конца выполнения цикла: " & _
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
'    If I = Int(Limit / 4) Then
'        'сообщение
'        MsgBoxEx "Активно тружусь ..." _
'    & vbCr & "Ищу столбцы, сопоставляю значения, вычисляю диапазоны", 0, "Выполнено " & _
'    Int(87 * I / Limit) & "%", 5
'    End If
'
'    If I = Int(Limit / 2) Then
'        'сообщение
'        MsgBoxEx "Всё в порядке" _
'    & vbCr & "Выполнено " & Int(87 * I / Limit) & "%", 0, "Достигли середины промежуточного цикла", 5
'    End If
'
'    If I = Int(Limit / 4 * 3) Then
'        'сообщение
'        MsgBoxEx "Обрабатывается " & iwLastRow & " строк и " & Limit & " колонок" _
'    & vbCr & "Выполнено " & Int(87 * I / Limit) & "%", 0, "Скоро...", 5
'    End If

Next i

'статус бар
Application.StatusBar = "Форматирование ячеек. Выполнено: 87 %"

'форматы
ThisWorkbook.Sheets(SheetName).Activate
Columns("Q:DD").Select
    Selection.NumberFormat = _
        "_-* #,##0.00 _?_-;-* #,##0.00 _?_-;_-* ""-""?? _?_-;_-@_-"
Columns("ED").Select
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
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(CONCATENATE(RC[-1],"" "",RC[5])=RC[-2]," _
        & "VLOOKUP(RC[-2],RC1:RC114,MATCH(R8C4,R9C1:R9C114,0),0)>0," _
        & "VLOOKUP(RC[-2],RC1:RC114,MATCH(R7C4,R9C1:R9C114,0),0)>0," _
        & "RC[4]=TRUE),"""",VLOOKUP(RC[-1],INDIRECT(CONCATENATE(""'"",VALUE(RC[5])," _
        & """ произ. календарь'!$A:$BM"")),HLOOKUP(RC[20]," _
        & "INDIRECT(CONCATENATE(""'"",VALUE(RC[5]),"" произ. календарь'!$2:$3"")),2,0),0))"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
    
'статус бар
Application.StatusBar = "Добавление формул анализа нормы часов. Выполнено: 90 %"
'анализ часов
K = 4
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
        "=OR(RC[-1]="""",OR(RC[-1]=VALUE(RC[21]),VLOOKUP(RC[-3],RC1:RC114," _
        & "MATCH(R8C4,R9C1:R9C114,0),0)>0))"
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
        & """ произ. календарь'!$A$18:$BM$31"")),HLOOKUP(RC[18]," _
        & "INDIRECT(CONCATENATE(""'"",VALUE(RC[3]),"" произ. календарь'!$18:$19"")),2,0),0))"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
    
'статус бар
Application.StatusBar = "Добавление формул анализа нормы дней. Выполнено: 92 %"
'анализ дней
K = 6
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
        "=OR(RC[-3]="""",OR(RC[-1]=VALUE(RC[18])," _
        & "VLOOKUP(RC[-5],RC1:RC114,MATCH(R8C4,R9C1:R9C114,0),0)>0))"
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
Columns("EC:EC").Select
    Selection.NumberFormat = _
        "_-* #,##0.00 _?_-;-* #,##0.00 _?_-;_-* ""-""?? _?_-;_-@_-"
Columns("EF:EF").Select
    Selection.NumberFormat = _
        "_-* #,##0.00 _?_-;-* #,##0.00 _?_-;_-* ""-""?? _?_-;_-@_-"

'статус бар
Application.StatusBar = "Выполнено: 100 %"

'завершение
ThisWorkbook.Sheets(SheetName).Activate
MsgBoxEx "Расчётная ведомость " _
    & "по компании " & vbCr & ThisWorkbook.Sheets("Preferences").Range("C7").Value2 _
    & vbCr & "за " & ThisWorkbook.Sheets(SheetName).Range("B2").Value2 & " год" _
    & vbCr & "добавлена успешно", 0, "Выполнено", 25

ThisWorkbook.Sheets("Calculation22").Activate

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
 Exit Sub
  
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub





















