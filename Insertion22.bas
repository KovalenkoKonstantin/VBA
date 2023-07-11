Attribute VB_Name = "Insertion22"
Sub Data_Insertion_22()
 
 Start = Now()
 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim ws, this As Worksheet
 Dim MyRange, MyCell As Range
 Dim SheetName, CompanyName As String
  
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 SheetName = "Processing22"
 DistinctYear = 2022
 Limit = 122 'последняя колонка базы
 begin = 12 'первый ряд вставки
 CompanyName = ThisWorkbook.Sheets("Preferences").Range("C7").Value2 'имя проекта
 
 Dim aw(1 To 122) As Variant
 Dim iw(1 To 122) As Variant
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False

FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="Выберите расчётную ведомость по компании " & CompanyName & " за " & DistinctYear & " год")
 
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
 If Range("G2").Value2 <> DistinctYear Or Range("A11").Value2 <> CompanyName Then
    Range("G2").Select
    With Selection:
        .Clear
    End With
    MsgBoxEx "Выбрана неправильная расчётная ведомость." _
    & vbCr & "Процесс прерван.", vbCritical, "Bad Day", 20
    GoTo ExitHandler
 ElseIf Range("G2").Value2 = DistinctYear Then
 Range("G2").Select
    With Selection:
        .Clear
    End With
    MsgBoxEx "Выбрана правильная расчётная ведомость." _
    & vbCr & "Продолжаем.", 0, "Succes", 5
 End If

ThisWorkbook.Sheets(SheetName).Activate
On Error Resume Next
ActiveSheet.ShowAllData

ThisWorkbook.Activate
''статус бар
'Application.StatusBar = "Выполнено: 1 %"

'определение колонок рабочей книги
On Error Resume Next
For I = 1 To 20
    If Worksheets(SheetName).Cells(I, 1) = "Сотрудник" Then
        DataRow = I
    End If
Next I

For I = 1 To Limit
    If Worksheets(SheetName).Cells(DataRow, I) = "Сотрудник" Then
        aw(1) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Месяц" Then
        aw(2) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "расчётная норма часов" Then
        aw(3) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "анализ часов" Then
        aw(4) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "расчётная норма дней" Then
        aw(5) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "анализ дней" Then
        aw(6) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Исключение всех кроме 20,26,44 счёта" Then
        aw(7) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Имя Отчество" Then
        aw(8) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Анализ изменения фамилии" Then
        aw(9) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Приказ об увольнении.Статья ТК РФ" Then
        aw(10) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Подразделение история" Then
        aw(11) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Должность" Then
        aw(12) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Вид занятости" Then
        aw(13) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Дата рождения" Then
        aw(14) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Способ отражения зарплаты в бух учете" Then
        aw(15) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "График работы" Then
        aw(16) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Норма дней" Then
        aw(17) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Норма часов" Then
        aw(18) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Отработано дней" Then
        aw(19) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Отработано часов" Then
        aw(20) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, I) = "Начислено" Then
        aw(21) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Оплата больничных листов за счет работодателя" Then
        aw(22) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Оплата больничных листов" Then
        aw(23) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Оклад по дням" Then
        aw(24) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Отсутствие по болезни (больничный еще не закрыт)" Then
        aw(25) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Надбавка за сложность и напряженность" Then
        aw(26) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Оплата отпуска по календарным дням" Then
        aw(27) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Премия за объем продаж" Then
        aw(28) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Надбавка за профессионализм и качество выполняемых работ" Then
        aw(29) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Премия по итогам года" Then
        aw(30) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Отпуск за свой счет" Then
        aw(31) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Компенсация отпуска (Отпуск основной)" Then
        aw(32) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Квартальная премия" Then
        aw(33) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Премия разовая" Then
        aw(34) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Выходное пособие при увольнении" Then
        aw(35) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Пособие по уходу за ребенком до полутора лет" Then
        aw(36) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Оклад по дням (пропорционально отработанным дням)" Then
        aw(37) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Надбавка за профессионализм и качество выполняемых работ (пропорционально отработанным дням)" Then
        aw(38) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Надбавка за работу со сведениями, составляющими государственную тайну" Then
        aw(39) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Оклад по дням 26 Нерезиденты" Then
        aw(40) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Премия не учитываемая" Then
        aw(41) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Северная надбавка" Then
        aw(42) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Районный коэффициент" Then
        aw(43) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Премия с учетом районного коэффициента(квартальная)" Then
        aw(44) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Ежегодный дополнительный оплачиваемый отпуск" Then
        aw(45) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Премия по итогам года с учетом РК" Then
        aw(46) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Вознаграждение за изобретение" Then
        aw(47) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Договор (работы, услуги)" Then
        aw(48) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Компенсация за фитнес" Then
        aw(49) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Премия с учетом районного коэффициента (месячная)" Then
        aw(50) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Прогул" Then
        aw(51) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Доплата за совмещение должностей, исполнение обязанностей" Then
        aw(52) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Месячная премия" Then
        aw(53) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Военные сборы" Then
        aw(54) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Дополнительный учебный отпуск (оплачиваемый)" Then
        aw(55) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Оплата работы в праздничные и выходные дни" Then
        aw(56) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "вознаграждение членам Совета Директоров" Then
        aw(57) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Премия Германия У.Е" Then
        aw(58) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Отсутствие по невыясненной причине" Then
        aw(59) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Отпуск без оплаты согласно ТК РФ" Then
        aw(60) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Оплата дней ухода за детьми-инвалидами" Then
        aw(61) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Пособие по уходу за ребёнком до 3 лет без оплаты" Then
        aw(62) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Оклад по часам (пропорциально отработанному времени)" Then
        aw(63) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Надбавка за сложность и напряженность (по часам пропорционально отработаннму времени)" Then
        aw(64) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Надбавка за профессионализм и качество выполняемых работ (по часам пропорц отраб времени)" Then
        aw(65) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Отсутствие по болезни" Then
        aw(66) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Доплата за работу в ночное время" Then
        aw(67) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "По итогам работы за год" Then
        aw(68) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Премия директорам (Чефранова, Басов, Набережный, Таранов)" Then
        aw(69) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Надбавка за соблюдение конфиденциальности в отношении с" Then
        aw(70) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Отпуск по беременности и родам" Then
        aw(71) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Доплата за работу в праздничные дни (дневное время)" Then
        aw(72) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Доплата за работу в праздничные дни (ночное время)" Then
        aw(73) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Доплата за переработки при суммированном учете рабочего времени" Then
        aw(74) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Неоплачиваемые дни отпуска по беременности и родам" Then
        aw(75) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Доплата за ненормированный рабочий день" Then
        aw(76) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Материальная помощь" Then
        aw(77) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, I) = "Сумма по Доход в натуральной форме" Then
        aw(78) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Питание за счет средств предприятия" Then
        aw(79) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Доход в натуральной форме (обл. НДФЛ)" Then
        aw(80) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Проездные" Then
        aw(81) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "НДФЛ к зачету в счет будущих платежей" Then
        aw(82) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Питание договорников" Then
        aw(83) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Подарок" Then
        aw(84) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Проездные Германия" Then
        aw(85) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Зачтено излишне удержанного НДФЛ" Then
        aw(86) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "сотовая связь" Then
        aw(87) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, I) = "Удержано" Then
        aw(88) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "НДФЛ" Then
        aw(89) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "НДФЛ с превышения" Then
        aw(90) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Удержано из зарплаты за оплату медстраховки" Then
        aw(91) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Удержание из з/п командировочных расходов" Then
        aw(92) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Удержание по исп. листу процентом" Then
        aw(93) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Удержание по заявлению" Then
        aw(94) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Выплата родственникам умершего сотрудника" Then
        aw(95) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Удержано из зарплаты за оплату физкультурно-оздоровительных услуг" Then
        aw(96) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Удержание за обучение из зарплаты" Then
        aw(97) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Удержание из зарплаты 73,03" Then
        aw(98) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Удержание по исп. листу фикс. суммой" Then
        aw(99) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Удержание по исполнительному документу" Then
        aw(100) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, I) = "ВЗНОСЫ" Then
        aw(101) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "ПФР до превышения" Then
        aw(102) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "ПФР с превышения" Then
        aw(103) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "ФФОМС" Then
        aw(104) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "ФСС" Then
        aw(105) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "ФСС НС" Then
        aw(106) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, I) = "% Страховых взносов" Then
        aw(107) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "База взносов" Then
        aw(108) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, I) = "Надбавка за стаж работы по защите государственной тайны" Then
        aw(109) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Оплата за дополнительный день (дни) отдыха донору" Then
        aw(110) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Командировка" Then
        aw(111) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Премия квартальная объем продаж (ПМ)" Then
        aw(112) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Премия Германия" Then
        aw(113) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Оплата по окладу (по часам)" Then
        aw(114) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Оклад по часам" Then
        aw(115) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Надбавка за сложность и напряженность (по часам)" Then
        aw(116) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Надбавка за сложность и напряженность (пропорционально отработанным дням)" Then
        aw(117) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Медицинский осмотр" Then
        aw(118) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Компенсация расходов по договорам подряда" Then
        aw(119) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Премия месячная (с учетом РК)" Then
        aw(120) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Премия квартальная" Then
        aw(121) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Премия квартальная (с учетом РК)" Then
        aw(122) = I
    End If
    
Next I
 
 importWB.Sheets(1).Activate

'определение колонок импортируемой книги
For I = 1 To 20
    If importWB.Sheets(1).Cells(I, 1) = "Организация" Then
        ImportFirstDataRow = I
    End If
Next I
For I = 1 To 20
    If importWB.Sheets(1).Cells(I, 1) = "Сотрудник" Then
        ImportSecondDataRow = I
    End If
Next I

For I = 1 To Limit
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "Сотрудник" Then '-
        iw(1) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Месяц" Then
        iw(2) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "расчётная норма часов" Then
        iw(2) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "анализ часов" Then
        iw(4) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "расчётная норма дней" Then
        iw(5) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "анализ дней" Then
        iw(6) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Исключение всех кроме 20,26,44 счёта" Then
        iw(7) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Имя Отчество" Then
        iw(8) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Анализ изменения фамилии" Then
        iw(9) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Приказ об увольнении.Статья ТК РФ" Then
        iw(10) = I
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "Подразделение история" Then '-
        iw(11) = I
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "Должность" Then '-
        iw(12) = I
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "Вид занятости" Then '-
        iw(13) = I
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "Дата рождения" Then '-
        iw(14) = I
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "Способ отражения зарплаты в бух учете" Then '-
        iw(15) = I
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "График работы" Then '-
        iw(16) = I
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "Норма дней" Then '-
        iw(17) = I
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "Норма часов" Then '-
        iw(18) = I
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "Дней" Then '-
        iw(19) = I
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "Часов" Then '-
        iw(20) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Начислено" Then
        iw(21) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Оплата больничных листов за счет работодателя" Then
        iw(22) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Оплата больничных листов" Then
        iw(23) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Оклад по дням" Then
        iw(24) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Отсутствие по болезни (больничный еще не закрыт)" Then
        iw(25) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Надбавка за сложность и напряженность" Then
        iw(26) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Оплата отпуска по календарным дням" Then
        iw(27) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Премия за объем продаж" Then
        iw(28) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Надбавка за профессионализм и качество выполняемых работ" Then
        iw(29) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Премия по итогам года" Then
        iw(30) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Отпуск за свой счет" Then
        iw(31) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Компенсация отпуска (Отпуск основной)" Then
        iw(32) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Квартальная премия" Then
        iw(33) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Премия разовая" Then
        iw(34) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Выходное пособие при увольнении" Then
        iw(35) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Пособие по уходу за ребенком до полутора лет" Then
        iw(36) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Оклад по дням (пропорционально отработанным дням)" Then
        iw(37) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Надбавка за профессионализм и качество выполняемых работ (пропорционально отработанным дням)" Then
        iw(38) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Надбавка за работу со сведениями, составляющими государственную тайну" Then
        iw(39) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Оклад по дням 26 Нерезиденты" Then
        iw(40) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Премия не учитываемая" Then
        iw(41) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Северная надбавка" Then
        iw(42) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Районный коэффициент" Then
        iw(43) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Премия с учетом районного коэффициента(квартальная)" Then
        iw(44) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Ежегодный дополнительный оплачиваемый отпуск" Then
        iw(45) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Премия по итогам года с учетом РК" Then
        iw(46) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Вознаграждение за изобретение" Then
        iw(47) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Договор (работы, услуги)" Then
        iw(48) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Компенсация за фитнес" Then
        iw(49) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Премия с учетом районного коэффициента (месячная)" Then
        iw(50) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Прогул" Then
        iw(51) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Доплата за совмещение должностей, исполнение обязанностей" Then
        iw(52) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Месячная премия" Then
        iw(53) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Военные сборы" Then
        iw(54) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Дополнительный учебный отпуск (оплачиваемый)" Then
        iw(55) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Оплата работы в праздничные и выходные дни" Then
        iw(56) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "вознаграждение членам Совета Директоров" Then
        iw(57) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Премия Германия У.Е" Then
        iw(58) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Отсутствие по невыясненной причине" Then
        iw(59) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Отпуск без оплаты согласно ТК РФ" Then
        iw(60) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Оплата дней ухода за детьми-инвалидами" Then
        iw(61) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Пособие по уходу за ребёнком до 3 лет без оплаты" Then
        iw(62) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Оклад по часам (пропорциально отработанному времени)" Then
        iw(63) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Надбавка за сложность и напряженность (по часам пропорционально отработаннму времени)" Then
        iw(64) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Надбавка за профессионализм и качество выполняемых работ (по часам пропорц отраб времени)" Then
        iw(65) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Отсутствие по болезни" Then
        iw(66) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Доплата за работу в ночное время" Then
        iw(67) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "По итогам работы за год" Then
        iw(68) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Премия директорам (Чефранова, Басов, Набережный, Таранов)" Then
        iw(69) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Надбавка за соблюдение конфиденциальности в отношении с" Then
        iw(70) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Отпуск по беременности и родам" Then
        iw(71) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Доплата за работу в праздничные дни (дневное время)" Then
        iw(72) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Доплата за работу в праздничные дни (ночное время)" Then
        iw(73) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Доплата за переработки при суммированном учете рабочего времени" Then
        iw(74) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Неоплачиваемые дни отпуска по беременности и родам" Then
        iw(75) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Доплата за ненормированный рабочий день" Then
        iw(76) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Материальная помощь" Then
        iw(77) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Сумма по Доход в натуральной форме" Then
        iw(78) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Питание за счет средств предприятия" Then
        iw(79) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Доход в натуральной форме (обл. НДФЛ)" Then
        iw(80) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Проездные" Then
        iw(81) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "НДФЛ к зачету в счет будущих платежей" Then
        iw(82) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Питание договорников" Then
        iw(83) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Подарок" Then
        iw(84) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Проездные Германия" Then
        iw(85) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Зачтено излишне удержанного НДФЛ" Then
        iw(86) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "сотовая связь" Then
        iw(87) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Удержано" Then
        iw(88) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "НДФЛ" Then
        iw(89) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "НДФЛ с превышения" Then
        iw(90) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Удержано из зарплаты за оплату медстраховки" Then
        iw(91) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Удержание из з/п командировочных расходов" Then
        iw(92) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Удержание по исп. листу процентом" Then
        iw(93) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Удержание по заявлению" Then
        iw(94) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Выплата родственникам умершего сотрудника" Then
        iw(95) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Удержано из зарплаты за оплату физкультурно-оздоровительных услуг" Then
        iw(96) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Удержание за обучение из зарплаты" Then
        iw(97) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Удержание из зарплаты 73,03" Then
        iw(98) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Удержание по исп. листу фикс. суммой" Then
        iw(99) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Удержание по исполнительному документу" Then
        iw(100) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "ВЗНОСЫ" Then
        iw(101) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "ПФР до превышения" Then
        iw(102) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "ПФР с превышения" Then
        iw(103) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "ФФОМС" Then
        iw(104) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "ФСС" Then
        iw(105) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "ФСС НС" Then
        iw(106) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "% Страховых взносов" Then
        iw(107) = I
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, I) = "База взносов" Then '-
        iw(108) = I
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Надбавка за стаж работы по защите государственной тайны" Then '-
        iw(109) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Оплата за дополнительный день (дни) отдыха донору" Then '-
        iw(110) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Командировка" Then '-
        iw(111) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Премия квартальная объем продаж (ПМ)" Then '-
        iw(112) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Премия Германия" Then '-
        iw(113) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Оплата по окладу (по часам)" Then '-
        iw(114) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Оклад по часам" Then '-
        iw(115) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Надбавка за сложность и напряженность (по часам)" Then '-
        iw(116) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Надбавка за сложность и напряженность (пропорционально отработанным дням)" Then '-
        iw(117) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Медицинский осмотр" Then '-
        iw(118) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Компенсация расходов по договорам подряда" Then '-
        iw(119) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Премия месячная (с учетом РК)" Then '-
        iw(120) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Премия квартальная" Then '-
        iw(121) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Премия квартальная (с учетом РК)" Then '-
        iw(122) = I
    End If

Next I

'удаление предыдущих данных
 ThisWorkbook.Sheets(SheetName).Activate
awLastRow = Cells(Rows.Count, "A").End(xlUp).row
awLastCol = Cells(begin, Columns.Count).End(xlUp).column
Range(Cells(begin, 1), Cells(awLastRow, Limit)).Select
 With Selection
        .Clear
 End With

'  'сообщение
' MsgBoxEx "Выполнено 5%", 0, "5%. Мы только начали...", 5

 'вставка данных
 importWB.Sheets(1).Activate
 iwLastRow = Cells(Rows.Count, "A").End(xlUp).row

For I = 1 To Limit
'статус бар
Application.StatusBar = "Промежуточный цикл. Выполнено: " & Int(100 * I / Limit) & "%." & _
" Общий прогресс: " & Int(90 * I / Limit) & "%"
 importWB.Activate
 Range(Cells(begin - 1, iw(I)), Cells(iwLastRow, iw(I))).Copy

 ThisWorkbook.Sheets(SheetName).Activate
 Range(Cells(begin, aw(I)), Cells(iwLastRow, aw(I))).Select
    With Selection
           .PasteSpecial Paste:=xlPasteAll
           .UnMerge
           .Font.Name = "Times New Roman"
           .WrapText = False
           .MergeCells = False
           .Font.Size = 8
    End With
    If I = Int(Limit / 4) Then
        'сообщение
        MsgBoxEx "Активно тружусь ..." _
    & vbCr & "Ищу столбцы, сопоставляю значения, вычисляю диапазоны", 0, "Выполнено " & _
    Int(90 * I / Limit) & "%", 5
    End If

    If I = Int(Limit / 2) Then
        'сообщение
        MsgBoxEx "Всё в порядке" _
    & vbCr & "Выполнено " & Int(90 * I / Limit) & "%", 0, "Достигли середины", 5
    End If

    If I = Int(Limit / 4 * 3) Then
        'сообщение
        MsgBoxEx "Обрабатывается " & iwLastRow & " строк и " & Limit & " колонок" _
    & vbCr & "Выполнено " & Int(90 * I / Limit) & "%", 0, "Скоро...", 5
    End If

Next I

'сообщение
MsgBoxEx "Ускоряемся ..." _
& vbCr & "Выполнено 90%", 0, "Нет, я не завис.", 5

'статус бар
Application.StatusBar = "Выполнено: 90 %"

'форматы
ThisWorkbook.Sheets(SheetName).Activate
Columns("Q:DD").Select
    Selection.NumberFormat = _
        "_-* #,##0.00 _?_-;-* #,##0.00 _?_-;_-* ""-""?? _?_-;_-@_-"

'вставка проверочных формул
ThisWorkbook.Sheets(SheetName).Activate
'месяц
Cells(begin, aw(2)).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(IFERROR(SEARCH("" 20"",RC[-1],1)>0,FALSE),TRIM(MID(IF(IFERROR(SEARCH("" 20"",RC[-1],1)>0,FALSE),RC[-1],R[-1]C),1,SEARCH("" "",IF(IFERROR(SEARCH("" 20"",RC[-1],1)>0,FALSE),RC[-1],R[-1]C),1)-1)),R[-1]C)"
    Cells(begin, aw(2)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(2)), Cells(iwLastRow, aw(2)))

'расчётная норма часов
K = 3
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
       "=IF(OR(CONCATENATE(RC[-1],"" " & DistinctYear & """)=RC[-2]," _
       & "VLOOKUP(RC[-2],RC1:RC108,MATCH(R8C4,R9C1:R9C108,0),0)>0,VLOOKUP(RC[-2]," _
       & "RC1:RC108,MATCH(R7C4,R9C1:R9C108,0),0)>0,RC[4]=TRUE),""""," _
       & "VLOOKUP(RC[-1],'" & DistinctYear & " произ. календарь'!C1:C65,HLOOKUP(RC[13]," _
       & "'" & DistinctYear & " произ. календарь'!R2:R3,2,0),0))"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))

'анализ часов
K = 4
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
        "=OR(RC[-1]="""",OR(RC[-1]=VALUE(RC[14]),VLOOKUP(RC[-3],RC1:RC108," _
        & "MATCH(R8C4,R9C1:R9C108,0),0)>0))"
    Range("A1").Copy
    Cells(begin, aw(K)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))

'статус бар
Application.StatusBar = "Выполнено: 93 %"

'расчётная норма дней
K = 5
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(CONCATENATE(RC[-3],"" " & DistinctYear & """)=RC[-4],VLOOKUP(RC[-4]," _
        & "RC1:RC108,MATCH(R8C4,R9C1:R9C108,0),0)>0,RC[2]=TRUE),"""",VLOOKUP(RC[-3]," _
        & "'" & DistinctYear & " произ. календарь'!R18C1:R31C65,HLOOKUP(RC[11]," _
        & "'" & DistinctYear & " произ. календарь'!R18:R19,2,0),0))"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))

'анализ дней
K = 6
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
        "=OR(RC[-3]="""",OR(RC[-1]=VALUE(RC[11]),VLOOKUP(RC[-5],RC1:RC108,MATCH(R8C4,R9C1:R9C108,0),0)>0))"
    Range("A1").Copy
    Cells(begin, aw(K)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))

'Исключение всех кроме 20,26,44 счёта
K = 7
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
       "=NOT(OR(IFERROR(SEARCH(20,RC[8],1),FALSE),IFERROR(SEARCH(26,RC[8],1),FALSE),IFERROR(SEARCH(44,RC[8],1),FALSE)))"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))

'Имя Отчество
K = 8
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
       "=IF(CONCATENATE(RC[-6],"" " & DistinctYear & """)=RC[-7],"""",(MID(RC[-7],SEARCH("" "",RC[-7],1),LEN(RC[-7]))))"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))

'Анализ изменения фамилии
K = 9
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-8],RC[-7],RC[5])"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
Columns("I:I").Select
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
Application.StatusBar = "Выполнено: 96 %"

''Умершие
'K = 9
'Cells(begin, aw(K)).Select
'    ActiveCell.FormulaR1C1 = "=IF(AND(RC[89]>0,RC[95]=""""),""пИчаль"","""")"
'    Cells(begin, aw(K)).Select
'    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))

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
Application.StatusBar = "Выполнено: 97 %"

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

''сообщение
'MsgBoxEx "Почти всё готово" _
'& vbCr & "Выполнено 98%", 0, "98%", 5

'статус бар
Application.StatusBar = "Выполнено: 100 %"

'завершение
ThisWorkbook.Sheets(SheetName).Activate
MsgBoxEx "Расчётная ведомость " _
    & "по компании " & vbCr & ThisWorkbook.Sheets("Preferences").Range("C7").Value2 _
    & vbCr & "за " & ThisWorkbook.Sheets(SheetName).Range("B2").Value2 & " год" _
    & vbCr & "добавлена успешно", 0, "Выполнено", 25

ThisWorkbook.Sheets("Calculation22").Activate
'If Range("E1").Value2 = True Then
'    MsgBox ("Ошибок нет")
'Finish = (Now() - Start) * 24 * 60 * 60
'MsgBox (Finish)
'ElseIf Range("E1").Value2 = False Then
'    result = MsgBox("Расчётная ведомость загружена по компании" _
'    & vbCr & ThisWorkbook.Sheets("Calculation22").Range("E2").Value2 _
'    & vbCr & "Отчёт по среднесписочной численности сотрудников загружен по компании" _
'    & vbCr & ThisWorkbook.Sheets("ССЧ22").Range("AG5").Value2 _
'    & vbCr & "Загрузить корректный отчёт по средней списочности компании" _
'    & vbCr & ThisWorkbook.Sheets("Calculation22").Range("E2").Value2 _
'    & "?", vbYesNo)
'    If result = vbYes Then
'        Application.Run "Data_insertion_SS4"
'    Else
'        MsgBox "Действие отменено!" _
'    & vbCr & "Выберите корректный отчёт с расчётной ведомостью по компании " _
'    & vbCr & ThisWorkbook.Sheets("ССЧ22").Range("AG5").Value2
'    End If
'    GoTo ExitHandler2
'End If

ExitHandler:
    On Error Resume Next
    importWB.Close
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
 ThisWorkbook.Sheets("Preferences").Activate
 Exit Sub
  
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler
End Sub



















