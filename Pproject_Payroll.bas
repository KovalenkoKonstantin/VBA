Attribute VB_Name = "Pproject_Payroll"
Sub Project_Payroll_Insertion()
 Start = Now()
 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim ws, this As Worksheet
 Dim MyRange, MyCell As Range
 Dim SheetName, ProjectName As String
 Dim K As Variant 'ключ из 11 строки

 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 SheetName = "РВ_Проекта"
 Limit = 135 'последняя колонка базы
 begin = 12 'первый ряд вставки
' LimitIW = 170 'последняя колонка импортируемой книги
ProjectName = ThisWorkbook.Sheets("Preferences").Range("C13").Value2 'имя проекта

 
 Dim aw(1 To 135) As Variant
 Dim iw(1 To 135) As Variant
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual

'статус бар
Application.StatusBar = "Определение колонок рабочей книги"

'определение колонок рабочей книги
On Error Resume Next
For i = 1 To 20
    If Worksheets(SheetName).Cells(i, 1) = "Сотрудник" Then
        DataRow = i
    End If
Next i

For i = 1 To Limit
'статус бар
Application.StatusBar = "Определение колонок рабочей книги." & _
" Выполнено: " & Int(90 * i / Limit) & "%"

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
    If Worksheets(SheetName).Cells(DataRow, i) = "Год" Then
        aw(9) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Анализ изменения фамилии" Then
        aw(10) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Проект" Then
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
    If Worksheets(SheetName).Cells(DataRow, i) = "% процент страховых взносов" Then
        aw(107) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "База взносов на проекте" Then
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
    If Worksheets(SheetName).Cells(DataRow, i) = "Отработано дней на проекте" Then
        aw(113) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Отработано часов на проекте" Then
        aw(114) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Сокращённое Ф.И.О." Then
        aw(115) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Оклад" Then
        aw(116) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Надбавка" Then
        aw(117) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия" Then
        aw(118) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Основная заработная плата" Then
        aw(119) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Месяц числом" Then
        aw(120) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Временное решение" Then
        aw(121) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Таб.№" Then
        aw(122) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "ФОТ" Then
        aw(123) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Анализ расхождений по зарплате" Then
        aw(124) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия квартальная" Then
        aw(125) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия квартальная (с учетом РК)" Then
        aw(126) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия месячная (с учетом РК)" Then
        aw(127) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия по итогам года (с учетом РК)" Then
        aw(128) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "ФОТ на проекте" Then
        aw(129) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Надбавка за сложность и напряженность (по часам пропорц. отработанному времени)" Then
        aw(130) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия разовая (с учетом РК)" Then
        aw(131) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия месячная" Then
        aw(132) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Доплата за работу в ночное время (праздничные и выходные дни)" Then
        aw(133) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия полугодовая (с учетом РК)" Then
        aw(134) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Премия полугодовая" Then
        aw(135) = i
    End If
    
Next i

 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="Выберите расчётную ведомость ПРОЕКТА " & ProjectName)

 If TypeName(FilesToOpen) = "Boolean" Then
 GoTo ExitHandler
 End If

ThisWorkbook.Sheets(SheetName).Activate
On Error Resume Next
ActiveSheet.ShowAllData

'вставка листов
Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

On Error Resume Next

 importWB.Sheets(1).Activate
 
 'статус бар
Application.StatusBar = "Определение колонок импортируемой книги"

'определение колонок импортируемой книги
For i = 1 To 20
    If importWB.Sheets(1).Cells(i, 1) = "Организация" Then
        ImportFirstDataRow = i
    End If
Next i
For i = 1 To 20
    If importWB.Sheets(1).Cells(i, 1) = "ФИОСотрудник" Then
        ImportSecondDataRow = i
    End If
Next i

'определение последней колонки в импортируемой книге
LimitIW = Cells(ImportSecondDataRow, Columns.Count).End(xlToLeft).column

For i = 1 To LimitIW
'статус бар
Application.StatusBar = "Определение колонок импортируемой книги" & _
" Выполнено: " & Int(90 * i / LimitIW) & "%"

    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "ФИОСотрудник" Then '-
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
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Год" Then
        iw(9) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Анализ изменения фамилии" Then
        iw(10) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Проект" Then '-
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
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Отработано дней" Then '-
        iw(19) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Отработано часов" Then '-
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
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "% процент страховых взносов" Then
        iw(107) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "База взносов на проекте" Then '-
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
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Отработано дней на проекте" Then '-
        iw(113) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Отработано часов на проекте" Then '-
        iw(114) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Сокращённое Ф.И.О." Then '-
        iw(115) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Оклад" Then '-
        iw(116) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Надбавка" Then '-
        iw(117) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Премия" Then '-
        iw(118) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Основная заработная плата" Then '-
        iw(119) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Месяц числом" Then '-
        iw(120) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Временное решение" Then '-
        iw(121) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Таб.№" Then '-
        iw(122) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "ФОТ" Then '-
        iw(123) = i
    End If
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "Анализ расхождений по зарплате" Then '-
        iw(124) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия квартальная" Then
        iw(125) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия квартальная (с учетом РК)" Then
        iw(126) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия месячная (с учетом РК)" Then
        iw(127) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия по итогам года (с учетом РК)" Then
        iw(128) = i
    End If
' ----------------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportSecondDataRow, i) = "ФОТ на проекте" Then
        iw(129) = i
    End If
' ----------------------------------------------------------------------------------------------------

    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Надбавка за сложность и напряженность (по часам пропорц. отработанному времени)" Then
        iw(130) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия разовая (с учетом РК)" Then
        iw(131) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия месячная" Then
        iw(132) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Доплата за работу в ночное время (праздничные и выходные дни)" Then
        iw(133) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия полугодовая (с учетом РК)" Then
        iw(134) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Премия полугодовая" Then
        iw(135) = i
    End If
    
Next i

'статус бар
Application.StatusBar = "Удаление предыдущих данных"

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
Application.StatusBar = "Вставка данных." & _
" Выполнено: " & Int(90 * i / Limit) & "%"

 importWB.Activate
 Range(Cells(begin, iw(i)), Cells(iwLastRow, iw(i))).Copy
 
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
Application.StatusBar = "Форматирование диапазонов"

'форматы
ThisWorkbook.Sheets(SheetName).Activate
Columns("Q:DD").Select
    Selection.NumberFormat = _
        "_-* #,##0.00 _?_-;-* #,##0.00 _?_-;_-* ""-""?? _?_-;_-@_-"
        
'статус бар
Application.StatusBar = "Добавление проверочных формул"

'вставка проверочных формул
ThisWorkbook.Sheets(SheetName).Activate

'статус бар
Application.StatusBar = "Добавление проверочных формул нормы месяца"

'месяц
Cells(begin, aw(2)).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(IFERROR(SEARCH("" 20"",RC[-1],1)>0,FALSE),TRIM(MID(IF(IFERROR(SEARCH("" 20"",RC[-1],1)>0,FALSE),RC[-1],R[-1]C),1,SEARCH("" "",IF(IFERROR(SEARCH("" 20"",RC[-1],1)>0,FALSE),RC[-1],R[-1]C),1)-1)),R[-1]C)"
    Cells(begin, aw(2)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(2)), Cells(iwLastRow - 1, aw(2)))
    
'расчётная норма часов
'статус бар
Application.StatusBar = "Добавление проверочных формул нормы часов"
K = 3
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(CONCATENATE(RC[-1],"" "",RC[7])=RC[-2]," & _
        "VLOOKUP(RC[-2],RC1:RC108,MATCH(R8C4,R9C1:R9C108,0),0)>0," & _
        "RC[4]=TRUE),"""",IF(RC[7]=2021,VLOOKUP(RC[-1],'2021 произ. календарь'!C1:C65," & _
        "HLOOKUP(RC[13],'2021 произ. календарь'!R2:R3,2,0),0)," & _
        "IF(RC[7]=2022,VLOOKUP(RC[-1],'2022 произ. календарь'!C1:C65,HLOOKUP(RC[13]," & _
        "'2022 произ. календарь'!R2:R3,2,0),0),IF(RC[7]=2023,VLOOKUP(R" & _
        "C[-1],'2023 произ. календарь'!C1:C65,HLOOKUP(RC[13]," & _
        "'2023 произ. календарь'!R2:R3,2,0),0)," & _
        """за пределами производственных календарей""))))" & _
        ""
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
    
'анализ часов
'статус бар
Application.StatusBar = "Добавление проверочных формул анализа расхождений между часами"
K = 4
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
        "=OR(RC[-1]="""",OR(RC[-1]=VALUE(RC[14]),VLOOKUP(RC[-3],RC1:RC108,MATCH(R8C4,R9C1:R9C108,0),0)>0))"
    Range("A1").Copy
    Cells(begin, aw(K)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
    
'Анализ расхождений по зарплате
'статус бар
Application.StatusBar = "Добавление проверочных формул анализа расхождений по заработной плате"
K = 124
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-2]="""",TRUE,RC[-3]=RC[-1])"
    Range("A1").Copy
    Cells(begin, aw(K)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
    
'расчётная норма дней
'статус бар
Application.StatusBar = "Добавление проверочных формул нормы дней"
K = 5
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(CONCATENATE(RC[-3],"" "",RC[5])=RC[-4]," & _
        "VLOOKUP(RC[-4],RC1:RC108,MATCH(R8C4,R9C1:R9C108,0),0)>0," & _
        "RC[2]=TRUE),"""",IF(RC[5]=2021," & _
        "VLOOKUP(RC[-3],'2021 произ. календарь'!R18C1:R31C65," & _
        "HLOOKUP(RC[11],'2021 произ. календарь'!R18:R19,2,0),0),IF(RC[5]=2022," & _
        "VLOOKUP(RC[-3],'2022 произ. календарь'!R18C1:R31C65,HLOOKUP(RC[11]," & _
        "'2022 произ. календарь'!R18:R19,2,0),0),IF(RC[5" & _
        "]=2023,VLOOKUP(RC[-3],'2023 произ. календарь'!R18C1:R31C65,HLOOKUP(RC[11]," & _
        "'2023 произ. календарь'!R18:R19,2,0),0)," & _
        """за пределами производственных календарей""))))" & _
        ""
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
    
'анализ дней
'статус бар
Application.StatusBar = "Добавление проверочных формул анализа расхождений по дням"
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
'статус бар
Application.StatusBar = "Добавление формул включения в расчёт 20,46,44 счетов"
K = 7
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
       "=NOT(OR(IFERROR(SEARCH(20,RC[8],1),FALSE),IFERROR(SEARCH(26,RC[8],1),FALSE),IFERROR(SEARCH(44,RC[8],1),FALSE)))"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
    
'Имя Отчество
'статус бар
Application.StatusBar = "Добавление проверочных формул анализа имени и отчества"
K = 8
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
       "=IF(CONCATENATE(RC[-6],"" 2022"")=RC[-7],"""",(MID(RC[-7],SEARCH("" "",RC[-7],1),LEN(RC[-7]))))"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
    
'Анализ изменения фамилии
'статус бар
Application.StatusBar = "Добавление проверочных формул анализа изменения фамилии"
K = 10
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
    
'Год
'статус бар
Application.StatusBar = "Добавление формул выделения года"
K = 9
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = "=VALUE(IF(IFERROR(SEARCH("" 20"",RC[-9],1)>0,FALSE),MID(RC[-9],SEARCH("" "",RC[-9],1)+1,4),R[-1]C))"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))

'Сокращённое Ф.И.О.
'статус бар
Application.StatusBar = "Добавление формул сокращения Ф.И.О. для анализа табеля рабочего времени"
K = 115
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = "=IFERROR(CONCATENATE(MID(RC[-18],1,FIND("" "",RC[-18])+1),"". "",MID(RC[-18],FIND("" "",RC[-18],FIND("" "",RC[-18])+1)+1,1),"".""),"""")"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
    
'Основная заработная плата
'статус бар
Application.StatusBar = "Добавление формул вычисления заработной платы"
K = 119
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = "=ROUND(RC[10]/(RC[7]/RC[-2]),0)" & _
    "+ROUND(RC[11]/(RC[7]/RC[-2]),0)"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))

'Месяц числом
'статус бар
Application.StatusBar = "Добавление формулы преобразования месяца в число"
K = 120
Cells(begin, aw(K)).Select
'    ActiveCell.FormulaR1C1 = "=SUM(RC[5]:RC[6])"
    ActiveCell.FormulaR1C1 = "=IFERROR(MONTH(DATEVALUE(RC[-19]&""1"")),"""")"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
    
'Временное решение
'статус бар
Application.StatusBar = "Добавление формул переноса ФОТ из расчётной ведомости"
K = 121
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = "=SUMIFS(РВ!C4,РВ!C2,RC[-21],РВ!C[-6],RC[-20],РВ!C11,RC[-12])"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))

'Сумма по Доход в натуральной форме
'статус бар
Application.StatusBar = "Добавление формул по доходу в натуральном виде"
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
Application.StatusBar = "Завершение"

'завершение
importWB.Close
ThisWorkbook.Sheets(SheetName).Activate
If Range("DZ4").Value2 = True Then
MsgBoxEx "Расчётная ведомость " _
    & "по компании " & vbCr & ThisWorkbook.Sheets("РВ_Проекта").Range("A12").Value2 _
    & vbCr & "за " & ThisWorkbook.Sheets("РВ_Проекта").Range("B2").Value2 & " - " _
    & ThisWorkbook.Sheets("РВ_Проекта").Range("B3").Value2 & " года" _
    & vbCr & "по проекту:" _
    & vbCr & ThisWorkbook.Sheets("РВ_Проекта").Range("C2").Value2 _
    & vbCr & "добавлена успешно.", 0, "Microsoft Excel", 15
End If

ThisWorkbook.Sheets(SheetName).Activate
If Range("DZ4").Value2 = False Then
    MsgBoxEx "Проверки не пройдены", vbCritical, "Ошибка", 15
End If

ExitHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    Application.Calculation = xlAutomatic
 ThisWorkbook.Sheets("Preferences").Activate
 Exit Sub
 
'ExitHandler2:
'    Application.ScreenUpdating = True
'    Application.EnableEvents = True
'    ActiveSheet.DisplayPageBreaks = True
'    Application.DisplayStatusBar = True
'    Application.DisplayAlerts = True
' ThisWorkbook.Sheets("Preferences").Activate
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler
End Sub




