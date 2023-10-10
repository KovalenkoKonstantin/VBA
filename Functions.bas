Attribute VB_Name = "Functions"
Public Function РЕГРЕСС(Salary As Double, Percentage As Double, month As Integer, Limit As Double) As Double
Dim result, FCC, PF, FFOMC, Tr, Extra As Double

'вариант 7,8%
If Percentage = 0.078 Then
    FCC = 0.015
    PF = 0.06
    FFOMC = 0.001
    Tr = 0.002
    Extra = 0
End If

'вариант 30,2%
If Percentage = 0.302 Then
    FCC = 0.029
    PF = 0.22
    FFOMC = 0.051
    Tr = 0.002
    Extra = 0.1
End If

'вариант 14,2%
If Percentage = 0.142 Then
    FCC = 0.02
    PF = 0.08
    FFOMC = 0.04
    Tr = 0.002
    Extra = 0
End If

'main
temp = 0
For i = 1 To month
    temp = temp + Salary
    result = Salary * FFOMC + Salary * Tr
    If temp <= Limit Then
        result = result + Salary * FCC + Salary * PF
    ElseIf Limit - temp + Salary > 0 Then
        result = result + (Limit - temp + Salary) * FCC + (Limit - temp + Salary) * PF + (temp - Limit) * Extra
    Else
        result = result + Salary * Extra
    End If
    result = result / Salary
Next i

РЕГРЕСС = result

End Function

Public Sub DescriptForРегресс()
    Application.MacroOptions "РЕГРЕСС", Description:="РЕГРЕСС", ArgumentDescriptions:=Array("Оклад сотрудника", "Предельный процент социальных выплат", "Месяц(как число)", "Предельное значение отчислений согласно законодательству")
End Sub


Function СУММАПРОПИСЬЮ(n As Double) As String
 
 Dim Nums1, Nums2, Nums3, Nums4 As Variant
 
 Nums1 = Array("", "один ", "два ", "три ", "четыре ", "пять ", "шесть ", "семь ", "восемь ", "девять ")
 Nums2 = Array("", "десять ", "двадцать ", "тридцать ", "сорок ", "пятьдесят ", "шестьдесят ", "семьдесят ", _
                        "восемьдесят ", "девяносто ")
 Nums3 = Array("", "сто ", "двести ", "триста ", "четыреста ", "пятьсот ", "шестьсот ", "семьсот ", _
                        "восемьсот ", "девятьсот ")
 Nums4 = Array("", "одна ", "две ", "три ", "четыре ", "пять ", "шесть ", "семь ", "восемь ", "девять ")
 Nums5 = Array("десять ", "одиннадцать ", "двенадцать ", "тринадцать ", "четырнадцать ", _
                        "пятнадцать ", "шестнадцать ", "семнадцать ", "восемнадцать ", "девятнадцать ")
 
 If n <= 0 Then
   СУММАПРОПИСЬЮ = "ноль"
   Exit Function
 End If
 'разделяем число на разряды, используя вспомогательную функцию Class
 ed = Class(n, 1)
 dec = Class(n, 2)
 sot = Class(n, 3)
 tys = Class(n, 4)
 dectys = Class(n, 5)
 sottys = Class(n, 6)
 mil = Class(n, 7)
 decmil = Class(n, 8)
 
 'проверяем миллионы
 Select Case decmil
   Case 1
     mil_txt = Nums5(mil) & "миллионов "
     GoTo www
   Case 2 To 9
     decmil_txt = Nums2(decmil)
 End Select
 Select Case mil
   Case 1
     mil_txt = Nums1(mil) & "миллион "
   Case 2, 3, 4
     mil_txt = Nums1(mil) & "миллиона "
   Case 5 To 20
     mil_txt = Nums1(mil) & "миллионов "
 End Select
www:
 sottys_txt = Nums3(sottys)
 'проверяем тысячи
 Select Case dectys
   Case 1
     tys_txt = Nums5(tys) & "тысяч "
     GoTo eee
   Case 2 To 9
     dectys_txt = Nums2(dectys)
 End Select
 Select Case tys
   Case 0
     If dectys > 0 Then tys_txt = Nums4(tys) & "тысяч "
   Case 1
     tys_txt = Nums4(tys) & "тысяча "
   Case 2, 3, 4
     tys_txt = Nums4(tys) & "тысячи "
   Case 5 To 9
     tys_txt = Nums4(tys) & "тысяч "
 End Select
 If dectys = 0 And tys = 0 And sottys <> 0 Then sottys_txt = sottys_txt & "тысяч "
eee:
 sot_txt = Nums3(sot)
 'проверяем десятки
 Select Case dec
   Case 1
     ed_txt = Nums5(ed)
     GoTo rrr
   Case 2 To 9
     dec_txt = Nums2(dec)
 End Select
 
 ed_txt = Nums1(ed)
rrr:
 'формируем итоговую строку
 СУММАПРОПИСЬЮ = UCase(Left(decmil_txt & mil_txt & sottys_txt & dectys_txt & tys_txt & sot_txt & dec_txt & ed_txt, 1)) & LCase(Mid(decmil_txt & mil_txt & sottys_txt & dectys_txt & tys_txt & sot_txt & dec_txt & ed_txt, 2))
End Function
 
'вспомогательная функция для выделения из числа разрядов
Private Function Class(M, i)
  Class = Int(Int(M - (10 ^ i) * Int(M / (10 ^ i))) / 10 ^ (i - 1))
End Function

Public Function CountByColor(DataRange As Range, ColorSample As Range) As Long
    Dim cell As Range, n As Long
     
    For Each cell In DataRange
        If cell.Font.Color = ColorSample.Font.Color Then n = n + 1
    Next cell
    CountByColor = n
End Function

Function SumByColor(DataRange As Range, ColorSample As Range) As Double
    Dim cell As Range, total As Double
     
    For Each cell In DataRange
        If IsNumeric(cell) And cell.Interior.Color = ColorSample.Interior.Color Then total = total + cell.Value
    Next cell
    SumByColor = total
End Function

Function AverageByColor(DataRange As Range, ColorSample As Range) As Double
    Dim cell As Range, total As Double, n As Long
     
    For Each cell In DataRange
        If IsNumeric(cell) And cell.Interior.Color = ColorSample.Interior.Color Then
            total = total + cell.Value
            n = n + 1
        End If
    Next cell
    AverageByColor = total / n
End Function
Public Function MsgBoxEx(Prompt, Optional Buttons As VbMsgBoxStyle = 0, Optional Title, Optional TimeOut = 0) As VbMsgBoxResult
'---------------------------------------------------------------------------------------
' Procedure : MsgBoxEx
' Purpose   : MsgBox с таймаутом (используется метод Popup WScript.Shell)
'             Создаёт .VBS-файл во временной папке, запускает его, возвращает коды результата, удаляет временный файл
' Arguments : Первые 3 аргумента такие же, как у MsgBox, 4-й - таймаут в секундах.
'             Если 4-й аргумент не задан или <=0, то ожидает ввода пользователя как обычный MsgBox
' Ret.Value : Такие же, как у Msgbox, но возвращает -1 по истечении таймаута.
' Errors    : Возвращает ошибку 735 - "Can't save file to TEMP" если временная папка не доступна
' Author    : Казанский, [email]exceleved@yandex.ru[/email]
' URL       : http://www.cyberforum.ru/post5874942.html
' Date      : 09.03.2014
'---------------------------------------------------------------------------------------
   Dim sTmp$, ff%
   With CreateObject("WScript.Shell")
      sTmp = Environ("temp")
      If sTmp = "" Then
         sTmp = Environ("tmp")
         If sTmp = "" Then
            sTmp = .SpecialFolders("MyDocuments")
            If sTmp = "" Then Err.Raise 735, "MsgBoxEx", "Can't save file to TEMP"
         End If
      End If
      sTmp = sTmp & Format$(Now, """\~MsgBoxEx""YYYYMMDDHHMMSS"".vbs""")   'уникальное имя файла
      ff = FreeFile
      If IsMissing(Title) Then Title = ""
      Prompt = Str2Code(Prompt): TimeOut = Int(TimeOut): Title = Str2Code(Title): Buttons = Int(Buttons)
      Open sTmp For Output As ff
      Print #ff, "WScript.Quit CreateObject(""WScript.Shell"").Popup (""" & Prompt & """, " & TimeOut & ", """ & Title & """, " & Buttons & ")" ' Popup(<Text>,<SecondsToWait>,<Title>,<Type>) ' http://www.script-coding.com/WSH/WshShell.html#3.2.
      Close #ff
      MsgBoxEx = .Run(sTmp, 0, True) ' Run(<Command>,<WindowStyle>,<WaitOnReturn>) ' http://www.script-coding.com/WSH/WshShell.html#3.4.
   End With
   On Error Resume Next
   Kill sTmp
End Function
Private Function Str2Code$(sTxt) ' заменить CR+LF, LF+CR, CR, LF на " & vblf & " для использования в VBS
   Str2Code = Replace(Replace(Replace(Replace(Replace(sTxt, """", """"""), vbCrLf, vbLf), vbLf & vbCr, vbLf), vbCr, vbLf), vbLf, """ & vblf & """)
End Function

Function ЧислоПропись(Число As Currency) As String

'до 999 999 999 999

On Error GoTo Число_Error

Dim strМиллиарды As String, strМиллионы As String, strТысячи As String, strЕдиницы As String, strСотые As String

Dim Поз As Integer

 

strЧисло = Format(Int(Число), "000000000000")

 

'Миллиарды'

Поз = 1

strМиллиарды = Сотни(Mid(strЧисло, Поз, 1))

strМиллиарды = strМиллиарды & Десятки(Mid(strЧисло, Поз + 1, 2), "м")

strМиллиарды = strМиллиарды & ИмяРазряда(strМиллиарды, Mid(strЧисло, Поз + 1, 2), "миллиард ", "миллиарда ", "миллиардов ")

 

'Миллионы'

Поз = 4

strМиллионы = Сотни(Mid(strЧисло, Поз, 1))

strМиллионы = strМиллионы & Десятки(Mid(strЧисло, Поз + 1, 2), "м")

strМиллионы = strМиллионы & ИмяРазряда(strМиллионы, Mid(strЧисло, Поз + 1, 2), "миллион ", "миллиона ", "миллионов ")

 

'Тысячи'

Поз = 7

strТысячи = Сотни(Mid(strЧисло, Поз, 1))

strТысячи = strТысячи & Десятки(Mid(strЧисло, Поз + 1, 2), "ж")

strТысячи = strТысячи & ИмяРазряда(strТысячи, Mid(strЧисло, Поз + 1, 2), "тысяча ", "тысячи ", "тысяч ")

 

'Единицы'

Поз = 10

strЕдиницы = Сотни(Mid(strЧисло, Поз, 1))

strЕдиницы = strЕдиницы & Десятки(Mid(strЧисло, Поз + 1, 2), "м")

If strМиллиарды & strМиллионы & strТысячи & strЕдиницы = "" Then strЕдиницы = "ноль "

'strЕдиницы = strЕдиницы & ИмяРазряда(" ", Mid(strЧисло, Поз + 1, 2), "рубль ", "рубля ", "рублей ")

 

 

'Сотые'

'strСотые = strКопейки & " " & ИмяРазряда(strКопейки, Right(strКопейки, 2), ‘"копейка", "копейки", "копеек")

 

ЧислоПропись = strМиллиарды & strМиллионы & strТысячи & strЕдиницы

ЧислоПропись = UCase(Left(ЧислоПропись, 1)) & Right(ЧислоПропись, Len(ЧислоПропись) - 1)

 

Exit Function

 

Число_Error:

    MsgBox Err.Description

End Function

 

Function Сотни(n As String) As String

Сотни = ""

Select Case n

    Case 0: Сотни = ""

    Case 1: Сотни = "сто "

    Case 2: Сотни = "двести "

    Case 3: Сотни = "триста "

    Case 4: Сотни = "четыреста "

    Case 5: Сотни = "пятьсот "

    Case 6: Сотни = "шестьсот "

    Case 7: Сотни = "семьсот "

    Case 8: Сотни = "восемьсот "

    Case 9: Сотни = "девятьсот "

End Select

End Function

 

Function Десятки(n As String, Sex As String) As String

Десятки = ""

Select Case Left(n, 1)

    Case "0": Десятки = "": n = Right(n, 1)

    Case "1": Десятки = ""

    Case "2": Десятки = "двадцать ": n = Right(n, 1)

    Case "3": Десятки = "тридцать ": n = Right(n, 1)

    Case "4": Десятки = "сорок ": n = Right(n, 1)

    Case "5": Десятки = "пятьдесят ": n = Right(n, 1)

    Case "6": Десятки = "шестьдесят ": n = Right(n, 1)

    Case "7": Десятки = "семьдесят ": n = Right(n, 1)

    Case "8": Десятки = "восемьдесят ": n = Right(n, 1)

    Case "9": Десятки = "девяносто ": n = Right(n, 1)

End Select

 

Dim Двадцатка As String

Двадцатка = ""

Select Case n

    Case "0": Двадцатка = ""

    Case "1"

        Select Case Sex

            Case "м": Двадцатка = "один "

            Case "ж": Двадцатка = "одна "

            Case "с": Двадцатка = "одно "

        End Select

    Case "2":

        Select Case Sex

            Case "м": Двадцатка = "два "

            Case "ж": Двадцатка = "две "

            Case "с": Двадцатка = "два "

        End Select

    Case "3": Двадцатка = "три "

    Case "4": Двадцатка = "четыре "

    Case "5": Двадцатка = "пять "

    Case "6": Двадцатка = "шесть "

    Case "7": Двадцатка = "семь "

    Case "8": Двадцатка = "восемь "

    Case "9": Двадцатка = "девять "

    Case "10": Двадцатка = "десять "

    Case "11": Двадцатка = "одиннадцать "

    Case "12": Двадцатка = "двенадцать "

    Case "13": Двадцатка = "тринадцать "

    Case "14": Двадцатка = "четырнадцать "

    Case "15": Двадцатка = "пятнадцать "

    Case "16": Двадцатка = "шестнадцать "

    Case "17": Двадцатка = "семнадцать "

    Case "18": Двадцатка = "восемнадцать "

    Case "19": Двадцатка = "девятнадцать "

End Select

 

Десятки = Десятки & Двадцатка

End Function

 

Function ИмяРазряда(Строка As String, n As String, Имя1 As String, Имя24 As String, ИмяПроч As String) As String

 

If Строка <> "" Then

    ИмяРазряда = ""

    Select Case Left(n, 1)

        Case "0", "2", "3", "4", "5", "6", "7", "8", "9": n = Right(n, 1)

    End Select

 

    Select Case n

        Case "1": ИмяРазряда = Имя1

        Case "2", "3", "4": ИмяРазряда = Имя24

        Case Else: ИмяРазряда = ИмяПроч

    End Select

End If

 

End Function

