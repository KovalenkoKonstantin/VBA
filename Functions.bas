Attribute VB_Name = "Functions"
Public Function �������(Salary As Double, Percentage As Double, month As Integer, Limit As Double) As Double
Dim result, FCC, PF, FFOMC, Tr, Extra As Double

'������� 7,8%
If Percentage = 0.078 Then
    FCC = 0.015
    PF = 0.06
    FFOMC = 0.001
    Tr = 0.002
    Extra = 0
End If

'������� 30,2%
If Percentage = 0.302 Then
    FCC = 0.029
    PF = 0.22
    FFOMC = 0.051
    Tr = 0.002
    Extra = 0.1
End If

'������� 14,2%
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

������� = result

End Function

Public Sub DescriptFor�������()
    Application.MacroOptions "�������", Description:="�������", ArgumentDescriptions:=Array("����� ����������", "���������� ������� ���������� ������", "�����(��� �����)", "���������� �������� ���������� �������� ����������������")
End Sub


Function �������������(n As Double) As String
 
 Dim Nums1, Nums2, Nums3, Nums4 As Variant
 
 Nums1 = Array("", "���� ", "��� ", "��� ", "������ ", "���� ", "����� ", "���� ", "������ ", "������ ")
 Nums2 = Array("", "������ ", "�������� ", "�������� ", "����� ", "��������� ", "���������� ", "��������� ", _
                        "����������� ", "��������� ")
 Nums3 = Array("", "��� ", "������ ", "������ ", "��������� ", "������� ", "�������� ", "������� ", _
                        "��������� ", "��������� ")
 Nums4 = Array("", "���� ", "��� ", "��� ", "������ ", "���� ", "����� ", "���� ", "������ ", "������ ")
 Nums5 = Array("������ ", "����������� ", "���������� ", "���������� ", "������������ ", _
                        "���������� ", "����������� ", "���������� ", "������������ ", "������������ ")
 
 If n <= 0 Then
   ������������� = "����"
   Exit Function
 End If
 '��������� ����� �� �������, ��������� ��������������� ������� Class
 ed = Class(n, 1)
 dec = Class(n, 2)
 sot = Class(n, 3)
 tys = Class(n, 4)
 dectys = Class(n, 5)
 sottys = Class(n, 6)
 mil = Class(n, 7)
 decmil = Class(n, 8)
 
 '��������� ��������
 Select Case decmil
   Case 1
     mil_txt = Nums5(mil) & "��������� "
     GoTo www
   Case 2 To 9
     decmil_txt = Nums2(decmil)
 End Select
 Select Case mil
   Case 1
     mil_txt = Nums1(mil) & "������� "
   Case 2, 3, 4
     mil_txt = Nums1(mil) & "�������� "
   Case 5 To 20
     mil_txt = Nums1(mil) & "��������� "
 End Select
www:
 sottys_txt = Nums3(sottys)
 '��������� ������
 Select Case dectys
   Case 1
     tys_txt = Nums5(tys) & "����� "
     GoTo eee
   Case 2 To 9
     dectys_txt = Nums2(dectys)
 End Select
 Select Case tys
   Case 0
     If dectys > 0 Then tys_txt = Nums4(tys) & "����� "
   Case 1
     tys_txt = Nums4(tys) & "������ "
   Case 2, 3, 4
     tys_txt = Nums4(tys) & "������ "
   Case 5 To 9
     tys_txt = Nums4(tys) & "����� "
 End Select
 If dectys = 0 And tys = 0 And sottys <> 0 Then sottys_txt = sottys_txt & "����� "
eee:
 sot_txt = Nums3(sot)
 '��������� �������
 Select Case dec
   Case 1
     ed_txt = Nums5(ed)
     GoTo rrr
   Case 2 To 9
     dec_txt = Nums2(dec)
 End Select
 
 ed_txt = Nums1(ed)
rrr:
 '��������� �������� ������
 ������������� = UCase(Left(decmil_txt & mil_txt & sottys_txt & dectys_txt & tys_txt & sot_txt & dec_txt & ed_txt, 1)) & LCase(Mid(decmil_txt & mil_txt & sottys_txt & dectys_txt & tys_txt & sot_txt & dec_txt & ed_txt, 2))
End Function
 
'��������������� ������� ��� ��������� �� ����� ��������
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
' Purpose   : MsgBox � ��������� (������������ ����� Popup WScript.Shell)
'             ������ .VBS-���� �� ��������� �����, ��������� ���, ���������� ���� ����������, ������� ��������� ����
' Arguments : ������ 3 ��������� ����� ��, ��� � MsgBox, 4-� - ������� � ��������.
'             ���� 4-� �������� �� ����� ��� <=0, �� ������� ����� ������������ ��� ������� MsgBox
' Ret.Value : ����� ��, ��� � Msgbox, �� ���������� -1 �� ��������� ��������.
' Errors    : ���������� ������ 735 - "Can't save file to TEMP" ���� ��������� ����� �� ��������
' Author    : ���������, [email]exceleved@yandex.ru[/email]
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
      sTmp = sTmp & Format$(Now, """\~MsgBoxEx""YYYYMMDDHHMMSS"".vbs""")   '���������� ��� �����
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
Private Function Str2Code$(sTxt) ' �������� CR+LF, LF+CR, CR, LF �� " & vblf & " ��� ������������� � VBS
   Str2Code = Replace(Replace(Replace(Replace(Replace(sTxt, """", """"""), vbCrLf, vbLf), vbLf & vbCr, vbLf), vbCr, vbLf), vbLf, """ & vblf & """)
End Function

Function ������������(����� As Currency) As String

'�� 999 999 999 999

On Error GoTo �����_Error

Dim str��������� As String, str�������� As String, str������ As String, str������� As String, str����� As String

Dim ��� As Integer

 

str����� = Format(Int(�����), "000000000000")

 

'���������'

��� = 1

str��������� = �����(Mid(str�����, ���, 1))

str��������� = str��������� & �������(Mid(str�����, ��� + 1, 2), "�")

str��������� = str��������� & ����������(str���������, Mid(str�����, ��� + 1, 2), "�������� ", "��������� ", "���������� ")

 

'��������'

��� = 4

str�������� = �����(Mid(str�����, ���, 1))

str�������� = str�������� & �������(Mid(str�����, ��� + 1, 2), "�")

str�������� = str�������� & ����������(str��������, Mid(str�����, ��� + 1, 2), "������� ", "�������� ", "��������� ")

 

'������'

��� = 7

str������ = �����(Mid(str�����, ���, 1))

str������ = str������ & �������(Mid(str�����, ��� + 1, 2), "�")

str������ = str������ & ����������(str������, Mid(str�����, ��� + 1, 2), "������ ", "������ ", "����� ")

 

'�������'

��� = 10

str������� = �����(Mid(str�����, ���, 1))

str������� = str������� & �������(Mid(str�����, ��� + 1, 2), "�")

If str��������� & str�������� & str������ & str������� = "" Then str������� = "���� "

'str������� = str������� & ����������(" ", Mid(str�����, ��� + 1, 2), "����� ", "����� ", "������ ")

 

 

'�����'

'str����� = str������� & " " & ����������(str�������, Right(str�������, 2), �"�������", "�������", "������")

 

������������ = str��������� & str�������� & str������ & str�������

������������ = UCase(Left(������������, 1)) & Right(������������, Len(������������) - 1)

 

Exit Function

 

�����_Error:

    MsgBox Err.Description

End Function

 

Function �����(n As String) As String

����� = ""

Select Case n

    Case 0: ����� = ""

    Case 1: ����� = "��� "

    Case 2: ����� = "������ "

    Case 3: ����� = "������ "

    Case 4: ����� = "��������� "

    Case 5: ����� = "������� "

    Case 6: ����� = "�������� "

    Case 7: ����� = "������� "

    Case 8: ����� = "��������� "

    Case 9: ����� = "��������� "

End Select

End Function

 

Function �������(n As String, Sex As String) As String

������� = ""

Select Case Left(n, 1)

    Case "0": ������� = "": n = Right(n, 1)

    Case "1": ������� = ""

    Case "2": ������� = "�������� ": n = Right(n, 1)

    Case "3": ������� = "�������� ": n = Right(n, 1)

    Case "4": ������� = "����� ": n = Right(n, 1)

    Case "5": ������� = "��������� ": n = Right(n, 1)

    Case "6": ������� = "���������� ": n = Right(n, 1)

    Case "7": ������� = "��������� ": n = Right(n, 1)

    Case "8": ������� = "����������� ": n = Right(n, 1)

    Case "9": ������� = "��������� ": n = Right(n, 1)

End Select

 

Dim ��������� As String

��������� = ""

Select Case n

    Case "0": ��������� = ""

    Case "1"

        Select Case Sex

            Case "�": ��������� = "���� "

            Case "�": ��������� = "���� "

            Case "�": ��������� = "���� "

        End Select

    Case "2":

        Select Case Sex

            Case "�": ��������� = "��� "

            Case "�": ��������� = "��� "

            Case "�": ��������� = "��� "

        End Select

    Case "3": ��������� = "��� "

    Case "4": ��������� = "������ "

    Case "5": ��������� = "���� "

    Case "6": ��������� = "����� "

    Case "7": ��������� = "���� "

    Case "8": ��������� = "������ "

    Case "9": ��������� = "������ "

    Case "10": ��������� = "������ "

    Case "11": ��������� = "����������� "

    Case "12": ��������� = "���������� "

    Case "13": ��������� = "���������� "

    Case "14": ��������� = "������������ "

    Case "15": ��������� = "���������� "

    Case "16": ��������� = "����������� "

    Case "17": ��������� = "���������� "

    Case "18": ��������� = "������������ "

    Case "19": ��������� = "������������ "

End Select

 

������� = ������� & ���������

End Function

 

Function ����������(������ As String, n As String, ���1 As String, ���24 As String, ������� As String) As String

 

If ������ <> "" Then

    ���������� = ""

    Select Case Left(n, 1)

        Case "0", "2", "3", "4", "5", "6", "7", "8", "9": n = Right(n, 1)

    End Select

 

    Select Case n

        Case "1": ���������� = ���1

        Case "2", "3", "4": ���������� = ���24

        Case Else: ���������� = �������

    End Select

End If

 

End Function

