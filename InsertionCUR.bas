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
 Limit = 153 '��������� ������� ����
 begin = 12 '������ ��� �������
 CompanyName = ThisWorkbook.Sheets("Preferences").Range("C7").Value2 '��� �������
 
' Dim aw(1 To 147) As Variant
' Dim iw(1 To 147) As Variant

' ��������� ������� ��� ������������
Dim aw() As Variant
Dim iw() As Variant

'������������� ������ ��������
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
 MultiSelect:=True, Title:="�������� ��������� ��������� �� �������� " _
 & CompanyName & " �� " & Year(Date) & " ���")
 
 '������ ���
Application.StatusBar = "������ ������..."

 If TypeName(FilesToOpen) = "Boolean" Then ',���� ������ ������ ������ ����� �� ���������
 GoTo ExitHandler
 End If

'������� ������
Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

On Error Resume Next

 '�������� ������������ ������ ������
 importWB.Sheets(1).Activate
 Range("G2").Select
 ActiveCell.FormulaR1C1 = "=YEAR(MID(RC[-4],SEARCH("" "",RC[-4],1)+1,10))"
' If Range("G2").Value2 <> DistinctYear Or Range("A11").Value2 <> CompanyName Then
 If Range("A11").Value2 <> CompanyName Then
    Range("G2").Select
    With Selection:
        .Clear
    End With
    MsgBoxEx "������� ������������ ��������� ���������. ������������ �������� �� ���������." _
    & vbCr & "������� �������.", vbCritical, "Bad Day", 20
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

'����������� ������� ������� �����
On Error Resume Next
For i = 1 To 20
    If Worksheets(SheetName).Cells(i, 1) = "���������" Then
        DataRow = i
    End If
Next i

' ���������� ������� � ������� � ���������������� ���������
    dict.Add "���������", 1
    dict.Add "�����", 2
    dict.Add "��������� ����� �����", 3
    dict.Add "������ �����", 4
    dict.Add "��������� ����� ����", 5
    dict.Add "������ ����", 6
    dict.Add "���������� ���� ����� 20,26,44 �����", 7
    dict.Add "��� ��������", 8
    dict.Add "������ ��������� �������", 9
    dict.Add "������ �� ����������.������ �� ��", 10
    dict.Add "������������� �������", 11
    dict.Add "���������", 12
    dict.Add "��� ���������", 13
    dict.Add "���� ��������", 14
    dict.Add "������ ��������� �������� � ��� �����", 15
    dict.Add "������ ������", 16
    dict.Add "����� ����", 17
    dict.Add "����� �����", 18
    dict.Add "����", 19
    dict.Add "�����", 20
    dict.Add "���������", 21
    dict.Add "������ ���������� ������ �� ���� ������������", 22
    dict.Add "������ ���������� ������", 23
    dict.Add "����� �� ����", 24
    dict.Add "���������� �� ������� (���������� ��� �� ������)", 25
    dict.Add "�������� �� ��������� � �������������", 26
    dict.Add "������ ������� �� ����������� ����", 27
    dict.Add "������ �� ����� ������", 28
    dict.Add "�������� �� ��������������� � �������� ����������� �����", 29
    dict.Add "������ �� ������ ����", 30
    dict.Add "������ �� ���� ����", 31
    dict.Add "����������� ������� (������ ��������)", 32
    dict.Add "����������� ������", 33
    dict.Add "������ �������", 34
    dict.Add "�������� ������� ��� ����������", 35
    dict.Add "������� �� ����� �� �������� �� �������� ���", 36
    dict.Add "����� �� ���� (��������������� ������������ ����)", 37
    dict.Add "�������� �� ��������������� � �������� ����������� ����� (��������������� ������������ ����)", 38
    dict.Add "�������� �� ������ �� ����������, ������������� ��������������� �����", 39
    dict.Add "����� �� ���� 26 �����������", 40
    dict.Add "������ �� �����������", 41
    dict.Add "�������� ��������", 42
    dict.Add "�������� �����������", 43
    dict.Add "������ � ������ ��������� ������������(�����������)", 44
    dict.Add "��������� �������������� ������������ ������", 45
    dict.Add "������ �� ������ ���� � ������ ��", 46
    dict.Add "�������������� �� �����������", 47
    dict.Add "������� (������, ������)", 48
    dict.Add "����������� �� ������", 49
    dict.Add "������ � ������ ��������� ������������ (��������)", 50
    dict.Add "������", 51
    dict.Add "������� �� ���������� ����������, ���������� ������������", 52
    dict.Add "�������� ������", 53
    dict.Add "������� �����", 54
    dict.Add "�������������� ������� ������ (������������)", 55
    dict.Add "������ ������ � ����������� � �������� ���", 56
    dict.Add "�������������� ������ ������ ����������", 57
    dict.Add "������ �������� �.�", 58
    dict.Add "���������� �� ������������ �������", 59
    dict.Add "������ ��� ������ �������� �� ��", 60
    dict.Add "������ ���� ����� �� ������-����������", 61
    dict.Add "������� �� ����� �� ������� �� 3 ��� ��� ������", 62
    dict.Add "����� �� ����� (������������� ������������� �������)", 63
    dict.Add "�������� �� ��������� � ������������� (�� ����� ��������������� ������������ �������)", 64
    dict.Add "�������� �� ��������������� � �������� ����������� ����� (�� ����� �������. �����. �������)", 65
    dict.Add "���������� �� �������", 66
    dict.Add "������� �� ������ � ������ �����", 67
    dict.Add "�� ������ ������ �� ���", 68
    dict.Add "������ ���������� (���������, �����, ����������, �������)", 69
    dict.Add "�������� �� ���������� ������������������ � ��������� �", 70
    dict.Add "������ �� ������������ � �����", 71
    dict.Add "������� �� ������ � ����������� ��� (������� �����)", 72
    dict.Add "������� �� ������ � ����������� ��� (������ �����)", 73
    dict.Add "������� �� ����������� ��� ������������� ����� �������� �������", 74
    dict.Add "�������������� ��� ������� �� ������������ � �����", 75
    dict.Add "������� �� ��������������� ������� ����", 76
    dict.Add "������������ ������", 77
    dict.Add "����� �� ����� � ����������� �����", 78
    dict.Add "������� �� ���� ������� �����������", 79
    dict.Add "����� � ����������� ����� (���. ����)", 80
    dict.Add "���������", 81
    dict.Add "���� � ������ � ���� ������� ��������", 82
    dict.Add "������� ������������", 83
    dict.Add "�������", 84
    dict.Add "��������� ��������", 85
    dict.Add "������� ������� ����������� ����", 86
    dict.Add "������� �����", 87
    dict.Add "��������", 88
    dict.Add "����", 89
    dict.Add "���� � ����������", 90
    dict.Add "�������� �� �������� �� ������ ������������", 91
    dict.Add "��������� �� �/� ��������������� ��������", 92
    dict.Add "��������� �� ���. ����� ���������", 93
    dict.Add "��������� �� ���������", 94
    dict.Add "������� ������������� �������� ����������", 95
    dict.Add "�������� �� �������� �� ������ ������������-��������������� �����", 96
    dict.Add "��������� �� �������� �� ��������", 97
    dict.Add "��������� �� �������� 73,03", 98
    dict.Add "��������� �� ���. ����� ����. ������", 99
    dict.Add "��������� �� ��������������� ���������", 100
    dict.Add "������", 101
    dict.Add "��� �� ����������", 102
    dict.Add "��� � ����������", 103
    dict.Add "�����", 104
    dict.Add "���", 105
    dict.Add "��� ��", 106
    dict.Add "% ��������� �������", 107
    dict.Add "���� �������", 108
    dict.Add "�������� �� ���� ������ �� ������ ��������������� �����", 109
    dict.Add "������ �� �������������� ���� (���) ������ ������", 110
    dict.Add "������������", 111
    dict.Add "������ ����������� ����� ������ (��)", 112
    dict.Add "������ ��������", 113
    dict.Add "������ �� ������ (�� �����)", 114
    dict.Add "����� �� �����", 115
    dict.Add "�������� �� ��������� � ������������� (�� �����)", 116
    dict.Add "�������� �� ��������� � ������������� (��������������� ������������ ����)", 117
    dict.Add "����������� ������", 118
    dict.Add "����������� �������� �� ��������� �������", 119
    dict.Add "������ �������� (� ������ ��)", 120
    dict.Add "������ �����������", 121
    dict.Add "������ ����������� (� ������ ��)", 122
    dict.Add "������ �� ������ ���� (� ������ ��)", 123
    dict.Add "�������� �� ��������� � ������������� (�� ����� �������. ������������� �������)", 124
    dict.Add "������ ������� (� ������ ��)", 125
    dict.Add "���� ����������", 126
    dict.Add "������ ��������", 127
    dict.Add "���", 128
    dict.Add "��� �����", 133
    dict.Add "������� �� ������ � ������ ����� (����������� � �������� ���)", 134
    dict.Add "������ ����������� (� ������ ��)", 135
    dict.Add "������ �����������", 136
    dict.Add "����������� ������� (������ �����, ���������� � ������� �������� ������)", 137
    dict.Add "���� ������� ����", 138
    dict.Add "% ��������� ������� ����", 139
    dict.Add "����� � ����������� ����� (���)", 140
    dict.Add "������ ������� �� ����������� ���� (�� 2025)", 141
    dict.Add "������ ���������� ������ �� ���� ������������ (�� 2025)", 142
    dict.Add "����������� ������� (������ ��������) (�� 2025)", 143
    dict.Add "�������������� ������� ������ (������������) (�� 2025)", 144
    dict.Add "������ �� �������������� ���� (���) ������ ������ (�� 2025)", 145
    dict.Add "����������� ������� (������ �����, ���������� � ������� �������� ������) (�� 2025)", 146
    dict.Add "������ ������� �� ����������� ���� (���� ��)", 147
    dict.Add "����������� ������� ��� ���������� �� ����������� ����", 148
    dict.Add "����������� ������� ��� ���������� �� ����������� ���� (���� ��)", 149
    dict.Add "����������� ������� (������ �����, ���������� � ������� �������� ������) (���� ��)", 150
    dict.Add "����������� ������� (������ �����, ���������� � ������� �������� ������) (���� ��)", 151
    dict.Add "����������� ������� ��� ���������� �� ����������� ���� (���� ��)", 152
    dict.Add "������ ������� �� ����������� ���� (���� ��)", 153


' ������� �������� � ������� � ���������� �������
    For Each Key In dict.Keys
        For i = 1 To Limit
            Item = dict(Key)
            If Worksheets(SheetName).Cells(DataRow, i) = Key Then
                aw(Item) = i
            End If
        Next i
    Next Key

 
 importWB.Sheets(1).Activate

'����������� ������� ������������� �����
For i = 1 To 20
    If importWB.Sheets(1).Cells(i, 1) = "�����������" Then
        ImportFirstDataRow = i
    End If
Next i
For i = 1 To 20
    If importWB.Sheets(1).Cells(i, 1) = "���������" Then
        ImportSecondDataRow = i
    End If
Next i

' ������� �������� � ������� � ���������� �������
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


'�������� ���������� ������
ThisWorkbook.Sheets(SheetName).Activate
awLastRow = Cells(Rows.Count, "A").End(xlUp).row
awLastCol = Cells(begin, Columns.Count).End(xlUp).column
Range(Cells(begin, 1), Cells(awLastRow, Limit)).Select
 With Selection
        .Clear
 End With

 '������ ���
Application.StatusBar = "������� ������"

 '������� ������
 importWB.Sheets(1).Activate
 iwLastRow = Cells(Rows.Count, "A").End(xlUp).row

For i = 1 To Limit
'������ ���
Application.StatusBar = "������������� ����. ���������: " & Int(100 * i / Limit) & "%." & _
" ����� ��������: " & Int(87 * i / Limit) & "%" & _
" ��������� ����� �� ����� ���������� ���������: " & _
Int((100 - Int(87 * i / Limit)) * (((Now() - Start) * 24 * 60 * 60) / (Int(87 * i / Limit)))) & " ������"
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

'������ ���
Application.StatusBar = "�������������� �����. ���������: 87 %"

'�������
ThisWorkbook.Sheets(SheetName).Activate
Columns("Q:DD").Select
    Selection.NumberFormat = _
        "_-* #,##0.00 _?_-;-* #,##0.00 _?_-;_-* ""-""?? _?_-;_-@_-"

'������� ����������� ������
ThisWorkbook.Sheets(SheetName).Activate
'������ ���
Application.StatusBar = "���������� ����������� ������ ������. ���������: 88 %"
'�����
Cells(begin, aw(2)).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(IFERROR(SEARCH("" 20"",RC[-1],1)>0,FALSE)," _
        & "TRIM(MID(IF(IFERROR(SEARCH("" 20"",RC[-1],1)>0,FALSE)," _
        & "RC[-1],R[-1]C),1,SEARCH("" "",IF(IFERROR(SEARCH("" 20"",RC[-1],1)>0,FALSE)," _
        & "RC[-1],R[-1]C),1)-1)),R[-1]C)"
    Cells(begin, aw(2)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(2)), Cells(iwLastRow, aw(2)))

'������ ���
Application.StatusBar = "���������� ������ ��������� ����� �����. ���������: 89 %"
'��������� ����� �����
K = 3
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = "=IF(OR(CONCATENATE(RC[-1],"" "",RC[5])=RC[-2]," _
    & "VLOOKUP(RC[-2],RC1:RC214,MATCH(R8C4,R9C1:R9C214,0),0)>0," _
    & "VLOOKUP(RC[-2],RC1:RC214,MATCH(R7C4,R9C1:R9C214,0),0)>0," _
    & "VLOOKUP(RC[-2],RC1:RC214,MATCH(R6C4,R9C1:R9C214,0),0)>0," _
    & "RC[4]=TRUE),"""",VLOOKUP(RC[-1],INDIRECT(CONCATENATE(""'"",VALUE(RC[5])," _
    & """ �����. ���������'!$A:$BR"")),HLOOKUP(RC[20],INDIRECT" _
    & "(CONCATENATE(""'"",VALUE(RC" & _
        "[5]),"" �����. ���������'!$2:$3"")),2,0),0))" & _
        ""
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
    
'������ ���
Application.StatusBar = "���������� ������ ������� ����� �����. ���������: 90 %"
'������ �����
K = 4
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
        "=OR(RC[-1]="""",OR(RC[-1]=" _
        & "VALUE(RC[21]),VLOOKUP(RC[-3],RC1:RC114," _
        & "MATCH(R8C4,R9C1:R9C114,0),0)>0),SUM(RC[22]:RC[124])=0," _
        & "NOT(ISNA(MATCH(RC[-3]&RC[-2]&RC[4],���.����!C7,0))))"
    Range("A1").Copy
    Cells(begin, aw(K)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
   
'������ ���
Application.StatusBar = "���������� ������ ����� ����. ���������: 91 %"
'��������� ����� ����
K = 5
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(CONCATENATE(RC[-3],"" "",RC[3])=RC[-4]," _
        & "VLOOKUP(RC[-4],RC1:RC114,MATCH(R8C4,R9C1:R9C114,0),0)>0," _
        & "RC[2]=TRUE),"""",VLOOKUP(RC[-3],INDIRECT(CONCATENATE(""'"",VALUE(RC[3])," _
        & """ �����. ���������'!$A$18:$BR$31"")),HLOOKUP(RC[18]," _
        & "INDIRECT(CONCATENATE(""'"",VALUE(RC[3]),"" �����. ���������'!$18:$19"")),2,0),0))"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
    
'������ ���
Application.StatusBar = "���������� ������ ������� ����� ����. ���������: 92 %"
'������ ����
K = 6
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
        "=OR(RC[-3]="""",OR(RC[-1]=VALUE(RC[18]),VLOOKUP(RC[-5]," _
        & "RC1:RC114,MATCH(R8C4,R9C1:R9C114,0),0)>0),SUM(RC[20]:RC[122])=0," _
        & "NOT(ISNA(MATCH(RC[-5]&RC[-4]&RC[2],���.����!C7,0))))"
    Range("A1").Copy
    Cells(begin, aw(K)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))

'������ ���
Application.StatusBar = "���������� ������ ���������� �� ������� ���� ����� 20,26,44 ������. ���������: 93 %"
'���������� ���� ����� 20,26,44 �����
K = 7
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
       "=NOT(OR(IFERROR(SEARCH(20,RC[15],1),FALSE)," _
       & "IFERROR(SEARCH(26,RC[15],1),FALSE),IFERROR(SEARCH(44,RC[15],1),FALSE)))"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))
    
'������ ���
Application.StatusBar = "���������� ������ ��������� ����� �������� �����������. ���������: 94 %"
'��� ��������
K = 8
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(CONCATENATE(RC[-12],"" "",RC[-6])=RC[-13],""""," _
        & "(MID(RC[-13],SEARCH("" "",RC[-13],1),LEN(RC[-13]))))"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))

'������ ���
Application.StatusBar = "���������� ������ � �������������� ������� ��������� �������. ���������: 95 %"
'������ ��������� �������
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


'������ ���
Application.StatusBar = "���������� ������ �� ������ � ����������� �����. ���������: 96 %"
'����� �� ����� � ����������� �����
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
    
'������ ���
Application.StatusBar = "���������� ������ ��������� ����. ���������: 97 %"
'���
K = 128
Cells(begin, aw(K)).Select
    ActiveCell.FormulaR1C1 = _
       "=VALUE(IF(IFERROR(SEARCH("" 20"",RC[-7],1)>0,FALSE)," _
        & "MID(RC[-7],SEARCH("" "",RC[-7],1)+1,4),R[-1]C))"
    Cells(begin, aw(K)).Select
    Selection.AutoFill Destination:=Range(Cells(begin, aw(K)), Cells(iwLastRow, aw(K)))


'������ ���
Application.StatusBar = "�������������� ����������. ���������: 98 %"

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
    
'������ ���
Application.StatusBar = "�������������� ����������. ���������: 99 %"
' �������������� ������� ���� ��������� �������
'������ � ������� ���������
columnsToFormat = Array("EC", "DW", "DZ", "ED", "EI", "EH", "EO", "EQ", "EP", _
            "ER", "ES", "EU", "EV", "EW")

' �������� �� ������� ������� � ������� � ����� ������
For bound = LBound(columnsToFormat) To UBound(columnsToFormat)
    Columns(columnsToFormat(bound) & ":" & columnsToFormat(bound)).NumberFormat = _
        "_-* #,##0.00 _?_-;-* #,##0.00 _?_-;_-* ""-""?? _?_-;_-@_-"
Next bound

'������ ���
Application.StatusBar = "���������: 100 %"

'����������

'ThisWorkbook.Sheets(SheetName).Protect Password:=ps
'ThisWorkbook.Sheets(SheetName).Visible = False
'MsgBoxEx "��������� ��������� " _
'    & "�� �������� " & vbCr & ThisWorkbook.Sheets("Preferences").Range("C7").Value2 _
'    & vbCr & "�� " & ThisWorkbook.Sheets(SheetName).Range("B2").Value2 & " ���" _
'    & vbCr & "��������� �������", 0, "���������", 25

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








