Attribute VB_Name = "Cleaning"
Sub DecreaseWeightProcessing21()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "Processing21"
 SearchingString = "���� �������" '���� ��������� ��������� �������
 begin = 12 '������ ��� �������
 
 '����������� ������� ������� �����
On Error Resume Next
For I = 1 To 20
    If Worksheets(SheetName).Cells(I, 1) = "���������" Then
        DataRow = I
    End If
Next I

 '����������� ��������� ��������� �������� ������� �����
On Error Resume Next
For I = 1 To 200
    If Worksheets(SheetName).Cells(DataRow, I) = SearchingString Then
        awLastCol = I
    End If
Next I

'�������� ���������� ������
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 [a12] = 1
 '����������
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub
Sub DecreaseWeightProcessing22()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "Processing22"
 SearchingString = "���� �������" '���� ��������� ��������� �������
 begin = 12 '������ ��� �������
 
 '����������� ������� ������� �����
On Error Resume Next
For I = 1 To 20
    If Worksheets(SheetName).Cells(I, 1) = "���������" Then
        DataRow = I
    End If
Next I

 '����������� ��������� ��������� �������� ������� �����
On Error Resume Next
For I = 1 To 200
    If Worksheets(SheetName).Cells(DataRow, I) = SearchingString Then
        awLastCol = I
    End If
Next I

'�������� ���������� ������
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 [a12] = 1
 '����������
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub
Sub DecreaseWeightProcessing23()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "Processing23"
 SearchingString = "���� �������" '���� ��������� ��������� �������
 begin = 12 '������ ��� �������
 
 '����������� ������� ������� �����
On Error Resume Next
For I = 1 To 20
    If Worksheets(SheetName).Cells(I, 1) = "���������" Then
        DataRow = I
    End If
Next I

 '����������� ��������� ��������� �������� ������� �����
On Error Resume Next
For I = 1 To 200
    If Worksheets(SheetName).Cells(DataRow, I) = SearchingString Then
        awLastCol = I
    End If
Next I

'�������� ���������� ������
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 [a12] = 1
 '����������
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub
Sub DecreaseWeightProcessing24()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "Processing24"
 SearchingString = "���� �������" '���� ��������� ��������� �������
 begin = 12 '������ ��� �������
 
 '����������� ������� ������� �����
On Error Resume Next
For I = 1 To 20
    If Worksheets(SheetName).Cells(I, 1) = "���������" Then
        DataRow = I
    End If
Next I

 '����������� ��������� ��������� �������� ������� �����
On Error Resume Next
For I = 1 To 200
    If Worksheets(SheetName).Cells(DataRow, I) = SearchingString Then
        awLastCol = I
    End If
Next I

'�������� ���������� ������
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 [a12] = 1
 '����������
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub

Sub DecreaseWeight���21()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "���21"
 SearchingString = "���������� ���� �������" '���� ��������� ��������� �������

 begin = 15 '������ ��� �������
 
 '����������� ������� ������� �����
On Error Resume Next
For I = 1 To 20
    If Worksheets(SheetName).Cells(I, 1) = "���������" Then
        DataRow = I
    End If
Next I

 '����������� ��������� ��������� �������� ������� �����
On Error Resume Next
For I = 1 To 200
    If Worksheets(SheetName).Cells(DataRow, I) = SearchingString Then
        awLastCol = I
    End If
Next I

'�������� ���������� ������
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 '����������
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub

Sub DecreaseWeight���22()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "���22"
 SearchingString = "���������� ���� �������" '���� ��������� ��������� �������

 begin = 15 '������ ��� �������
 
 '����������� ������� ������� �����
On Error Resume Next
For I = 1 To 20
    If Worksheets(SheetName).Cells(I, 1) = "���������" Then
        DataRow = I
    End If
Next I

 '����������� ��������� ��������� �������� ������� �����
On Error Resume Next
For I = 1 To 200
    If Worksheets(SheetName).Cells(DataRow, I) = SearchingString Then
        awLastCol = I
    End If
Next I

'�������� ���������� ������
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 '����������
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub
Sub DecreaseWeight���23()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "���23"
 SearchingString = "���������� ���� �������" '���� ��������� ��������� �������

 begin = 15 '������ ��� �������
 
 '����������� ������� ������� �����
On Error Resume Next
For I = 1 To 20
    If Worksheets(SheetName).Cells(I, 1) = "���������" Then
        DataRow = I
    End If
Next I

 '����������� ��������� ��������� �������� ������� �����
On Error Resume Next
For I = 1 To 200
    If Worksheets(SheetName).Cells(DataRow, I) = SearchingString Then
        awLastCol = I
    End If
Next I

'�������� ���������� ������
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 '����������
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub
Sub DecreaseWeight���24()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "���24"
 SearchingString = "���������� ���� �������" '���� ��������� ��������� �������

 begin = 15 '������ ��� �������
 
 '����������� ������� ������� �����
On Error Resume Next
For I = 1 To 20
    If Worksheets(SheetName).Cells(I, 1) = "���������" Then
        DataRow = I
    End If
Next I

 '����������� ��������� ��������� �������� ������� �����
On Error Resume Next
For I = 1 To 200
    If Worksheets(SheetName).Cells(DataRow, I) = SearchingString Then
        awLastCol = I
    End If
Next I

'�������� ���������� ������
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 ThisWorkbook.Sheets(SheetName).Range("C5").Select
 With Selection
        .Clear
 End With
 
 
 '����������
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub

Sub DecreaseWeight��_�������()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "��_�������"
 SearchingString = "���� ������� �� �������" '���� ��������� ��������� �������
 begin = 12 '������ ��� �������
 
 '����������� ������� ������� �����
On Error Resume Next
For I = 1 To 20
    If Worksheets(SheetName).Cells(I, 1) = "���������" Then
        DataRow = I
    End If
Next I

 '����������� ��������� ��������� �������� ������� �����
On Error Resume Next
For I = 1 To 200
    If Worksheets(SheetName).Cells(DataRow, I) = SearchingString Then
        awLastCol = I
    End If
Next I

'�������� ���������� ������
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 [a12] = 1
 '����������
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub
Sub DecreaseWeightExpenditures()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "Expenditures"
 SearchingString = "���� �������" '���� ��������� ��������� �������
 begin = 12 '������ ��� �������
 
 '����������� ������� ������� �����
On Error Resume Next
For I = 1 To 20
    If Worksheets(SheetName).Cells(I, 1) = "���������" Then
        DataRow = I
    End If
Next I

 '����������� ��������� ��������� �������� ������� �����
On Error Resume Next
For I = 1 To 200
    If Worksheets(SheetName).Cells(DataRow, I) = SearchingString Then
        awLastCol = I
    End If
Next I

'�������� ���������� ������
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 [a12] = 1
 '����������
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 5

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub
Sub DecreaseWeightBudget()

 Dim ThisWorkbook As Workbook
 Dim SheetName As String
 Set ThisWorkbook = ActiveWorkbook
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

 SheetName = "������"
 SearchingString = "������ ������" '���� ��������� ��������� �������
 begin = 5 '������ ��� �������
 
 '����������� ������� ������� �����
On Error Resume Next
For I = 1 To 20
    If Worksheets(SheetName).Cells(I, 1) = "���������" Then
        DataRow = I
    End If
Next I

 '����������� ��������� ��������� �������� ������� �����
On Error Resume Next
For I = 1 To 200
    If Worksheets(SheetName).Cells(DataRow, I) = SearchingString Then
        awLastCol = I
    End If
Next I

'�������� ���������� ������
 ThisWorkbook.Sheets(SheetName).Activate
 awLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(begin, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With
 [a5] = 1
 '����������
ThisWorkbook.Sheets("Preferences").Activate
MsgBoxEx "Data cleaned", 0, "Done!", 2

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    
End Sub
