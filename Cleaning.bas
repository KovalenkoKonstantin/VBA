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
 SearchingString = "���� �������"
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
 SearchingString = "���������� ���� �������"

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
 SearchingString = "���������� ���� �������"

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