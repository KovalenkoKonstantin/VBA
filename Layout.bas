Attribute VB_Name = "Layout"
Sub LayoutOn()
    Dim ThisWorkbook As Workbook
    Dim ws As Worksheet
    Dim CompanyName As String
    Dim aws As Integer
    Dim tottal As Integer

    ' �������� ��������� ����� (��������������, ��� � ��� ���� ��������� HideSys)
    HideSys

    ' ������������� ������ �� �������� �����
    Set ThisWorkbook = ActiveWorkbook

    ' �������� ��� �������� �� ������ C7 �� ����� "Preferences"
    CompanyName = ThisWorkbook.Sheets("Preferences").Range("C7").Value2

    ' ��������� ������: ������� � ExitHandler � ������ ������
    On Error GoTo ExitHandler

    ' ��������� ���������� ������, �������, ������� ������� � ��������������
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayAlerts = False

    ' ������������ ���������� ������� ������
    tottal = 0
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            tottal = tottal + 1
        End If
    Next ws

    ' �������� �� ���� ������ � �����
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Activate
            aws = Application.ActiveSheet.Index

            ' ��������� ��������� ������
            Application.StatusBar = "�������������� " & aws & " ���� �� " & tottal & " ������. ���������: " & _
                                    Int(aws / tottal * 100) & " %. ��������� ����� �� ����� ���������� ���������: " & _
                                    Int((tottal - aws) * 3) & " ������(�)."

            ' ����������� �����������
            With ActiveSheet.PageSetup
                .CenterHeader = "&""Times New Roman,�������""&KFF0000������ �������� �� ����������."
                .RightFooter = "&""Times New Roman,�������""&KFF0000��������� �������� � ����� ���������� � ���� �������� ����������, ����������� � ������������ ����� " & CompanyName
            End With
        End If
    Next ws

ExitHandler:
    ' ��������������� ��������� Excel
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True

    ' ������������ �� ���� "Preferences"
    ThisWorkbook.Sheets("Preferences").Activate
    ActiveWindow.View = xlNormalView

    Exit Sub

ErrHandler:
    ' ������� ��������� �� ������ � ������������ � ExitHandler
    MsgBox Err.Description
    Resume ExitHandler
End Sub
Sub LayoutInfotecs()
    Dim ThisWorkbook As Workbook
    Dim ws As Worksheet
    Dim CompanyName As String
    Dim aws As Integer
    Dim tottal As Integer
    
    ' �������� ��������� ����� (��������������, ��� � ��� ���� ��������� HideSys)
    HideSys
    
    ' ������������� ������ �� �������� �����
    Set ThisWorkbook = ActiveWorkbook
    
    ' �������� ��� �������� �� ������ C7 �� ����� "Preferences"
    CompanyName = ThisWorkbook.Sheets("Preferences").Range("C7").Value2
    
    ' ��������� ������: ������� � ExitHandler � ������ ������
    On Error GoTo ExitHandler
    
    ' ��������� ���������� ������, �������, ������� ������� � ��������������
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayAlerts = False
    
    ' ������������ ���������� ������� ������
    tottal = 0
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            tottal = tottal + 1
        End If
    Next ws
    
    ' �������� �� ���� ������ � �����
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Activate
            aws = Application.ActiveSheet.Index
            
            ' ��������� ��������� ������
            Application.StatusBar = "�������������� " & aws & " ���� �� " & tottal & " ������. ���������: " & _
                                    Int(aws / tottal * 100) & " %. ��������� ����� �� ����� ���������� ���������: " & _
                                    Int((tottal - aws) * 3) & " ������(�)."
            
            ' ����������� ������ ����������
            With ActiveSheet.PageSetup
                .RightFooter = "&""Times New Roman,�������""&KFF0000��������� " & CompanyName
            End With
        End If
    Next ws
    
ExitHandler:
    ' ��������������� ��������� Excel
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    
    ' ������������ �� ���� "Preferences"
    ThisWorkbook.Sheets("Preferences").Activate
    ActiveWindow.View = xlNormalView
    
    Exit Sub
    
ErrHandler:
    ' ������� ��������� �� ������ � ������������ � ExitHandler
    MsgBox Err.Description
    Resume ExitHandler
End Sub

Sub LayoutOff()
    Dim ThisWorkbook As Workbook
    Dim ws As Worksheet
    Dim aws As Integer
    Dim tottal As Integer

    ' �������� ��������� ����� (��������������, ��� � ��� ���� ��������� HideSys)
    HideSys

    ' ������������� ������ �� �������� �����
    Set ThisWorkbook = ActiveWorkbook

    ' ��������� ������: ������� � ExitHandler � ������ ������
    On Error GoTo ExitHandler

    ' ��������� ���������� ������, �������, ������� ������� � ��������������
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayAlerts = False

    ' ������������ ���������� ������� ������
    tottal = 0
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            tottal = tottal + 1
        End If
    Next ws

    ' �������� �� ���� ������ � �����
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Activate
            aws = Application.ActiveSheet.Index

            ' ��������� ��������� ������
            Application.StatusBar = "�������������� " & aws & " ���� �� " & tottal & " ������. ���������: " & _
                                    Int(aws / tottal * 100) & " %. ��������� ����� �� ����� ���������� ���������: " & _
                                    Int((tottal - aws) * 3) & " ������."

            ' ������� ��������� ������������
            ActiveSheet.PageSetup.CenterHeaderPicture.Filename = ""
            With ActiveSheet.PageSetup
                .CenterHeader = ""
                .RightHeader = ""
                .RightFooter = ""
            End With
        End If
    Next ws

ExitHandler:
    ' ��������������� ��������� Excel
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True

    ' ������������ �� ���� "Preferences"
    ThisWorkbook.Sheets("Preferences").Activate
    ActiveWindow.View = xlNormalView

    Exit Sub

ErrHandler:
    ' ������� ��������� �� ������ � ������������ � ExitHandler
    MsgBox Err.Description
    Resume ExitHandler
End Sub


