Attribute VB_Name = "Layout"
Sub LayoutOn()

 Dim ThisWorkbook As Workbook
 Dim ws As Worksheet
 HideSys
 tottal = Application.Sheets.Count
 
' Dim xSht As Variant
'    Dim I As Long
'    For Each xSht In ActiveWorkbook.Sheets
'        If xSht.Visible Then I = I + 1
'    Next
'tottal = I

Set ThisWorkbook = ActiveWorkbook
CompanyName = ThisWorkbook.Sheets("Preferences").Range("C7").Value2 '��� ��������
 On Error GoTo ExitHandler
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False

For Each ws In ThisWorkbook.Worksheets
    If ws.Visible = True Then
        ws.Activate
        aws = Application.ActiveSheet.Index
    
    '������ ���
    Application.StatusBar = "�������������� " + Str(aws) _
    + " ���� �� " + Str(tottal) + " ������. ���������: " _
    + Str(Int(aws / tottal * 100)) + " %. " + "��������� ����� �� ����� ���������� ���������: " _
    + Str(Int((Str(tottal) - Str(aws)) * 3)) + " ������(�)."

'        ActiveSheet.PageSetup.CenterHeaderPicture.Filename = _
'        "D:\���������.png"

            With ActiveSheet.PageSetup
'                .CenterHeader = "&G"
                .CenterHeader = _
                "&""Times New Roman,�������""&KFF0000������ �������� �� ����������."
'                .RightHeader = _
'                "&""Times New Roman,�������""&KFF0000����� ������������������� ���������, ��������������� ��� ����������� ���������."
                .RightFooter = _
                "&""Times New Roman,�������""&KFF0000��������� �������� � ����� ���������� � ���� �������� ���������� ����������� � ������������ ����� " & CompanyName
            End With
    End If
Next ws

ExitHandler:
Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
ThisWorkbook.Sheets("Preferences").Activate
ActiveWindow.View = xlNormalView
 
Exit Sub
 
ErrHandler:
MsgBox Err.Description
Resume ExitHandler

End Sub


Sub LayoutOff()

 Dim ThisWorkbook As Workbook
 Dim ws As Worksheet
 HideSys
 tottal = Application.Sheets.Count

Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False

For Each ws In ThisWorkbook.Worksheets
    If ws.Visible = True Then
        ws.Activate
        aws = Application.ActiveSheet.Index
        
        '������ ���
        Application.StatusBar = "�������������� " + Str(aws) _
        + " ���� �� " + Str(tottal) + " ������. ���������: " _
        + Str(Int(aws / tottal * 100)) + " %. " + "��������� ����� �� ����� ���������� ���������: " _
        + Str(Int((Str(tottal) - Str(aws)) * 3)) + " ������."
        
        ActiveSheet.PageSetup.CenterHeaderPicture.Filename = ""
        
            With ActiveSheet.PageSetup
                .CenterHeader = ""
                .RightHeader = ""
                .RightFooter = ""
            End With
    End If
Next ws
ExitHandler:
Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
ThisWorkbook.Sheets("Preferences").Activate
ActiveWindow.View = xlNormalView
 
Exit Sub
 
ErrHandler:
MsgBox Err.Description
Resume ExitHandler

End Sub



