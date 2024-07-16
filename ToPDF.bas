Attribute VB_Name = "ToPDF"
Sub SaveToPDF()
    Dim Start As Double
    Dim Finish As Double
    Dim SaveName As String
    Dim Path As String
    Dim ThisWorkbook As Workbook
    
    ' ���������� ����� ������ ���������� ���������
    Start = Now()
    
    ' ������������� ������ �� �������� �����
    Set ThisWorkbook = ActiveWorkbook
    
    ' ��������� ������: ������� � ExitHandler � ������ ������
    On Error GoTo ExitHandler
    
    ' ���������� ���� "Preferences" � �������� ��� ��� ���������� �� ������ H30
    ThisWorkbook.Sheets("Preferences").Activate
    SaveName = ActiveSheet.Range("H30").Text
    
    ' �������� ���� � �����, ��� ��������� �������� �����
    Path = ThisWorkbook.Path
    
    ' ��������� ���������� ������, �������, ������� �������, ��������� ������ � ��������������
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
    
    ' �������� ��������� ��� ������ ������ � ������-������ ������ �������
    SelectPaleYellowSheets
    
    ' ������������ �������� ���� � PDF
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=Path & "\" & SaveName & ".pdf", _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
    ' ������������ �� ���� "Preferences"
    Sheets("Preferences").Select
    
    ' ��������� ����� ���������� ��������� � ��������
    Finish = (Now() - Start) * 24 * 60 * 60
    
ExitHandler:
    ' ���������, ���� �� ���������� ��������� ������� ������� (����� 0.1 �������)
    If Finish < 0.1 Then
        MsgBox "������������ ���������. ���� ������."
    Else
'        MsgBox "���� �������� � ������� PDF � �������� �����", vbInformation, "Done"
    End If
    
    ' ��������������� ��������� Excel
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    
    ' ������������ �� ���� "Preferences"
    ThisWorkbook.Sheets("Preferences").Activate
End Sub

Sub SelectPaleYellowSheets()
    Dim ws As Worksheet
    Dim paleYellowSheets As Collection
    Dim i As Integer
    Dim paleYellowColor As Long
    
    ' ������������� ���� ������-�������
    paleYellowColor = 13434879
    
    ' ������� ��������� ��� �������� ���� ������, ������� ����� ��������
    Set paleYellowSheets = New Collection
    
    ' �������� �� ���� ������ � �����
    For Each ws In ThisWorkbook.Sheets
        ' ���������, ����� �� ���� ������-������ ������
        If ws.Tab.Color = paleYellowColor Then
            paleYellowSheets.Add ws.Name, ws.Name
        End If
    Next ws
    
' ���������, ���� �� ���� �� ���� ���� � ���������
    If paleYellowSheets.Count > 0 Then
        ' ���������� ������ ���� �� ���������
        ThisWorkbook.Sheets(paleYellowSheets(1)).Activate
        
        ' �������� ������ �� �����, ������� ��������� � ���������
        For i = 1 To paleYellowSheets.Count
            ThisWorkbook.Sheets(paleYellowSheets(i)).Select Replace:=False
        Next i
    Else
        MsgBox "��� ������ � ������-������ ������ �������."
    End If
End Sub

