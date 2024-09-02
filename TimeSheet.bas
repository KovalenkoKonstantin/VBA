Attribute VB_Name = "TimeSheet"
Sub TimeSheet()
    ' ������ ���������
    Start = Now() ' ���������� ������� ����� ��� ��������� ����������������� ����������

    Dim FilesToOpen ' ���������� ��� �������� ���� � ����������� ������
    Dim ThisWorkbook As Workbook, importWB As Workbook ' ���������� ��� ������� �����
    Dim SheetName As String ' ���������� ��� �������� ����� �����
    Dim ws As Worksheet ' ���������� ��� ������ � �������

    Set ThisWorkbook = ActiveWorkbook ' ������������� ������ ThisWorkbook �� ������� �������� �����
    On Error GoTo ExitHandler ' ������������ ������, ��������� ����� ���������� � ExitHandler ��� ������������� ������
    SheetName = "������" ' ��� �����, � ������� ����� ��������
    awLastCol = 63 ' ��������� ������� ��� ��������

    ' ��������� ���������� ������ � ������ ������� ��� ��������� ����������
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False

    ' ������ ���� ������� � ���������� ���
    ThisWorkbook.Sheets(SheetName).Visible = True
    ThisWorkbook.Sheets(SheetName).Activate

    ' ��������� ������ ������ ������
    FilesToOpen = Application.GetOpenFilename _
        (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
        MultiSelect:=True, Title:="�������� ���� � �������� ��� ��������������")

    ' ���� ������ �� �������, ������� �� ���������
    If TypeName(FilesToOpen) = "Boolean" Then
        GoTo ExitHandler
    End If

    ThisWorkbook.Sheets(SheetName).Activate ' ���������� ���� ��� ���������� ������
    On Error Resume Next ' ���������� ������ �� ��������� �������
    ActiveSheet.ShowAllData ' ���������� ��� ������, ���� ������� �������

    ' ������ ������ �� ������� ���������� �����
    Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

    On Error Resume Next ' ����� ���������� ������

    importWB.Sheets(1).Activate ' ���������� ������ ���� ������������� �����

    ' ������� ���������� ������ �� �����
    ThisWorkbook.Sheets(SheetName).Activate
    awLastRow = Cells(Rows.Count, "AC").End(xlUp).row ' ������� ����� ��������� ����������� ������ � ������� AC
    Range(Cells(1, 1), Cells(awLastRow, awLastCol)).Select ' �������� �������� ��� �������
    With Selection
        .Clear ' ������� ��������� ��������
    End With

    ' ������ ����� ������ �� ���� ������ ��������������� �����
    For Each ws In importWB.Sheets ' ��� ������� ����� � ������������� �����
        ws.Activate ' ���������� ������� ����
        iwLastRow = Cells(Rows.Count, "AC").End(xlUp).row ' ������� ����� ��������� ����������� ������ � ������� AC

        importWB.Activate
        Range(Cells(1, 1), Cells(iwLastRow, awLastCol)).Copy ' �������� ������ � �������� �����

        ThisWorkbook.Sheets(SheetName).Activate ' ������������ � ��������� �����
        awFirstRow = Cells(Rows.Count, "AC").End(xlUp).row ' ������� ����� ��������� ����������� ������ �� ������� �����
        awFirstCol = 1 ' ������ �������, ������� � 1

        ' ��������� ������������� ������ � ������ ������ ���
        Range(Cells(awFirstRow + 1, awFirstCol), Cells(iwLastRow + awFirstRow, awLastCol)).Select
        With Selection
            .PasteSpecial Paste:=xlPasteAll ' ��������� ������������� ������
        End With
    Next ws

    ' ��������� ����, �� �������� ������������� ������
    importWB.Close
    ThisWorkbook.Sheets(SheetName).Activate ' ����� ���������� �������� ����

    ' ���������� ��������� ����� ����������
    ' MsgBoxEx "������� ������� ���������", 0, "�����������", 15

ExitHandler: ' ���������� ������ �� ���������
    ' �������� ������� ��������� ����������, ������� ���� ���������
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ' ���� ��� ����� ������ ����, �� ���������������
    ' ThisWorkbook.Sheets(SheetName).Visible = False
    ThisWorkbook.Sheets("Preferences").Activate ' ���������� ���� "Preferences"
    
    ' ������������ � ���������� ����� ���������� (����������������)
    ' Finish = (Now() - Start) * 24 * 60 * 60
    ' MsgBox (Finish)

    Exit Sub ' ��������� ���������
    
ErrHandler: ' ���������� ������
    MsgBox Err.Description ' ���� ��������� ������, ���������� ��������� � ���������
    Resume ExitHandler ' ������������ � ExitHandler ��� �������
End Sub
