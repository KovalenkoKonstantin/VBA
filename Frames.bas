Attribute VB_Name = "Frames"
Sub Frames_Insertion_old()

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False


Dim sh As Worksheet
Dim ThisWorkbook As Workbook
Dim new_name As String

Set ThisWorkbook = ActiveWorkbook
new_name1 = ThisWorkbook.Sheets("Preferences").Range("H21").Value2
new_name2 = ThisWorkbook.Sheets("Preferences").Range("H22").Value2

For Each sh In ThisWorkbook.Worksheets
    For i = 1 To 50
        sh.Activate
        On Error Resume Next
        ActiveSheet.Shapes.Range(Array("Rectangle " & i)).Select
            Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = new_name1
        i = i + 1
    Next i
    
    For i = 2 To 50
        sh.Activate
        On Error Resume Next
        ActiveSheet.Shapes.Range(Array("Rectangle " & i)).Select
            Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = new_name2
        i = i + 1
    Next i
Next

ThisWorkbook.Sheets("Preferences").Activate

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True

End Sub

Sub Frames_Insertion()
    ' ���������� ���������� ������ � ������ ������� ��� ��������� ������������������
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False

    Dim sh As Worksheet
    Dim ThisWorkbook As Workbook
    Dim new_name1 As String
    Dim new_name2 As String

    ' ��������� ������� ������� �����
    Set ThisWorkbook = ActiveWorkbook
    
    ' ��������� �������� �� ����� "Preferences"
    new_name1 = ThisWorkbook.Sheets("Preferences").Range("H21").Value2
    new_name2 = ThisWorkbook.Sheets("Preferences").Range("H22").Value2

    ' ������� ���� ������ � ������� �����
    For Each sh In ThisWorkbook.Worksheets
        ' ������� ������ new_name1 � ������ 50 �����
        UpdateShapesText sh, new_name1, 1, 50
        
        ' ������� ������ new_name2 � ������ � 2 �� 50
        UpdateShapesText sh, new_name2, 2, 50
    Next sh

    ' ������� �� ���� "Preferences"
    ThisWorkbook.Sheets("Preferences").Activate

    ' ��������� ���� ����������� ����������
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
End Sub

' ��������� ��� ���������� ������ ����� �� �����
Private Sub UpdateShapesText(sh As Worksheet, textValue As String, startIndex As Integer, endIndex As Integer)
    Dim i As Integer
    For i = startIndex To endIndex
        On Error Resume Next ' ������������ ������, ���� ������ �� �������
        sh.Shapes("Rectangle " & i).TextFrame2.TextRange.Characters.Text = textValue
        On Error GoTo 0 ' �������� ��������� ������ �����
    Next i
End Sub
