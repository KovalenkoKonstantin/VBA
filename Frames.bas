Attribute VB_Name = "Frames"
Sub Frames_Insertion()

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