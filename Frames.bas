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
    ' Отключение обновления экрана и других событий для повышения производительности
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False

    Dim sh As Worksheet
    Dim ThisWorkbook As Workbook
    Dim new_name1 As String
    Dim new_name2 As String

    ' Установка текущей рабочей книги
    Set ThisWorkbook = ActiveWorkbook
    
    ' Получение значений из листа "Preferences"
    new_name1 = ThisWorkbook.Sheets("Preferences").Range("H21").Value2
    new_name2 = ThisWorkbook.Sheets("Preferences").Range("H22").Value2

    ' Перебор всех листов в рабочей книге
    For Each sh In ThisWorkbook.Worksheets
        ' Вставка текста new_name1 в первые 50 фигур
        UpdateShapesText sh, new_name1, 1, 50
        
        ' Вставка текста new_name2 в фигуры с 2 по 50
        UpdateShapesText sh, new_name2, 2, 50
    Next sh

    ' Возврат на лист "Preferences"
    ThisWorkbook.Sheets("Preferences").Activate

    ' Включение всех отключенных параметров
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
End Sub

' Процедура для обновления текста фигур на листе
Private Sub UpdateShapesText(sh As Worksheet, textValue As String, startIndex As Integer, endIndex As Integer)
    Dim i As Integer
    For i = startIndex To endIndex
        On Error Resume Next ' Игнорировать ошибки, если фигура не найдена
        sh.Shapes("Rectangle " & i).TextFrame2.TextRange.Characters.Text = textValue
        On Error GoTo 0 ' Включить обработку ошибок снова
    Next i
End Sub
