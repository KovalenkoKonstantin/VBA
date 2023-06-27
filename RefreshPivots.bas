Attribute VB_Name = "RefreshPivots"
Sub create()
Dim Folder As String
Dim Path As String
Dim Slash As String
Dim Object As Object
Dim ThisWorkbook As Workbook

Set ThisWorkbook = ActiveWorkbook
Set Object = CreateObject("Scripting.FileSystemObject")

ThisWorkbook.Sheets("Merge").Activate
Folder = ActiveSheet.Range("AB2").Text
Path = ActiveWorkbook.Path

    If Object.FolderExists(Path & "\" & Folder) Then
        Object.DeleteFolder (Path & "\" & Folder)
        Object.CreateFolder (Path & "\" & Folder)
    Else
        Object.CreateFolder (Path & "\" & Folder)
    End If

End Sub

Sub «аблокировать¬сеЋисты()
Dim ws As Worksheet
 Application.ScreenUpdating = False
 Application.EnableEvents = False
 ActiveSheet.DisplayPageBreaks = False
 Application.DisplayStatusBar = False
 Application.DisplayAlerts = False
For Each ws In ActiveWorkbook.Worksheets
ws.Protect Password:="gfhjkm"
Next ws

'ThisWorkbook.Sheets("Merge").Activate
' ActiveWorkbook.Protect Password:="gfhjkm"
' Application.ScreenUpdating = True
' Application.EnableEvents = True
' ActiveSheet.DisplayPageBreaks = True
' Application.DisplayStatusBar = True
' Application.DisplayAlerts = True

End Sub

Sub —н€ть«ащиту¬сехЋистов()
Dim ws As Worksheet
 Application.ScreenUpdating = False
 Application.EnableEvents = False
 ActiveSheet.DisplayPageBreaks = False
 Application.DisplayStatusBar = False
 Application.DisplayAlerts = False
For Each ws In ActiveWorkbook.Worksheets
ws.UnProtect Password:="gfhjkm"
Next ws
ActiveWorkbook.UnProtect Password:="gfhjkm"
 Application.ScreenUpdating = True
 Application.EnableEvents = True
 ActiveSheet.DisplayPageBreaks = True
 Application.DisplayStatusBar = True
 Application.DisplayAlerts = True
'ThisWorkbook.Sheets("Merge").Activate
End Sub

Sub ќбновить()
 
 Dim ws As Worksheet
 Dim pt As PivotTable
' —н€ть«ащиту¬сехЋистов
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
    For Each pt In ws.PivotTables
    pt.RefreshTable
    Next pt
    Next ws
Application.StatusBar = "¬ыполнено 40%"
    For Each ws In ThisWorkbook.Worksheets
    For Each pt In ws.PivotTables
    pt.RefreshTable
    Next pt
    Next ws
Application.StatusBar = "¬ыполнено 80%"
    For Each ws In ThisWorkbook.Worksheets
    For Each pt In ws.PivotTables
    pt.RefreshTable
    Next pt
    Next ws
Application.StatusBar = False
' «аблокировать¬сеЋисты
' ThisWorkbook.Sheets("Merge").Activate

End Sub

Sub delete()
 Dim ws As Worksheet
 Dim pt As PivotTable
 Dim ThisWorkbook As Workbook
 Dim rCell As Range
 On Error GoTo ErrHandler
 Set ThisWorkbook = ActiveWorkbook
 
 Application.ScreenUpdating = False
 —н€ть«ащиту¬сехЋистов
  
 ThisWorkbook.Sheets("58").Activate
 Range("A1:BB300").Select
 With Selection
        .Clear
 End With
 
 ThisWorkbook.Sheets("58н").Activate
 Range("A1:BB300").Select
 With Selection
        .Clear
 End With
 
 ThisWorkbook.Sheets("58контр").Activate
 Range("A1:BB300").Select
 With Selection
        .Clear
 End With
 
 ThisWorkbook.Sheets("60").Activate
 Range("A1:BB300").Select
 With Selection
        .Clear
 End With
 
 ThisWorkbook.Sheets("60н").Activate
 Range("A1:BB300").Select
 With Selection
        .Clear
 End With
 
 ThisWorkbook.Sheets("60контр").Activate
 Range("A1:BB300").Select
 With Selection
        .Clear
 End With
 
 ThisWorkbook.Sheets("62").Activate
 Range("A1:BB300").Select
 With Selection
        .Clear
 End With
 
 ThisWorkbook.Sheets("62н").Activate
 Range("A1:BB300").Select
 With Selection
        .Clear
 End With
 
 ThisWorkbook.Sheets("62контр").Activate
 Range("A1:BB300").Select
 With Selection
        .Clear
 End With
 
 ThisWorkbook.Sheets("66").Activate
 Range("A1:BB300").Select
 With Selection
        .Clear
 End With
 
 ThisWorkbook.Sheets("66н").Activate
 Range("A1:BB300").Select
 With Selection
        .Clear
 End With
 
 ThisWorkbook.Sheets("66контр").Activate
 Range("A1:BB300").Select
 With Selection
        .Clear
 End With
 
 ThisWorkbook.Sheets("76").Activate
 Range("A1:BB300").Select
 With Selection
        .Clear
 End With
 
 ThisWorkbook.Sheets("76н").Activate
 Range("A1:BB300").Select
 With Selection
        .Clear
 End With
 
 ThisWorkbook.Sheets("76контр").Activate
 Range("A1:BB300").Select
 With Selection
        .Clear
 End With
 
For Each ws In ThisWorkbook.Worksheets
For Each pt In ws.PivotTables
pt.RefreshTable
Next pt
Next ws

For Each ws In ThisWorkbook.Worksheets
For Each pt In ws.PivotTables
pt.RefreshTable
Next pt
Next ws

 ThisWorkbook.Sheets("Merge").Activate
 
 
ExitHandler:
 Application.ScreenUpdating = True
 ThisWorkbook.Sheets("Merge").Activate
 Exit Sub
 
 «аблокировать¬сеЋисты

ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub



