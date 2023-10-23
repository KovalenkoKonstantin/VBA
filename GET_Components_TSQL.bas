Attribute VB_Name = "GET_Components_TSQL"


Sub Query1_Add_()

  Dim i As Long
  Dim list As Worksheet
  Dim ThisWorkbook As Workbook
  Dim SheetName, var As String
  Dim Sht As Worksheet
  Dim X As Range
  Set ThisWorkbook = ActiveWorkbook
  var = ThisWorkbook.Sheets("Труд").Range("I2").Value2

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
    
    ActiveWorkbook.Queries.Add Name:="Components", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Источник = Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec ComponentsRefresh 'Программно-аппаратный комплекс ViPNet Coordinator HW50 A 4.x (+3G)(+u%';""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Components;Extended Properties=""""" _
        , Destination:=Range("$Q$4")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Components]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "Components"
        .Refresh BackgroundQuery:=False
    End With
    
Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True

End Sub
Sub Components_SP_Query_()
Dim ThisWorkbook As Workbook
Dim var As String
Set ThisWorkbook = ActiveWorkbook
var = ThisWorkbook.Sheets("Труд").Range("I3").Value2

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual

ThisWorkbook.Sheets("Труд").Activate
Range("Q5:W40").Select
    Selection.ClearContents

ActiveWorkbook.Queries("Components").Formula = "let Источник = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"ComponentsRefresh '" & var & "';""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"
'"let Источник = " & _
'"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
'"LabourRefresh '" & var & "'""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"
        

ActiveWorkbook.Queries("Components").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
ThisWorkbook.Sheets("Preferences").Activate

End Sub


