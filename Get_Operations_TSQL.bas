Attribute VB_Name = "Get_Operations_TSQL"
Sub Query1_Add()

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
Application.Calculation = xlManual
    
    ActiveWorkbook.Queries.Add Name:="Operations", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    " & _
        "Источник = Sql.Database(""msk-sql-02"", ""RKM"", " & _
        "[Query=""select operation_name, labour_intensity_month_value#(lf)" & _
        "from LabourIntensity l#(lf)inner join Operations O on " & _
        "l.operation_id = O.operation_id#(lf)inner join Project P on " & _
        "l.project_id = P.project_id#(lf)where project_cipher " & _
        "like#(lf)      " & _
        "'Программно-аппаратный комплекс ViPNet Coordinator HW50 A 4.x (+3G)(+u%'" & _
        ";""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник" & _
        ""

    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Operations;Extended Properties=""""" _
        , Destination:=Range("$N$4")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Operations]")
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
        .ListObject.DisplayName = "Operations"
        .Refresh BackgroundQuery:=False
    End With
    
Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic

End Sub
Sub LabourIntensityQuery()
Dim ThisWorkbook As Workbook
Dim var As String
Set ThisWorkbook = ActiveWorkbook
var = ThisWorkbook.Sheets("Труд").Range("I2").Value2

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual


ActiveWorkbook.Queries("Operations").Formula = "let Источник = Sql.Database(""msk-sql-02"", ""RKM"", " & _
        "[Query=""select operation_name, labour_intensity_month_value#(lf)" & _
        "from LabourIntensity l#(lf)inner join Operations O on " & _
        "l.operation_id = O.operation_id#(lf)inner join Project P on " & _
        "l.project_id = P.project_id#(lf)where project_cipher " & _
        "like#(lf)      " & _
        "'" & var & "'" & _
        ";""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник" & _
        ""
        
ActiveWorkbook.Queries("Operations").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic

End Sub

Sub LabourIntensity_SP_Query()
Dim ThisWorkbook As Workbook
Dim var As String
Set ThisWorkbook = ActiveWorkbook
var = ThisWorkbook.Sheets("Труд").Range("I2").Value2

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual

ThisWorkbook.Sheets("Труд").Activate
Range("N5:O40").Select
    Selection.ClearContents

ActiveWorkbook.Queries("Operations").Formula = "let Источник = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"LabourRefresh '" & var & "'""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"
       
ActiveWorkbook.Queries("Operations").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
ThisWorkbook.Sheets("Preferences").Activate

End Sub

