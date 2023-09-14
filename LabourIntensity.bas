Attribute VB_Name = "LabourIntensity"
Sub Макрос2()
Attribute Макрос2.VB_ProcData.VB_Invoke_Func = " \n14"

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

'    Range("Query1[#All]").Select
'    Selection.ListObject.QueryTable.delete
'    Selection.ClearContents
    
    ActiveWorkbook.Queries.Add Name:="Query1", Formula:= _
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

'    ActiveWorkbook.Queries.Add Name:="Query1", Formula:= _
'        "let" & Chr(13) & "" & Chr(10) & "    " & _
'        "Источник = Sql.Database(""msk-sql-02"", ""RKM"", " & _
'        "[Query=""select operation_name, labour_intensity_month_value#(lf)" & _
'        "from LabourIntensity l#(lf)inner join Operations O on " & _
'        "l.operation_id = O.operation_id#(lf)inner join Project P on " & _
'        "l.project_id = P.project_id#(lf)where project_cipher " & _
'        "like#(lf)      " & _
'        "var" & _
'        ";""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник" & _
'        ""

'    ActiveWorkbook.Queries.Add Name:="Query1", Formula:= _
'        "let" & Chr(13) & "" & Chr(10) & "    " & _
'        "Источник = Sql.Database(""msk-sql-02"", ""RKM"", " & _
'        "[Query=""select operation_name, labour_intensity_month_value#(lf)" & _
'        "from LabourIntensity l#(lf)inner join Operations O on " & _
'        "l.operation_id = O.operation_id#(lf)inner join Project P on " & _
'        "l.project_id = P.project_id#(lf)where project_cipher " & _
'        "like#(lf)      " & _
'        "var" & _
'        ";""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник" & _
'        ""

    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Query1;Extended Properties=""""" _
        , Destination:=Range("$I$4")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Query1]")
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
        .ListObject.DisplayName = "Query1"
        .Refresh BackgroundQuery:=False
    End With

    With Selection
        .HorizontalAlignment = xlGeneral
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
'    With Selection
'        .HorizontalAlignment = xlGeneral
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ShrinkToFit = False
'        .ReadingOrder = xlContext
'        .MergeCells = False
'    End With
    With Selection.Font
        .Name = "Times New Roman"
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
'    With Selection.Font
'        .Name = "Times New Roman"
'        .Size = 12
'        .Strikethrough = False
'        .Superscript = False
'        .Subscript = False
'        .OutlineFont = False
'        .Shadow = False
'        .Underline = xlUnderlineStyleNone
'        .ThemeColor = xlThemeColorLight1
'        .TintAndShade = 0
'        .ThemeFont = xlThemeFontNone
'    End With
    
Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True

End Sub
Sub LabourIntensityQuery()
Attribute LabourIntensityQuery.VB_ProcData.VB_Invoke_Func = " \n14"
Dim ThisWorkbook As Workbook
Dim var As String
Set ThisWorkbook = ActiveWorkbook
var = ThisWorkbook.Sheets("Труд").Range("I2").Value2

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

'ActiveWorkbook.Queries("Query1").Formula = "let Источник = Sql.Database(""msk-sql-02"", ""RKM"", " & _
'        "[Query=""select operation_name, labour_intensity_month_value#(lf)" & _
'        "from LabourIntensity l#(lf)inner join Operations O on " & _
'        "l.operation_id = O.operation_id#(lf)inner join Project P on " & _
'        "l.project_id = P.project_id#(lf)where project_cipher " & _
'        "like#(lf)      " & _
'        "'Программно-аппаратный комплекс ViPNet Coordinator HW50 A 4.x (+3G)(+u%'" & _
'        ";""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник" & _
'        ""

ActiveWorkbook.Queries("Query1").Formula = "let Источник = Sql.Database(""msk-sql-02"", ""RKM"", " & _
        "[Query=""select operation_name, labour_intensity_month_value#(lf)" & _
        "from LabourIntensity l#(lf)inner join Operations O on " & _
        "l.operation_id = O.operation_id#(lf)inner join Project P on " & _
        "l.project_id = P.project_id#(lf)where project_cipher " & _
        "like#(lf)      " & _
        "'" & var & "'" & _
        ";""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник" & _
        ""
        
ActiveWorkbook.RefreshAll
'ActiveWorkbook.Queries("Query1").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True

End Sub
