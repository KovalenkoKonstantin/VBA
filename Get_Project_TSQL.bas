Attribute VB_Name = "Get_Project_TSQL"
'Sub Query3_Add()
'
'    ActiveWorkbook.Queries.Add Name:="Query3", Formula:= _
'        "let" & Chr(13) & "" & Chr(10) & "    Источник = " & _
'        "Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
'        "GetProjectRefresh;""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"
'    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
'        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Query3;Extended Properties=""""" _
'        , Destination:=Range("$AD$2")).QueryTable
'        .CommandType = xlCmdSql
'        .CommandText = Array("SELECT * FROM [Query3]")
'        .RowNumbers = False
'        .FillAdjacentFormulas = False
'        .PreserveFormatting = True
'        .RefreshOnFileOpen = False
'        .BackgroundQuery = True
'        .RefreshStyle = xlInsertDeleteCells
'        .SavePassword = False
'        .SaveData = True
'        .AdjustColumnWidth = True
'        .RefreshPeriod = 0
'        .PreserveColumnInfo = True
'        .ListObject.DisplayName = "Query3"
'        .Refresh BackgroundQuery:=False
'    End With
''    Range("V11").Select
''    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
'End Sub

Sub GetProjectRefresh_SP_Query()
Dim ThisWorkbook As Workbook
Dim var As String
Set ThisWorkbook = ActiveWorkbook

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual

ActiveWorkbook.Queries("Project").Formula = "let Источник = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetProjectRefresh;""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"

'ActiveWorkbook.RefreshAll
ActiveWorkbook.Queries("Project").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
'ThisWorkbook.Sheets("Preferences").Activate

End Sub
