Attribute VB_Name = "Get_Contractors_TSQ"
Sub GetContractorsRefresh_SP_Query()
Dim ThisWorkbook As Workbook
Dim var As String
Set ThisWorkbook = ActiveWorkbook

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual

ActiveWorkbook.Queries("Contractors").Formula = "let Источник = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetContractorsRefresh;""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"

'ActiveWorkbook.RefreshAll
ActiveWorkbook.Queries("Contractors").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
'ThisWorkbook.Sheets("Preferences").Activate

End Sub

