Attribute VB_Name = "Get_Organization_TSQL"
Sub GetOrganizationRefresh_SP_Query()
Dim ThisWorkbook As Workbook
Dim var As String
Set ThisWorkbook = ActiveWorkbook

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual

ActiveWorkbook.Queries("Organization").Formula = "let Источник = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetOrganizationRefresh;""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"

'ActiveWorkbook.RefreshAll
ActiveWorkbook.Queries("Organization").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
'ThisWorkbook.Sheets("Preferences").Activate

End Sub

