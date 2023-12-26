Attribute VB_Name = "TSQL"
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
'ThisWorkbook.Sheets("Preferences").Activate

End Sub

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
Sub GetEmployeeRefresh_SP_Query()
Dim ThisWorkbook As Workbook
Dim var As String
Set ThisWorkbook = ActiveWorkbook

var = ThisWorkbook.Sheets("ForDataBase").Range("BK1").Value2

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
'Application.Calculation = xlManual

ActiveWorkbook.Queries("Employee").Formula = "let Источник = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetEmployeeRefresh '" & var & "';""])" & Chr(13) & "" & _
"" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"

'ActiveWorkbook.Queries("Employee").Formula = "let Источник = " & _
'"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
'"GetEmployeeRefresh;""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"

'ActiveWorkbook.RefreshAll
ActiveWorkbook.Queries("Employee").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
'ThisWorkbook.Sheets("Preferences").Activate

End Sub
Sub GetEmployeeChangesRefresh_SP_Query()
Dim ThisWorkbook As Workbook
Dim var As String
Set ThisWorkbook = ActiveWorkbook

var = ThisWorkbook.Sheets("ForDataBase").Range("AY1").Value2

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
'Application.Calculation = xlManual

ActiveWorkbook.Queries("EmployeeChanges").Formula = "let Источник = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetEmployeeChangesRefresh '" & var & "';""])" & Chr(13) & "" & _
"" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"

'ActiveWorkbook.Queries("EmployeeChanges").Formula = "let Источник = " & _
'"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
'"GetEmployeeChangesRefresh;""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"

'ActiveWorkbook.RefreshAll
ActiveWorkbook.Queries("EmployeeChanges").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
'ThisWorkbook.Sheets("Preferences").Activate

End Sub
Sub GetEnterpriseRefresh_SP_Query()
Dim ThisWorkbook As Workbook
Dim var As String
Set ThisWorkbook = ActiveWorkbook

var = ThisWorkbook.Sheets("ForDataBase").Range("AO1").Value2

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual

ActiveWorkbook.Queries("Enterprise").Formula = "let Источник = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetEnterpriseRefresh '" & var & "';""])" & Chr(13) & "" & _
"" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"

'ActiveWorkbook.Queries("Enterprise").Formula = "let Источник = " & _
'"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
'"GetEnterpriseRefresh;""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"

'ActiveWorkbook.RefreshAll
ActiveWorkbook.Queries("Enterprise").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
'ThisWorkbook.Sheets("Preferences").Activate

End Sub
Sub GetExpendituresRefresh_SP_Query()
Dim ThisWorkbook As Workbook
Dim var As String
Set ThisWorkbook = ActiveWorkbook

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual

ActiveWorkbook.Queries("Expenditures").Formula = "let Источник = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetExpendituresRefresh;""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"

'ActiveWorkbook.RefreshAll
ActiveWorkbook.Queries("Expenditures").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
'ThisWorkbook.Sheets("Preferences").Activate

End Sub
Sub GetGozAttributeRefresh_SP_Query()
Dim ThisWorkbook As Workbook
Dim var As String
Set ThisWorkbook = ActiveWorkbook

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual

ActiveWorkbook.Queries("GozAttribute").Formula = "let Источник = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetGozAttributeRefresh;""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"

'ActiveWorkbook.RefreshAll
ActiveWorkbook.Queries("GozAttribute").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
'ThisWorkbook.Sheets("Preferences").Activate

End Sub
Sub LabourIntensity_SP_Query()
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
'ThisWorkbook.Sheets("Preferences").Activate

End Sub
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
Sub GetProjectRefresh_SP_Query()
Dim ThisWorkbook As Workbook
Dim var As String
Set ThisWorkbook = ActiveWorkbook
var = ThisWorkbook.Sheets("ForDataBase").Range("AG1").Value2

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
'Application.Calculation = xlManual
'
'ActiveWorkbook.Queries("Project").Formula = "let Источник = " & _
'"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
'"GetProjectRefresh_SP '" & var & "';""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"

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
'Application.Calculation = xlAutomatic
'ThisWorkbook.Sheets("Preferences").Activate

End Sub
Sub GetTaxRefresh_SP_Query()
Dim ThisWorkbook As Workbook
Dim var As String
Set ThisWorkbook = ActiveWorkbook

var = ThisWorkbook.Sheets("ForDataBase").Range("BN1").Value2

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
'Application.Calculation = xlManual

ActiveWorkbook.Queries("Tax").Formula = "let Источник = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetTaxRefresh '" & var & "';""])" & Chr(13) & "" & _
"" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"

'ActiveWorkbook.Queries("Tax").Formula = "let Источник = " & _
'"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
'"GetTaxRefresh;""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"

'ActiveWorkbook.RefreshAll
ActiveWorkbook.Queries("Tax").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
'Application.Calculation = xlAutomatic
'ThisWorkbook.Sheets("Preferences").Activate

End Sub
Sub GetTaxBaseRefresh_SP_Query()
Dim ThisWorkbook As Workbook
Dim var As String
Set ThisWorkbook = ActiveWorkbook

var = ThisWorkbook.Sheets("ForDataBase").Range("BS1").Value2

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
'Application.Calculation = xlManual

ActiveWorkbook.Queries("TaxBase").Formula = "let Источник = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetTaxBaseRefresh '" & var & "';""])" & Chr(13) & "" & _
"" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"

'ActiveWorkbook.Queries("TaxBase").Formula = "let Источник = " & _
'"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
'"GetTaxBaseRefresh;""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"

'ActiveWorkbook.RefreshAll
ActiveWorkbook.Queries("TaxBase").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
'Application.Calculation = xlAutomatic
'ThisWorkbook.Sheets("Preferences").Activate

End Sub

Sub GetWorktimeRefresh_SP_Query()
Dim ThisWorkbook As Workbook
Dim var As String
Set ThisWorkbook = ActiveWorkbook

var = ThisWorkbook.Sheets("ForDataBase").Range("BE1").Value2

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
'Application.Calculation = xlManual

ActiveWorkbook.Queries("Worktime").Formula = "let Источник = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetWorktimeRefresh '" & var & "';""])" & Chr(13) & "" & _
"" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"

'ActiveWorkbook.Queries("Worktime").Formula = "let Источник = " & _
'"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
'"GetWorktimeRefresh;""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"

'ActiveWorkbook.RefreshAll
ActiveWorkbook.Queries("Worktime").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
'ThisWorkbook.Sheets("Preferences").Activate

End Sub

Sub GetSalaryBudgetRefresh_SP_Query()
Dim ThisWorkbook As Workbook
Dim var As Integer
Dim var1 As Integer
Set ThisWorkbook = ActiveWorkbook

var = ThisWorkbook.Sheets("ForDataBase").Range("AS1").Value2
var1 = ThisWorkbook.Sheets("ForDataBase").Range("AT1").Value2

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
'Application.Calculation = xlManual

ActiveWorkbook.Queries("SalaryBudget").Formula = "let Источник = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetSalaryBudgetRefresh " & var & "," & var1 & ";""])" & Chr(13) & "" & _
"" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"

'ActiveWorkbook.Queries("SalaryBudget").Formula = "let Источник = " & _
'"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
'"GetSalaryBudgetRefresh;""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"

'ActiveWorkbook.RefreshAll
ActiveWorkbook.Queries("SalaryBudget").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
'Application.Calculation = xlAutomatic
'ThisWorkbook.Sheets("Preferences").Activate

End Sub

Sub GetSalaryRefresh_SP_Query()
Dim ThisWorkbook As Workbook
Dim var As Integer
Dim var1 As Integer
Set ThisWorkbook = ActiveWorkbook

var = ThisWorkbook.Sheets("ForDataBase").Range("BX1").Value2
var1 = ThisWorkbook.Sheets("ForDataBase").Range("BY1").Value2

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
'Application.Calculation = xlManual

ActiveWorkbook.Queries("Salary").Formula = "let Источник = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetSalaryBudgetRefresh " & var & "," & var1 & ";""])" & Chr(13) & "" & _
"" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"

'ActiveWorkbook.Queries("SalaryBudget").Formula = "let Источник = " & _
'"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
'"GetSalaryBudgetRefresh;""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    Источник"

'ActiveWorkbook.RefreshAll
ActiveWorkbook.Queries("Salary").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
'Application.Calculation = xlAutomatic
'ThisWorkbook.Sheets("Preferences").Activate

End Sub
