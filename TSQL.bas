Attribute VB_Name = "TSQL"
Sub Components_SP_Query_()
Dim ThisWorkbook As Workbook
Dim var As String
Set ThisWorkbook = ActiveWorkbook
var = ThisWorkbook.Sheets("����").Range("I3").Value2

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual

ThisWorkbook.Sheets("����").Activate
Range("Q5:W40").Select
    Selection.ClearContents

ActiveWorkbook.Queries("Components").Formula = "let �������� = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"ComponentsRefresh '" & var & "';""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    ��������"
'"let �������� = " & _
'"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
'"LabourRefresh '" & var & "'""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    ��������"
        

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

ActiveWorkbook.Queries("Contractors").Formula = "let �������� = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetContractorsRefresh;""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    ��������"

ActiveWorkbook.Queries("Contractors").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic

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

ActiveWorkbook.Queries("Employee").Formula = "let �������� = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetEmployeeRefresh '" & var & "';""])" & Chr(13) & "" & _
"" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    ��������"

ActiveWorkbook.Queries("Employee").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic

End Sub
Sub GetEmployeeChangesRefresh_SP_Query()
Dim ThisWorkbook As Workbook
Dim var As String
Set ThisWorkbook = ActiveWorkbook

var = ThisWorkbook.Sheets("ForDataBase").Range("CC1").Value2
var1 = ThisWorkbook.Sheets("ForDataBase").Range("CD1").Value2

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

ActiveWorkbook.Queries("EmployeeChanges").Formula = "let �������� = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetEmployeeChanges " & var & "," & var1 & ";""])" & Chr(13) & "" & _
"" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    ��������"

ActiveWorkbook.Queries("EmployeeChanges").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic

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

ActiveWorkbook.Queries("Enterprise").Formula = "let �������� = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetEnterpriseRefresh '" & var & "';""])" & Chr(13) & "" & _
"" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    ��������"

ActiveWorkbook.Queries("Enterprise").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic

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

ActiveWorkbook.Queries("Expenditures").Formula = "let �������� = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetExpendituresRefresh;""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    ��������"

ActiveWorkbook.Queries("Expenditures").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic

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

ActiveWorkbook.Queries("GozAttribute").Formula = "let �������� = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetGozAttributeRefresh;""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    ��������"

ActiveWorkbook.Queries("GozAttribute").Refresh

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
var = ThisWorkbook.Sheets("����").Range("I3").Value2

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual

ThisWorkbook.Sheets("����").Activate
Range("N5:O40").Select
    Selection.ClearContents

ActiveWorkbook.Queries("Operations").Formula = "let �������� = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"LabourRefresh '" & var & "'""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    ��������"
       
ActiveWorkbook.Queries("Operations").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic

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

ActiveWorkbook.Queries("Organization").Formula = "let �������� = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetOrganizationRefresh;""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    ��������"

ActiveWorkbook.Queries("Organization").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic

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

ActiveWorkbook.Queries("Project").Formula = "let �������� = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetProjectRefresh;""])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    ��������"

ActiveWorkbook.Queries("Project").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True

End Sub
Sub GetTaxRefresh_SP_Query()
Dim ThisWorkbook As Workbook
Dim var As String
Set ThisWorkbook = ActiveWorkbook

var = ThisWorkbook.Sheets("ForDataBase").Range("BN1").Value2
var1 = ThisWorkbook.Sheets("ForDataBase").Range("BO1").Value2

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

ActiveWorkbook.Queries("Tax").Formula = "let �������� = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetTaxRefresh " & var & ", " & var1 & ";""])" & Chr(13) & "" & _
"" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    ��������"

ActiveWorkbook.Queries("Tax").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True

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

ActiveWorkbook.Queries("TaxBase").Formula = "let �������� = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetTaxBaseRefresh '" & var & "';""])" & Chr(13) & "" & _
"" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    ��������"

ActiveWorkbook.Queries("TaxBase").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True

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

ActiveWorkbook.Queries("Worktime").Formula = "let �������� = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetWorktimeRefresh '" & var & "';""])" & Chr(13) & "" & _
"" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    ��������"

ActiveWorkbook.Queries("Worktime").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic

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

ActiveWorkbook.Queries("SalaryBudget").Formula = "let �������� = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetSalaryBudgetRefresh " & var & "," & var1 & ";""])" & Chr(13) & "" & _
"" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    ��������"

ActiveWorkbook.Queries("SalaryBudget").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True

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

ActiveWorkbook.Queries("Salary").Formula = "let �������� = " & _
"Sql.Database(""msk-sql-02"", ""RKM"", [Query=""exec " & _
"GetSalary " & var & "," & var1 & ";""])" & Chr(13) & "" & _
"" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    ��������"

ActiveWorkbook.Queries("Salary").Refresh

Application.StatusBar = False
Application.ScreenUpdating = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True

End Sub
