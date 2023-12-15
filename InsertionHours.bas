Attribute VB_Name = "InsertionHours"
Sub HoursInsertion()

 Dim FilesToOpen
 Dim ThisWorkbook As Workbook
 Dim SheetName, iw As String
 Dim iLinks As Variant
  
 Set ThisWorkbook = ActiveWorkbook
 SheetName = "РВ"

 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual

FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="Выберите данные по трудоёмкости")

Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))
iw = importWB.Name

On Error Resume Next

ThisWorkbook.Sheets(SheetName).Activate
ActiveSheet.ShowAllData

j = "Часы отнесённые на проект"

ColData = Rows(RowData).Find(j, LookIn:=xlValues).column

Cells(4, ColData).Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIFS('" & iw & "'!C6," _
        & "'" & iw & "'!C4,RC[-8]," _
        & "'" & iw & "'!C1,RC[17])"
Range("J4").Copy
    
ThisWorkbook.Sheets(SheetName).Range("J4:J103,J105:J204,J206:J305," _
& "J307:J406,J408:J507,J509:J608,J610:J709,J711:J810,J812:J911," _
& "J913:J1012,J1014:J1113,J1115:J1214,J1216:J1315,J1317:J1416," _
& "J1418:J1517,J1519:J1618,J1620:J1719,J1721:J1820,J1822:J1921," _
& "J1923:J2022,J2024:J2123,J2125:J2224,J2226:J2325").Select

Selection.PasteSpecial Paste:=xlPasteFormulas

'убираем связи с внешней книгой
iLinks = ActiveWorkbook.LinkSources(xlExcelLinks)
If Not IsEmpty(iLinks) Then
    For i = 1 To UBound(iLinks)
        ActiveWorkbook.BreakLink Name:=iLinks(i), Type:=xlExcelLinks
    Next i
End If

    importWB.Close
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    Application.Calculation = xlAutomatic
' ThisWorkbook.Sheets("Preferences").Activate

End Sub
