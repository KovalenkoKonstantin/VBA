Attribute VB_Name = "SaveAsNewFile"
Sub SaveFinalTableAsNewFile()
    Dim wsSource As Worksheet
    Dim tblFinal As ListObject
    Dim newFilePath As String
    Dim currentDate As String
    Dim currentTime As String
    Dim newWorkbook As Workbook
    Dim newWorksheet As Worksheet
    
    ' ������������� ������ �� ���� "��� ��������"
    Set wsSource = ThisWorkbook.Sheets("��� ��������")
    
    ' ������� ������� "Final" �� ���� �����
    Set tblFinal = wsSource.ListObjects("Final")
    
    ' �������� ������� ���� � ����� � ������ �������
    currentDate = Format(Date, "DD.MM.YYYY")
    currentTime = Format(Time, "HH-MM")
    
    ' ��������� ��� ������ �����
    newFilePath = ThisWorkbook.Path & "\��� �������� ��� (" & currentDate & " " & currentTime & ").xlsx"
    
    ' ������ ����� Workbook
    Set newWorkbook = Workbooks.Add
    ' ��������� ����� ���� � ����� ����
    Set newWorksheet = newWorkbook.Sheets(1)
    
    ' �������� ������ �� ������� "Final" � ����� ����
    tblFinal.Range.Copy Destination:=newWorksheet.Cells(1, 1)
    
    ' ������� �������������� ������ (���� ��� ����), ������ ��������
    newWorksheet.Cells.Copy
    newWorksheet.Cells.PasteSpecial Paste:=xlPasteValues
    
    ' ��������� ����� ���� � �� �� ����������
    newWorkbook.SaveAs Filename:=newFilePath, FileFormat:=xlOpenXMLWorkbook
    
    ' ��������� ����� ����
    newWorkbook.Close SaveChanges:=False
    
'    ' ������� ��������� � ����������
'    MsgBox "���� ������� ���: " & newFilePath, vbInformation
End Sub

