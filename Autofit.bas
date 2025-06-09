Attribute VB_Name = "Autofit"
Sub SetRowHeightToContent()
    Dim ws As Worksheet
    Dim mergedCell As Range
    Dim tempCell As Range
    Dim originalHeight As Double
    
    ' ������� ���� � ������ "3. �����"
    Set ws = ThisWorkbook.Worksheets("3. �����")
    ws.Activate
    
    ' ������������� ������ �� ������������ ������
    Set mergedCell = ws.Range("A31:E31")
    
    ' �������� ������� ������
    mergedCell.WrapText = True
    
    ' ��������� ������������ ������
    originalHeight = ws.Rows(31).RowHeight
    
    ' �������� ���������� ����� ������� ������ ��� �������
    ws.Rows(31).RowHeight = 300
    
    ' ������� ��������� ������ ��� �������
    Set tempCell = ws.Range("Z31")
    tempCell.Value = mergedCell.Value
    tempCell.Font.Size = mergedCell.Font.Size
    tempCell.Font.Name = mergedCell.Font.Name
    tempCell.WrapText = True
    tempCell.ColumnWidth = (mergedCell.Width / 7.5) ' ������������ � ������� ������ �������
    
    ' ��������� ���������� � ��������� ������
    tempCell.EntireRow.Autofit
    
    ' �������� ������������ ������
    Dim calculatedHeight As Double
    calculatedHeight = ws.Rows(31).RowHeight
    
    ' ������� ��������� ������
    tempCell.Clear
    
    ' ��������� ������������ ������ � ������ ������
    ws.Rows(31).RowHeight = calculatedHeight
    
End Sub
