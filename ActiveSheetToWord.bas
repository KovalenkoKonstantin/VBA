Attribute VB_Name = "ActiveSheetToWord"
Sub ExportActiveSheetToWord()
    Dim ws As Worksheet
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim printArea As Range
    Dim filePath As String
    Dim fileName As String

    ' ������������� ������ �� �������� ����
    Set ws = ThisWorkbook.ActiveSheet

    ' ���������� ������� ������
    On Error Resume Next
    Set printArea = ws.Range(ws.PageSetup.printArea)
    If printArea Is Nothing Then
        Set printArea = ws.UsedRange ' ���� ������� ������ �� ������, ���������� ��� ������� ������
    End If
    On Error GoTo 0
    
    ' ������� ���������� Word
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True ' ������ Word �������

    ' ������� ����� ��������
    Set wordDoc = wordApp.Documents.Add

    ' �������� ������� ������ �� Excel
    printArea.Copy

    ' ��������� � �������� Word
    wordDoc.Content.Paste

    ' ���������� ���� ����������
    filePath = ThisWorkbook.Path
    fileName = ws.Name & ".docx" ' ��� ����� �� �������� �����

    ' ��������� ��������
    wordDoc.SaveAs filePath & "\" & fileName
    wordDoc.Close
    wordApp.Quit

    ' ����������� ������
    Set wordDoc = Nothing
    Set wordApp = Nothing
    
    ' ����������� ������������
'    MsgBox "�������� ��������: " & filePath & "\" & fileName
End Sub
