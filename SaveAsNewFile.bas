Attribute VB_Name = "SaveAsNewFile"
Sub SaveFinalTableAsNewFile_v_1_0()
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


Sub SaveFinalTableAsNewFile_v_1_1()
    Dim wsSource As Worksheet
    Dim tblSelected As ListObject
    Dim newFilePath As String
    Dim currentDate As String
    Dim currentTime As String
    Dim newWorkbook As Workbook
    Dim newWorksheet As Worksheet
    Dim userChoice As String
    Dim folderPath As String
    
    ' ���� � ����� ��������
    folderPath = ThisWorkbook.Path & "\��������\"
    
    ' ���������, ���������� �� ����� "��������"
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "����� '��������' �� ������� � ������� ����������.", vbExclamation
        Exit Sub
    End If
    
    ' �������� ���������� ���� � ����� ��������: ��� � ��
    userChoice = MsgBox("�������� ���: ��� ��� ��." & vbCrLf & _
                        "������� '��' ��� ���, '���' ��� ��", vbYesNo + vbQuestion, "����� ����")
    
    ' �������� ������ ������������
    If userChoice = vbYes Then
        ' ������ ���
        userChoice = "���"
    ElseIf userChoice = vbNo Then
        ' ������ ��
        userChoice = "��"
    Else
        MsgBox "�������� ��������.", vbInformation
        Exit Sub
    End If
    
    ' ������������� ������ �� ���� "��� ��������"
    Set wsSource = ThisWorkbook.Sheets("��� ��������")
    
    ' ����������, ����� ������� �������� � ����������� �� ������ ������������
    If userChoice = "���" Then
        ' ������� "Final_���" ��� ���
        On Error Resume Next
        Set tblSelected = wsSource.ListObjects("Final_���")
        On Error GoTo 0
        If tblSelected Is Nothing Then
            MsgBox "������� 'Final_���' �� �������!", vbExclamation
            Exit Sub
        End If
    ElseIf userChoice = "��" Then
        ' ������� "Final_��" ��� ��
        On Error Resume Next
        Set tblSelected = wsSource.ListObjects("Final_��")
        On Error GoTo 0
        If tblSelected Is Nothing Then
            MsgBox "������� 'Final_��' �� �������!", vbExclamation
            Exit Sub
        End If
    End If
    
    ' �������� ������� ���� � ����� � ������ �������
    currentDate = Format(Date, "DD.MM.YYYY")
    currentTime = Format(Time, "HH-MM")
    
    ' ��������� ��� ������ ����� � ����������� �� ������
    newFilePath = folderPath & "��� �������� " & userChoice & " (" & currentDate & " " & currentTime & ").xlsx"
    
    ' �������� �� ������������� ����� � ��� ��������, ���� ���� ��� ����������
    If Dir(newFilePath) <> "" Then
        ' ���� ����������, ������� ���
        Kill newFilePath
    End If
    
    ' ������ ����� Workbook
    Set newWorkbook = Workbooks.Add
    ' ��������� ����� ���� � ����� ����
    Set newWorksheet = newWorkbook.Sheets(1)
    
    ' �������� ������ �� ��������� ������� � ����� ����
    tblSelected.Range.Copy Destination:=newWorksheet.Cells(1, 1)
    
    ' ������� �������������� ������ (���� ��� ����), ������ ��������
    newWorksheet.Cells.Copy
    newWorksheet.Cells.PasteSpecial Paste:=xlPasteValues
    
    ' ��������� ����� ���� � �� �� ���������� (����� ��������)
    newWorkbook.SaveAs Filename:=newFilePath, FileFormat:=xlOpenXMLWorkbook
    
    ' ��������� ����� ����
    newWorkbook.Close SaveChanges:=False
    
    ' ������� ��������� � ����������
'    MsgBox "���� ������� ���: " & newFilePath, vbInformation
End Sub

Sub SaveFinalTableAsNewFile()
    Dim wsSource As Worksheet
    Dim tblSelected As ListObject
    Dim newFilePath As String
    Dim currentDate As String
    Dim currentTime As String
    Dim newWorkbook As Workbook
    Dim newWorksheet As Worksheet
    Dim folderPath As String
    
    ' ���������� ����� ��� ������ ����
    �����.Show
'    Debug.Print (�����.userChoice)
    ' ���� ����� ���� �������, ������������ �� ������ ������
    If �����.userChoice = "" Then
        MsgBox "�������� ��������.", vbInformation
        Exit Sub
    End If
    
    ' ���� � ����� "��������"
    folderPath = ThisWorkbook.Path & "\��������\"
    
    ' ���������, ���������� �� ����� "��������"
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "����� '��������' �� ������� � ������� ����������.", vbExclamation
        Exit Sub
    End If
    
    ' ������������� ������ �� ���� "��� ��������"
    Set wsSource = ThisWorkbook.Sheets("��� ��������")
    
    ' ����������, ����� ������� �������� � ����������� �� ������ ������������
    If �����.userChoice = "���" Then
        ' ������� "Final_���" ��� ���
        On Error Resume Next
        Set tblSelected = wsSource.ListObjects("Final_���")
        On Error GoTo 0
        If tblSelected Is Nothing Then
            MsgBox "������� 'Final_���' �� �������!", vbExclamation
            Exit Sub
        End If
    ElseIf �����.userChoice = "��" Then
        ' ������� "Final_��" ��� ��
        On Error Resume Next
        Set tblSelected = wsSource.ListObjects("Final_��")
        On Error GoTo 0
        If tblSelected Is Nothing Then
            MsgBox "������� 'Final_��' �� �������!", vbExclamation
            Exit Sub
        End If
    End If
    
    ' �������� ������� ���� � ����� � ������ �������
    currentDate = Format(Date, "DD.MM.YYYY")
    currentTime = Format(Time, "HH-MM")
    
    ' ��������� ��� ������ ����� � ����������� �� ������
    newFilePath = folderPath & "��� �������� " & �����.userChoice & " (" & currentDate & " " & currentTime & ").xlsx"
    
    ' �������� �� ������������� ����� � ��� ��������, ���� ���� ��� ����������
    If Dir(newFilePath) <> "" Then
        ' ���� ����������, ������� ���
        Kill newFilePath
    End If
    
    ' ������ ����� Workbook
    Set newWorkbook = Workbooks.Add
    ' ��������� ����� ���� � ����� ����
    Set newWorksheet = newWorkbook.Sheets(1)
    
    ' �������� ������ �� ��������� ������� � ����� ����
    tblSelected.Range.Copy Destination:=newWorksheet.Cells(1, 1)
    
    ' ������� �������������� ������ (���� ��� ����), ������ ��������
    newWorksheet.Cells.Copy
    newWorksheet.Cells.PasteSpecial Paste:=xlPasteValues
    
    ' ��������� ����� ���� � �� �� ���������� (����� ��������)
    newWorkbook.SaveAs Filename:=newFilePath, FileFormat:=xlOpenXMLWorkbook
    
    ' ��������� ����� ����
    newWorkbook.Close SaveChanges:=False
    
    ' ������� ��������� � ����������
    MsgBox "���� ������� ���: " & newFilePath, vbInformation
End Sub

