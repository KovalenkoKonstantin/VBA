Attribute VB_Name = "Outlook"
Sub Message() '������ �� ��������� 1� ERP
    Application.StatusBar = "�������� �������"
    RunPython ("import OutlookMsg; OutlookMsg.Messaging()")
    Application.StatusBar = False
End Sub

Sub Request() '���������
    Application.StatusBar = "�������� �������"
    RunPython ("import Request; Request.MessagingRequest()")
    Application.StatusBar = False
End Sub

Sub DenisRequest() '������ ������
    Application.StatusBar = "�������� �������"
    RunPython ("import Denis; Denis.MessagingRequest()")
    Application.StatusBar = False
End Sub

Sub Negotiation() '������������ �������
    Application.StatusBar = "�������� �������"
    RunPython ("import negotiation; negotiation.MessagingRequest()")
    Application.StatusBar = False
End Sub

Sub WSS_MSG() '��������� WSS
    Application.StatusBar = "�������� �������"
    RunPython ("import WSS; WSS.MessagingRequest()")
    Application.StatusBar = False
End Sub
