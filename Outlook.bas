Attribute VB_Name = "Outlook"
Sub Message()
    Application.StatusBar = "�������� �������"
    RunPython ("import OutlookMsg; OutlookMsg.Messaging()")
    Application.StatusBar = False
End Sub
