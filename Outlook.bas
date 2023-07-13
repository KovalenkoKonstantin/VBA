Attribute VB_Name = "Outlook"
Sub Message()
    Application.StatusBar = "Создание запроса"
    RunPython ("import OutlookMsg; OutlookMsg.Messaging()")
    Application.StatusBar = False
End Sub
