Attribute VB_Name = "Outlook"
Sub Message()
    Application.StatusBar = "Создание запроса"
    RunPython ("import OutlookMsg; OutlookMsg.Messaging()")
    Application.StatusBar = False
End Sub

Sub Request()
    Application.StatusBar = "Создание запроса"
    RunPython ("import Request; Request.MessagingRequest()")
    Application.StatusBar = False
End Sub

Sub DenisRequest()
    Application.StatusBar = "Создание запроса"
    RunPython ("import Denis; Denis.MessagingRequest()")
    Application.StatusBar = False
End Sub
