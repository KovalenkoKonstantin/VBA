Attribute VB_Name = "Outlook"
Sub Message() 'запрос на поддержку 1С ERP
    Application.StatusBar = "Создание запроса"
    RunPython ("import OutlookMsg; OutlookMsg.Messaging()")
    Application.StatusBar = False
End Sub

Sub Request() 'технологи
    Application.StatusBar = "Создание запроса"
    RunPython ("import Request; Request.MessagingRequest()")
    Application.StatusBar = False
End Sub

Sub DenisRequest() 'запрос Денису
    Application.StatusBar = "Создание запроса"
    RunPython ("import Denis; Denis.MessagingRequest()")
    Application.StatusBar = False
End Sub

Sub Negotiation() 'согласование расчёта
    Application.StatusBar = "Создание запроса"
    RunPython ("import negotiation; negotiation.MessagingRequest()")
    Application.StatusBar = False
End Sub

Sub WSS_MSG() 'доработка WSS
    Application.StatusBar = "Создание запроса"
    RunPython ("import WSS; WSS.MessagingRequest()")
    Application.StatusBar = False
End Sub
