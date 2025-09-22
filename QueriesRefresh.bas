Attribute VB_Name = "QueriesRefresh"
Sub RefreshSelectedPowerQueries_Sequential()
    Dim queryNames As Variant
    Dim conn As WorkbookConnection
    Dim connNameClean As String
    Dim i As Long, j As Long
    Dim foundConn As WorkbookConnection
    Dim refreshedCount As Integer
    
    queryNames = Array( _
        "Сотрудники", _
        "Employee", _
        "SalaryBudget", _
        "EmployeeChanges", _
        "Worktime", _
        "Tax", _
        "TaxBase" _
    )
    
    refreshedCount = 0
    
    For i = LBound(queryNames) To UBound(queryNames)
        Set foundConn = Nothing
        ' Ищем в ThisWorkbook.Connections подключение с нужным именем (с учётом префикса)
        For Each conn In ThisWorkbook.Connections
            If Left(conn.Name, 9) = "Запрос — " Then
                connNameClean = Mid(conn.Name, 10)
            Else
                connNameClean = conn.Name
            End If
            
            If connNameClean = queryNames(i) Then
                Set foundConn = conn
                Exit For
            End If
        Next conn
        
        If Not foundConn Is Nothing Then
            Debug.Print "Обновляется: " & foundConn.Name
            On Error Resume Next
            foundConn.Refresh
            On Error GoTo 0
            
            ' Пауза 2 секунды, чтобы дать время завершить обновление (Power Query асинхронный)
            Application.Wait Now + TimeValue("0:00:02")
            DoEvents
            
            refreshedCount = refreshedCount + 1
        Else
            Debug.Print "Не найдено подключение для: " & queryNames(i)
        End If
    Next i

End Sub


Sub ListAllWorkbookConnections()
    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        Debug.Print conn.Name
    Next conn
End Sub
