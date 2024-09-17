VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim dropdownCell As Range
    Dim conditionCell As Range


    ' Определяем ячейку с выпадающим списком и ячейку условия
    Set dropdownCell = Me.Range("E3") ' Ячейка с выпадающим списком
    Set conditionCell = Me.Range("E1") ' Ячейка для проверки условия

    ' Проверяем, изменена ли ячейка с выпадающим списком
    If Not Intersect(Target, dropdownCell) Is Nothing Then
        ' Проверяем значение в ячейке E1
        If conditionCell.Value = False Then
            ' Если значение в E1 равно False, выполняем действие
            RefreshAllTables
        End If
    End If
    
    ' Наполняем диапазон с помощью автофила
    With Sheets("Расчет АО ""ПМ""")
                .Range("B4").AutoFill Destination:=.Range("B4:B500"), Type:=xlFillDefault
    End With
    Sheets("БД").Select
    
End Sub
