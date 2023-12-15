Attribute VB_Name = "GetNames"
Public Sub GetDistinctNames()
    Dim FullNameColumn As Range
    Dim dimension As Integer
    Dim ThisWorkbook, importWB As Workbook
    Dim DistinctList As Variant
    
    'Set FullNameColumn = ActiveSheet.UsedRange.Columns(4) ' Получаем первый столбец.
    
    Set ThisWorkbook = ActiveWorkbook
    SheetName = "РВ"
    ThisWorkbook.Sheets(SheetName).Activate
    Range("B4:B103").ClearContents
    
FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="Выберите данные по трудоёмкости")

Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

On Error Resume Next
importWB.Activate
Set FullNameColumn = ActiveSheet.UsedRange.Range("D5:D100") ' Диапазон значений.

    DistinctList = GetDistinctItems(FullNameColumn) ' Передаем диапазон в функцию.
'    Debug.Print Join(DistinctList, vbCrLf) ' Выводим результат.
    dimension = UBound(DistinctList, 1) 'размер массива

ThisWorkbook.Sheets(SheetName).Activate
  
  'идём циклом по элементам массива
  For i = 1 To dimension
  'предполагаю что длина имени не может быть меньше 5 знаков
    If Len(DistinctList(i)) > 5 Then
    'записываю элементы в новый массив
      j = j + 1
'    DistinctList(j) = DistinctList(i)
'    Debug.Print DistinctList(j)

'вывожу элементы на лист в ячейку B4
    Range("B" & i + 3).Value2 = DistinctList(j)
    End If
    
  Next
'    Debug.Print DistinctList
    
'    Range("E1").Resize(j + 1, 1) = WorksheetFunction.Transpose(DistinctList)
'    DistinctList = Range(Range("E1"), Range("E" & Rows.Count).End(xlUp)).Resize(, 1).Value
'    Range("E1").Resize(dimension - 2, 1).Select
'    Range("E1").Resize(dimension, 1).Value2 = WorksheetFunction.Transpose(DistinctList)
End Sub

Public Function GetDistinctItems(ByRef Range As Range) As Variant
    Dim Data As Variant: Data = Range.Value ' Преобразуем диапазон в массив.
    Dim Buffer As Object: Set Buffer = CreateObject("System.Collections.ArrayList") ' Создаем объект ArrayList.

    Dim Item
    For Each Item In Data
        If Not Buffer.Contains(Item) Then Buffer.Add Item ' Проверяем наличие элемента и добавляем если отсутствует.
    Next

    Buffer.Sort ': Buffer.Reverse ' Сортируем по возрастанию, а потом переворачиваем (по убыванию).
    GetDistinctItems = Buffer.ToArray() ' Выгружаем в виде массива.
End Function
