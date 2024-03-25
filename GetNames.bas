Attribute VB_Name = "GetNames"
Public Sub GetDistinctNames()

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
'Application.Calculation = xlManual

    Dim FullNameColumn As Range
    Dim dimension As Integer
    Dim ThisWorkbook, importWB As Workbook
    Dim DistinctList As Variant
    Dim FilesToOpen
    
FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="Выберите данные по трудоёмкости")
    
    'Set FullNameColumn = ActiveSheet.UsedRange.Columns(4) ' Получаем первый столбец.
    Set ThisWorkbook = ActiveWorkbook
    SheetName = "РВ"
    ThisWorkbook.Sheets(SheetName).Activate
    Range("B4:B103").ClearContents
    
Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

On Error Resume Next
importWB.Activate
Set FullNameColumn = ActiveSheet.UsedRange.Range("D5:D150") ' Диапазон значений.

    DistinctList = GetDistinctItems(FullNameColumn) ' Передаем диапазон в функцию.
'    Debug.Print Join(DistinctList, vbCrLf) ' Выводим результат.
    dimension = UBound(DistinctList, 1) 'размер массива

ThisWorkbook.Sheets(SheetName).Activate
    
  j = -1
  'идём циклом по элементам массива
  For i = 0 To dimension
  'предполагаю что длина имени не может быть меньше 5 знаков
    If Len(DistinctList(i)) > 5 Then
    'записываю элементы в новый массив
      j = j + 1
    DistinctList(j) = DistinctList(i)
    Debug.Print DistinctList(j)

'вывожу элементы на лист в ячейку B4
    Range("B" & i + 4).Value2 = DistinctList(j)
    End If
  Next
  
importWB.Close
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
'    Application.Calculation = xlAutomatic
 ThisWorkbook.Sheets("Preferences").Activate
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
