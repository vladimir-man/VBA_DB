Option Explicit
#Const DEBUG_MODE = True

Sub ReadSourceData()
    Dim wbSource As Workbook, wbDB As Workbook
    Dim wsIron As Worksheet, wsTitle As Worksheet
    Dim arrSpare As Variant, cols As Variant
    Dim carDonor As String, additionToTable As String
    Dim result() As Variant
    Dim i As Long, k As Long
    


    Set wsTitle = ThisWorkbook.Sheets("накладна отримання")
    Set wsIron = ThisWorkbook.Sheets("в металобрухт")
    carDonor = wsTitle.Range("B4").Value
    additionToTable = "'" & carDonor & "' Акт зміни якісного стану №"

    arrSpare = Range("B22:Q" & Cells(Rows.Count, "B").End(xlUp).Row).Value
   
    cols = Array(1, 2, 5, 10, 11, 12, 14, 15, 16) ' точный порядок столбцов
    ReDim result(1 To UBound(arrSpare, 1), 1 To UBound(cols) + 1)
    ' копируем реальные столбцы
    For k = LBound(cols) To UBound(cols)
        For i = 1 To UBound(arrSpare, 1)
            result(i, k + 1) = arrSpare(i, cols(k))
        Next i
    Next k

    #If DEBUG_MODE Then
    ' Тестовая выгрузка на отдельный лист с заголовками
    Dim wsTest As Worksheet
    On Error Resume Next
    Set wsTest = ThisWorkbook.Sheets("DebugOutput")
    On Error GoTo 0
    If wsTest Is Nothing Then
        Set wsTest = Worksheets.Add
        wsTest.Name = "DebugOutput"
    End If
    wsTest.Cells.Clear

    ' Заголовки для выбранных столбцов
    For k = LBound(cols) To UBound(cols)
        wsTest.Cells(1, k + 1).Value = "Col_" & cols(k)
    Next k

    ' Выгрузка данных под заголовками
    wsTest.Range("A2").Resize(UBound(result, 1), UBound(result, 2)).Value = result
#End If




End Sub



