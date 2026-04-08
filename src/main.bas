Option Explicit
#Const DEBUG_MODE = True

Sub ReadSourceData()
    Dim wbSource As Workbook, wbDB As Workbook
    Dim wsIron As Worksheet, wsTitle As Worksheet
    Dim arrSpare As Variant, cols As Variant
    Dim carDonor As String, additionToTable As String
    Dim result() As Variant
    Dim i As Long, k As Long, j As Long
    Dim temp() As Variant
    Dim startRow As Long, lastRow As Long
    Dim modeChoice As Integer

    


    Set wsTitle = ThisWorkbook.Sheets("накладна отримання")
    Set wsIron = ThisWorkbook.Sheets("в металобрухт")
    carDonor = wsTitle.Range("B4").Value
    additionToTable = "'" & carDonor & "' Акт зміни якісного стану №"

    arrSpare = wsIron.Range("B22:Q" & wsIron.Cells(wsIron.Rows.Count, "B").End(xlUp).Row).Value
    
    cols = Array(1, 2, 5, 10, 11, 12, 14, 15, 16) ' точный порядок столбцов
    ReDim result(1 To UBound(arrSpare, 1), 1 To UBound(cols) + 1)
    ' копируем реальные столбцы
   ' Временный массив с запасом
    ReDim temp(1 To UBound(arrSpare, 1), 1 To UBound(cols) + 1)
    
    k = 0
    For i = 1 To UBound(arrSpare, 1)
        If UBound(arrSpare, 2) >= 11 Then
            If Not IsEmpty(arrSpare(i, 11)) Then
                If InStr(1, arrSpare(i, 11), "Утіль", vbTextCompare) = 0 Then
                    k = k + 1
                    For j = LBound(cols) To UBound(cols)
                        temp(k, j + 1) = arrSpare(i, cols(j))
                    Next j
                End If
            End If
        End If
    Next i
    
    ' Теперь создаём итоговый массив точного размера
    If k > 0 Then
        ReDim result(1 To k, 1 To UBound(cols) + 1)
        For i = 1 To k
            For j = 1 To UBound(cols) + 1
                result(i, j) = temp(i, j)
            Next j
        Next i
    Else
        Erase result
    End If
    'For i = 1 To UBound(result, 1)
    'For j = 1 To UBound(result, 2)
     '   Debug.Print "Row " & i & ", Col " & j & " = " & result(i, j)
    'Next j
'Next i

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
    
     modeChoice = vbNo
    
    ' Определяем последнюю строку после вставки новых данных
    lastRow = wsTest.Cells(wsTest.Rows.Count, 9).End(xlUp).Row
    If modeChoice = vbYes Then
    ' Пользователь выбрал "Да" > создать новый файл
    ElseIf modeChoice = vbNo Then
    ' Пользователь выбрал "Нет" > открыть существующий
    startRow = wsTest.Cells(2, 10).Row
    ElseIf modeChoice = vbCancel Then
    ' Пользователь нажал "Отмена" > выход из макроса
    Exit Sub
    End If
    
    With wsTest.Range("J" & startRow & ":J" & lastRow)
    .Merge
    .Value = additionToTable
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = True
    End With
    ' Устанавливаем ширину столбца J (примерно 205 пикселей ? 27 символов)
    wsTest.Columns("J").ColumnWidth = 27
    #End If
    
   





End Sub




