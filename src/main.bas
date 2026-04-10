Option Explicit
#Const DEBUG_MODE = True

Public result() As Variant
Public additionToTable As String

Sub MainProcess()
    Dim modeChoice As Integer
    
    ' Окно выбора режима
    modeChoice = MsgBox("Створити новий файл на основі шаблону?", _
                        vbYesNoCancel + vbQuestion, "Вибір режиму")
    
    If modeChoice = vbCancel Then Exit Sub
    
    ' Сбор данных
    Call ReadSourceData
    
    ' Экспорт
    Call ExportToTemplate(result, modeChoice, additionToTable)
End Sub

Sub ReadSourceData()
    Dim wsTitle As Worksheet, wsIron As Worksheet
    Dim arrSpare As Variant, cols As Variant
    Dim carDonor As String
    Dim temp() As Variant
    Dim i As Long, j As Long, k As Long
    
    Set wsTitle = ThisWorkbook.Sheets("накладна отримання")
    Set wsIron = ThisWorkbook.Sheets("в металобрухт")
    
    carDonor = wsTitle.Range("B4").Value
    additionToTable = "'" & carDonor & "' Акт зміни якісного стану №"
    
    arrSpare = wsIron.Range("B22:Q" & wsIron.Cells(wsIron.Rows.Count, "B").End(xlUp).Row).Value
    cols = Array(1, 2, 5, 10, 11, 12, 14, 15, 16)
    
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
End Sub

Sub ExportToTemplate(result As Variant, modeChoice As Integer, additionToTable As String)
    Dim wbTarget As Workbook
    Dim wsTarget As Worksheet
    Dim filePath As Variant
    Dim savePath As Variant
    Dim lastRow As Long
    Dim startRow As Long, endRow As Long
    
    ' Выбор файла
    filePath = Application.GetOpenFilename( _
        FileFilter:="Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", _
        Title:="Выберите файл")
    
    If filePath = False Then Exit Sub
    Set wbTarget = Workbooks.Open(filePath)
    Set wsTarget = wbTarget.Sheets("Дані")
    
    ' Проверка на пустой массив
    If Not IsEmpty(result) Then
        If modeChoice = vbYes Then
            ' Новый акт > вставляем с A2
            startRow = 2
            endRow = startRow + UBound(result, 1) - 1
            wsTarget.Range("A" & startRow).Resize(UBound(result, 1), UBound(result, 2)).Value = result
            Call ClearAndPrepareSheet(wsTarget, additionToTable, startRow, endRow)
        ElseIf modeChoice = vbNo Then
            ' Дополнение > ищем первую пустую строку по колонке I
            lastRow = wsTarget.Cells(wsTarget.Rows.Count, "I").End(xlUp).Row
            If Application.CountA(wsTarget.Rows(lastRow)) > 0 Then
                lastRow = lastRow + 1
            End If
            
            startRow = lastRow
            endRow = startRow + UBound(result, 1) - 1
            wsTarget.Range("A" & startRow).Resize(UBound(result, 1), UBound(result, 2)).Value = result
            Call ClearAndPrepareSheet(wsTarget, additionToTable, startRow, endRow)
        End If
    End If
    
    ' Сохранить как
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="Акт_" & Format(Now, "yyyymmdd_hhmm") & ".xlsm", _
        FileFilter:="Excel Macro-Enabled Workbook (*.xlsm), *.xlsm")
    
    If savePath <> False Then
        wbTarget.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    End If
    
    wbTarget.Close SaveChanges:=False
End Sub

Private Sub ClearAndPrepareSheet(wsTarget As Worksheet, additionToTable As String, startRow As Long, endRow As Long)
    ' Объединяем только диапазон текущего массива
    With wsTarget.Range("J" & startRow & ":J" & endRow)
        .Merge
        .Value = additionToTable
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Bold = True
        .Font.Size = 12
        .Interior.Color = RGB(220, 230, 241)
    End With
    
    wsTarget.Columns("J").ColumnWidth = 27
End Sub


