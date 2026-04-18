Sub ExportToWordDynamicSummary()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim tbl As Object
    Dim metals As Variant
    Dim values As Variant
    Dim j As Long
    
    ' Запуск Word
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    
    ' Диалог выбора шаблона Word
    Dim fdOpen As FileDialog
    Set fdOpen = Application.FileDialog(msoFileDialogFilePicker)
    fdOpen.Title = "Выберите шаблон Word"
    fdOpen.Filters.Clear
    fdOpen.Filters.Add "Word Documents", "*.docx"
    If fdOpen.Show = -1 Then
        Set wdDoc = wdApp.Documents.Open(fdOpen.SelectedItems(1))
    Else
        MsgBox "Шаблон не выбран!"
        Exit Sub
    End If
    
    ' Берём данные из Excel
    Set ws = ThisWorkbook.Sheets("Дані")
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    
    ' Ссылка на первую таблицу в документе
    Set tbl = wdDoc.Tables(1)
    
    ' Вставляем строки данных перед итогами
For i = 2 To lastRow
    tbl.Rows.Add tbl.Rows(tbl.Rows.Count) ' добавляем перед "Всього"
    
    tbl.Cell(i, 1).Range.Text = ws.Cells(i, 1).Value
    tbl.Cell(i, 2).Range.Text = ws.Cells(i, 2).Value
    tbl.Cell(i, 3).Range.Text = ws.Cells(i, 3).Value
    tbl.Cell(i, 4).Range.Text = ws.Cells(i, 4).Value
    tbl.Cell(i, 5).Range.Text = ws.Cells(i, 5).Value
    tbl.Cell(i, 6).Range.Text = ws.Cells(i, 6).Value
    tbl.Cell(i, 7).Range.Text = ws.Cells(i, 7).Value
    tbl.Cell(i, 8).Range.Text = ws.Cells(i, 8).Value
    tbl.Cell(i, 9).Range.Text = ws.Cells(i, 9).Value
    tbl.Cell(i, 10).Range.Text = ws.Cells(i, 10).Value
    
    ' Сброс ориентации для всех ячеек вставленной строки
    Dim c As Long
    For c = 1 To tbl.Columns.Count
        With tbl.Cell(i, c).Range
            .Orientation = 0
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
        End With
    Next c
Next i

    
    
    Dim lastSummaryRow As Long
    Dim r As Long
    
    ' Определяем последнюю строку с данными по металлам в Excel
    lastSummaryRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    
    ' Переносим ненулевые строки в конец таблицы Word
    For r = 2 To lastSummaryRow
        If ws.Cells(r, "M").Value > 0 Then
            ' Добавляем строку в конец таблицы
            tbl.Rows.Add
            
            ' Название металла во второй столбец
            tbl.Cell(tbl.Rows.Count, 2).Range.Text = ws.Cells(r, "L").Value
            
            ' Значение в последний столбец
            tbl.Cell(tbl.Rows.Count, tbl.Columns.Count).Range.Text = ws.Cells(r, "M").Value
                ' Сброс ориентации для значения
            With tbl.Cell(tbl.Rows.Count, tbl.Columns.Count).Range
                .Orientation = 0
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
            End With
        End If
    Next r
    
                Dim col As Long
    Dim rowIndex As Long
    Dim topCell As Object
    
    ' Определяем последнюю строку с данными (до блока металлов)
    Dim lastDataRow As Long
    lastDataRow = tbl.Rows.Count - (lastSummaryRow - 1) ' минус количество строк металлов
    
    ' Проходим по каждому столбцу таблицы, только до lastDataRow
    For col = 1 To tbl.Columns.Count
        Set topCell = Nothing
        For rowIndex = 1 To lastDataRow
            With tbl.Cell(rowIndex, col)
                ' Проверяем, пустая ли ячейка
                If Trim(.Range.Text) = Chr(13) & Chr(7) Then
                    If Not topCell Is Nothing Then
                        topCell.Merge tbl.Cell(rowIndex, col)
                    End If
                Else
                    ' Запоминаем верхнюю непустую ячейку
                    Set topCell = tbl.Cell(rowIndex, col)
                End If
            End With
        Next rowIndex
    Next col



    
    wdDoc.Bookmarks("DatePlace").Range.Text = Format(Date, "dd.mm.yyyy")




    
    Dim fdSave As FileDialog
    Dim fileName As String
    
    fileName = "Заявка металобрухт_" & Format(Date, "dd_mmmm_yyyy") & ".docx"
    
    Set fdSave = wdApp.FileDialog(msoFileDialogSaveAs)
    fdSave.Title = "Сохранить документ"
    fdSave.InitialFileName = fileName
    
    If fdSave.Show = -1 Then
        wdDoc.SaveAs2 fdSave.SelectedItems(1), FileFormat:=wdFormatXMLDocument
    Else
        MsgBox "Файл не сохранён!"
    End If



    
    ' Очистка
    Set tbl = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub


