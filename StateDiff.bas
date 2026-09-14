Attribute VB_Name = "Module1"
Sub БС()
    Dim ws As Worksheet
    Dim wbBS As Workbook
    Dim wbБыло As Workbook
    Dim wbСтало As Workbook
    Dim wsСМР_Было As Worksheet
    Dim wsТМЦ_Было As Worksheet
    Dim wsСМР_Стало As Worksheet
    Dim wsТМЦ_Стало As Worksheet
    Dim lastRow As Long
    Dim copyRange As Range
    Dim targetRow As Long
    Dim tmciRow As Long
    Dim smrRow As Long
    Dim filePath As Variant
    Dim i As Integer
    Dim dataRow As Long
    Dim firstSMРRow As Long
    Dim lastSMРRow As Long
    Dim firstTMCRow As Long
    Dim lastTMCRow As Long
    
    Dim key As String
    Dim insertRow As Long
    Dim r As Long
    Dim newRowCount As Long
    
    Dim foundRow As Long
    Dim код_Стало As String
    Dim код_БС As String
    
    Dim имя_файла_Было As String
    Dim имя_файла_Стало As String
    
    ' Денежный формат (русская локализация, 2 знака после запятой)
    Dim moneyFormat As String
    moneyFormat = "# ##0,00_ ;-# ##0,00"
    
    ' Текущая книга - файл "БС"
    Set wbBS = ThisWorkbook
    
    ' ============ ЧАСТЬ 1: СОЗДАНИЕ ЗАГОЛОВКА ============
    
    ' 1. Создаем или очищаем лист "Было-Стало"
    On Error Resume Next
    Set ws = wbBS.Worksheets("Было-Стало")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = wbBS.Worksheets.Add(After:=wbBS.Worksheets(wbBS.Worksheets.Count))
        ws.Name = "Было-Стало"
    Else
        ws.Cells.Clear
    End If
    
    ' Задаем ширину колонок
    With ws
        .Columns(1).ColumnWidth = 8
        .Columns(2).ColumnWidth = 54
        .Columns(3).ColumnWidth = 7
        .Columns(4).ColumnWidth = 13
        .Columns(5).ColumnWidth = 13
        .Columns(6).ColumnWidth = 13
        .Columns(7).ColumnWidth = 13
        .Columns(8).ColumnWidth = 13
        .Columns(9).ColumnWidth = 13
        .Columns(10).ColumnWidth = 13
    End With
    
    ' 2. Формируем заголовок
    ' 2.1. Строка 1, колонки 1-2 - названия файлов "Было" и "Стало"
    With ws.Range("A1:B1")
        .Merge
        .Value = ""
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.Bold = False
        .Interior.ColorIndex = xlNone
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .RowHeight = 44
    End With
    
    ' 2.2. Строка 1, колонки 3-6 - "БЫЛО"
    With ws.Range("C1:F1")
        .Merge
        .Value = "БЫЛО"
        .Font.Name = "Arial"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.ColorIndex = 19
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 44
    End With
    
    ' 2.3. Строка 1, колонки 7-9 - "СТАЛО"
    With ws.Range("G1:I1")
        .Merge
        .Value = "СТАЛО"
        .Font.Name = "Arial"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.ColorIndex = 35
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 44
    End With
    
    ' 2.4. Строка 1, колонка 10 - "РАЗНИЦА"
    With ws.Range("J1")
        .Value = "РАЗНИЦА"
        .Font.Name = "Arial"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.ColorIndex = 3
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 44
    End With
    
    ' 2.5. Строка 2 - заголовки колонок
    Dim headers As Variant
    headers = Array("ИД.КЕР/ТМЦ", "Наименование КЕР/ТМЦ", "ЕдИзм КЕР/ТМЦ", _
                    "Объем СМР/ТМЦ Всего в СР", "СМР/ТМЦ за ЕдИзм", "СМР/ТМЦ Всего", _
                    "Объем СМР/ТМЦ Всего в СР", "СМР/ТМЦ за ЕдИзм", "СМР/ТМЦ Всего", _
                    "Стоимость всего руб с НДС")
    
    For i = 1 To 10
        With ws.Cells(2, i)
            .Value = headers(i - 1)
            .Interior.ColorIndex = 15
            .Font.Name = "Aptos Narrow"
            .Font.Size = 8
            .WrapText = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next i
    ws.Rows(2).RowHeight = 34
    
    ' 2.6. Строка 3 - ВСЕГО СМР+ТМЦ
    smrRow = 3
    For i = 1 To 10
        With ws.Cells(smrRow, i)
            .Interior.ColorIndex = 42
            .Font.Name = "Aptos Narrow"
            .Font.Size = 10
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next i
    ws.Cells(smrRow, 2).Value = "ВСЕГО СМР+ТМЦ"
    ws.Rows(smrRow).RowHeight = 13
    
    ' 2.7. Строка 4 - СМР
    smrRow = 4
    For i = 1 To 10
        With ws.Cells(smrRow, i)
            .Interior.ColorIndex = 37
            .Font.Name = "Aptos Narrow"
            .Font.Size = 10
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next i
    ws.Cells(smrRow, 2).Value = "СМР"
    ws.Rows(smrRow).RowHeight = 13
    
    ' ============ ЧАСТЬ 2: ЗАГРУЗКА ДАННЫХ ИЗ ФАЙЛА "БЫЛО" ============
    
    ' 1. Запрашиваем путь к файлу "Было"
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Укажите путь к файлу БЫЛО"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            MsgBox "Файл не выбран. Операция отменена."
            Exit Sub
        End If
    End With
    
    ' Сохраняем имя файла "Было"
    имя_файла_Было = Right(filePath, Len(filePath) - InStrRev(filePath, "\"))
    
    ' Открываем файл "Было"
    Application.ScreenUpdating = False
    Set wbБыло = Workbooks.Open(filePath)
    
    ' 2. Копируем данные с листа "СМР уник"
    On Error Resume Next
    Set wsСМР_Было = wbБыло.Worksheets("СМР уник")
    On Error GoTo 0
    
    If wsСМР_Было Is Nothing Then
        MsgBox "В файле ""Было"" нет листа ""СМР уник""!"
        wbБыло.Close False
        Exit Sub
    End If
    
    ' Определяем последнюю непустую строку на листе СМР уник (по колонке 1)
    lastRow = wsСМР_Было.Cells(wsСМР_Было.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "На листе ""СМР уник"" нет данных для копирования!"
        wbБыло.Close False
        Exit Sub
    End If
    
    ' Копируем данные (колонки 1-5, со 2-й строки)
    Set copyRange = wsСМР_Было.Range(wsСМР_Было.Cells(2, 1), wsСМР_Было.Cells(lastRow, 5))
    
    ' 3. Вставляем данные после строки "СМР" (строка 4)
    targetRow = 5
    firstSMРRow = targetRow
    
    ' Вставляем значения
    ws.Range(ws.Cells(targetRow, 1), ws.Cells(targetRow + copyRange.Rows.Count - 1, 5)).Value = copyRange.Value
    
    ' Применяем форматирование
    With ws.Range(ws.Cells(targetRow, 1), ws.Cells(targetRow + copyRange.Rows.Count - 1, 5))
        .Font.Name = "Aptos Narrow"
        .Font.Size = 8
    End With
    
    ' Форматирование по колонкам
    ws.Range(ws.Cells(targetRow, 1), ws.Cells(targetRow + copyRange.Rows.Count - 1, 1)).HorizontalAlignment = xlCenter
    With ws.Range(ws.Cells(targetRow, 2), ws.Cells(targetRow + copyRange.Rows.Count - 1, 2))
        .HorizontalAlignment = xlLeft
        .WrapText = True
    End With
    ws.Range(ws.Cells(targetRow, 3), ws.Cells(targetRow + copyRange.Rows.Count - 1, 3)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(targetRow, 4), ws.Cells(targetRow + copyRange.Rows.Count - 1, 6)).HorizontalAlignment = xlRight
    
    ' Запоминаем последнюю строку СМР
    lastSMРRow = targetRow + copyRange.Rows.Count - 1
    
    ' 4. Добавляем строку "ТМЦ"
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    targetRow = lastRow + 1
    tmciRow = targetRow
    
    ' Заливка только строки ТМЦ
    ws.Range(ws.Cells(targetRow, 1), ws.Cells(targetRow, 10)).Interior.ColorIndex = 37
    ws.Rows(targetRow).RowHeight = 13
    
    ' Заполняем ячейки
    With ws.Cells(targetRow, 2)
        .Value = "ТМЦ"
        .Font.Name = "Aptos Narrow"
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ws.Range(ws.Cells(targetRow, 1), ws.Cells(targetRow, 10)).VerticalAlignment = xlCenter
    ws.Range(ws.Cells(targetRow, 1), ws.Cells(targetRow, 1)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(targetRow, 3), ws.Cells(targetRow, 10)).HorizontalAlignment = xlCenter
    
    ' 5. Копируем данные с листа "ТМЦ уник"
    On Error Resume Next
    Set wsТМЦ_Было = wbБыло.Worksheets("ТМЦ уник")
    On Error GoTo 0
    
    If wsТМЦ_Было Is Nothing Then
        MsgBox "В файле ""Было"" нет листа ""ТМЦ уник""!"
        wbБыло.Close False
        Exit Sub
    End If
    
    lastRow = wsТМЦ_Было.Cells(wsТМЦ_Было.Rows.Count, 1).End(xlUp).Row
    
    If lastRow >= 2 Then
        Set copyRange = wsТМЦ_Было.Range(wsТМЦ_Было.Cells(2, 1), wsТМЦ_Было.Cells(lastRow, 5))
        
        targetRow = tmciRow + 1
        firstTMCRow = targetRow
        
        ws.Range(ws.Cells(targetRow, 1), ws.Cells(targetRow + copyRange.Rows.Count - 1, 5)).Value = copyRange.Value
        
        With ws.Range(ws.Cells(targetRow, 1), ws.Cells(targetRow + copyRange.Rows.Count - 1, 5))
            .Font.Name = "Aptos Narrow"
            .Font.Size = 8
        End With
        
        ws.Range(ws.Cells(targetRow, 1), ws.Cells(targetRow + copyRange.Rows.Count - 1, 1)).HorizontalAlignment = xlCenter
        With ws.Range(ws.Cells(targetRow, 2), ws.Cells(targetRow + copyRange.Rows.Count - 1, 2))
            .HorizontalAlignment = xlLeft
            .WrapText = True
        End With
        ws.Range(ws.Cells(targetRow, 3), ws.Cells(targetRow + copyRange.Rows.Count - 1, 3)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(targetRow, 4), ws.Cells(targetRow + copyRange.Rows.Count - 1, 6)).HorizontalAlignment = xlRight
        
        lastTMCRow = targetRow + copyRange.Rows.Count - 1
    Else
        MsgBox "На листе ""ТМЦ уник"" нет данных для копирования!"
        lastTMCRow = tmciRow
    End If
    
    wbБыло.Close False
    
    ' ============ ЧАСТЬ 3: РАСЧЕТЫ В КОЛОНКАХ 6, 9 И 10 ============
    
    ' 3.1. Формулы для строк СМР
    For dataRow = firstSMРRow To lastSMРRow
        ws.Cells(dataRow, 6).Formula = "=ROUND(D" & dataRow & "*E" & dataRow & ", 2)"
        ws.Cells(dataRow, 6).NumberFormatLocal = moneyFormat
        ws.Cells(dataRow, 6).Font.Name = "Aptos Narrow"
        ws.Cells(dataRow, 6).Font.Size = 8
        
        ws.Cells(dataRow, 9).Formula = "=ROUND(G" & dataRow & "*H" & dataRow & ", 2)"
        ws.Cells(dataRow, 9).NumberFormatLocal = moneyFormat
        ws.Cells(dataRow, 9).Font.Name = "Aptos Narrow"
        ws.Cells(dataRow, 9).Font.Size = 8
        
        ws.Cells(dataRow, 10).Formula = "=I" & dataRow & "-F" & dataRow
        ws.Cells(dataRow, 10).NumberFormatLocal = moneyFormat
        ws.Cells(dataRow, 10).Font.Name = "Aptos Narrow"
        ws.Cells(dataRow, 10).Font.Size = 8
    Next dataRow
    
    ' 3.2. Формулы для строк ТМЦ
    If lastTMCRow > tmciRow Then
        For dataRow = firstTMCRow To lastTMCRow
            ws.Cells(dataRow, 6).Formula = "=ROUND(D" & dataRow & "*E" & dataRow & ", 2)"
            ws.Cells(dataRow, 6).NumberFormatLocal = moneyFormat
            ws.Cells(dataRow, 6).Font.Name = "Aptos Narrow"
            ws.Cells(dataRow, 6).Font.Size = 8
            
            ws.Cells(dataRow, 9).Formula = "=ROUND(G" & dataRow & "*H" & dataRow & ", 2)"
            ws.Cells(dataRow, 9).NumberFormatLocal = moneyFormat
            ws.Cells(dataRow, 9).Font.Name = "Aptos Narrow"
            ws.Cells(dataRow, 9).Font.Size = 8
            
            ws.Cells(dataRow, 10).Formula = "=I" & dataRow & "-F" & dataRow
            ws.Cells(dataRow, 10).NumberFormatLocal = moneyFormat
            ws.Cells(dataRow, 10).Font.Name = "Aptos Narrow"
            ws.Cells(dataRow, 10).Font.Size = 8
        Next dataRow
    End If
    
    ' 3.3. Суммы
    If firstSMРRow <= lastSMРRow Then
        ws.Cells(4, 6).Formula = "=SUM(F" & firstSMРRow & ":F" & lastSMРRow & ")"
        ws.Cells(4, 6).NumberFormatLocal = moneyFormat
        ws.Cells(4, 6).Font.Name = "Aptos Narrow"
        ws.Cells(4, 6).Font.Size = 10
        
        ws.Cells(4, 9).Formula = "=SUM(I" & firstSMРRow & ":I" & lastSMРRow & ")"
        ws.Cells(4, 9).NumberFormatLocal = moneyFormat
        ws.Cells(4, 9).Font.Name = "Aptos Narrow"
        ws.Cells(4, 9).Font.Size = 10
    End If
    
    If firstTMCRow <= lastTMCRow Then
        ws.Cells(tmciRow, 6).Formula = "=SUM(F" & firstTMCRow & ":F" & lastTMCRow & ")"
        ws.Cells(tmciRow, 6).NumberFormatLocal = moneyFormat
        
        ws.Cells(tmciRow, 9).Formula = "=SUM(I" & firstTMCRow & ":I" & lastTMCRow & ")"
        ws.Cells(tmciRow, 9).NumberFormatLocal = moneyFormat
    End If
    
    ws.Cells(4, 10).Formula = "=I4-F4"
    ws.Cells(4, 10).NumberFormatLocal = moneyFormat
    ws.Cells(4, 10).Font.Name = "Aptos Narrow"
    ws.Cells(4, 10).Font.Size = 10
    
    ws.Cells(tmciRow, 10).Formula = "=I" & tmciRow & "-F" & tmciRow
    ws.Cells(tmciRow, 10).NumberFormatLocal = moneyFormat
    ws.Cells(tmciRow, 10).Font.Name = "Aptos Narrow"
    ws.Cells(tmciRow, 10).Font.Size = 10
    
    ws.Cells(3, 6).Formula = "=F4+F" & tmciRow
    ws.Cells(3, 6).NumberFormatLocal = moneyFormat
    ws.Cells(3, 6).Font.Name = "Aptos Narrow"
    ws.Cells(3, 6).Font.Size = 10
    
    ws.Cells(3, 9).Formula = "=I4+I" & tmciRow
    ws.Cells(3, 9).NumberFormatLocal = moneyFormat
    ws.Cells(3, 9).Font.Name = "Aptos Narrow"
    ws.Cells(3, 9).Font.Size = 10
    
    ws.Cells(3, 10).Formula = "=J4+J" & tmciRow
    ws.Cells(3, 10).NumberFormatLocal = moneyFormat
    ws.Cells(3, 10).Font.Name = "Aptos Narrow"
    ws.Cells(3, 10).Font.Size = 10
    
    ' Выравнивание
    ws.Range(ws.Cells(3, 6), ws.Cells(lastTMCRow, 6)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(3, 9), ws.Cells(lastTMCRow, 9)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(3, 10), ws.Cells(lastTMCRow, 10)).HorizontalAlignment = xlRight
    
    ' ============ ЧАСТЬ 4: ЗАГРУЗКА ДАННЫХ ИЗ ФАЙЛА "СТАЛО" ============
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Укажите путь к файлу СТАЛО"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            MsgBox "Файл не выбран. Операция отменена."
            Exit Sub
        End If
    End With
    
    ' Сохраняем имя файла "Стало"
    имя_файла_Стало = Right(filePath, Len(filePath) - InStrRev(filePath, "\"))
    
    ' Открываем файл "Стало"
    Set wbСтало = Workbooks.Open(filePath)
    
    ' Обновляем название файлов в ячейке A1 (две строки)
    ws.Range("A1").Value = имя_файла_Было & Chr(10) & имя_файла_Стало
    ws.Range("A1").WrapText = True
    
    ' ============ ОБРАБОТКА СМР ============
    
    On Error Resume Next
    Set wsСМР_Стало = wbСтало.Worksheets("СМР уник")
    On Error GoTo 0
    
    If wsСМР_Стало Is Nothing Then
        MsgBox "В файле ""Стало"" нет листа ""СМР уник""! Данные СМР не будут загружены."
    Else
        lastRow = wsСМР_Стало.Cells(wsСМР_Стало.Rows.Count, 1).End(xlUp).Row
        
        If lastRow >= 2 Then
            newRowCount = 0
            
            For r = 2 To lastRow
                код_Стало = Trim(wsСМР_Стало.Cells(r, 1).Value)
                If код_Стало <> "" Then
                    foundRow = 0
                    For dataRow = firstSMРRow To lastSMРRow
                        код_БС = Trim(ws.Cells(dataRow, 1).Value)
                        If код_БС = код_Стало Then
                            foundRow = dataRow
                            Exit For
                        End If
                    Next dataRow
                    
                    If foundRow > 0 Then
                        ws.Cells(foundRow, 7).Value = wsСМР_Стало.Cells(r, 4).Value
                        ws.Cells(foundRow, 8).Value = wsСМР_Стало.Cells(r, 5).Value
                        
                        ws.Cells(foundRow, 7).NumberFormatLocal = moneyFormat
                        ws.Cells(foundRow, 8).NumberFormatLocal = moneyFormat
                        ws.Cells(foundRow, 7).Font.Name = "Aptos Narrow"
                        ws.Cells(foundRow, 8).Font.Name = "Aptos Narrow"
                        ws.Cells(foundRow, 7).Font.Size = 8
                        ws.Cells(foundRow, 8).Font.Size = 8
                        ws.Cells(foundRow, 7).HorizontalAlignment = xlRight
                        ws.Cells(foundRow, 8).HorizontalAlignment = xlRight
                        
                        ws.Cells(foundRow, 9).Formula = "=ROUND(G" & foundRow & "*H" & foundRow & ", 2)"
                        ws.Cells(foundRow, 9).NumberFormatLocal = moneyFormat
                        ws.Cells(foundRow, 10).Formula = "=I" & foundRow & "-F" & foundRow
                        ws.Cells(foundRow, 10).NumberFormatLocal = moneyFormat
                    Else
                        insertRow = tmciRow
                        ws.Rows(insertRow).Insert Shift:=xlDown
                        
                        ws.Cells(insertRow, 1).Value = wsСМР_Стало.Cells(r, 1).Value
                        ws.Cells(insertRow, 2).Value = wsСМР_Стало.Cells(r, 2).Value
                        ws.Cells(insertRow, 3).Value = wsСМР_Стало.Cells(r, 3).Value
                        ws.Cells(insertRow, 7).Value = wsСМР_Стало.Cells(r, 4).Value
                        ws.Cells(insertRow, 8).Value = wsСМР_Стало.Cells(r, 5).Value
                        
                        With ws.Range(ws.Cells(insertRow, 1), ws.Cells(insertRow, 10))
                            .Font.Name = "Aptos Narrow"
                            .Font.Size = 8
                            .VerticalAlignment = xlCenter
                            .Interior.ColorIndex = xlNone
                        End With
                        
                        ws.Cells(insertRow, 1).HorizontalAlignment = xlCenter
                        ws.Cells(insertRow, 2).HorizontalAlignment = xlLeft
                        ws.Cells(insertRow, 2).WrapText = True
                        ws.Cells(insertRow, 3).HorizontalAlignment = xlCenter
                        ws.Cells(insertRow, 4).HorizontalAlignment = xlRight
                        ws.Cells(insertRow, 5).HorizontalAlignment = xlRight
                        ws.Cells(insertRow, 6).HorizontalAlignment = xlRight
                        ws.Cells(insertRow, 7).HorizontalAlignment = xlRight
                        ws.Cells(insertRow, 8).HorizontalAlignment = xlRight
                        ws.Cells(insertRow, 9).HorizontalAlignment = xlRight
                        ws.Cells(insertRow, 10).HorizontalAlignment = xlRight
                        
                        ws.Cells(insertRow, 7).NumberFormatLocal = moneyFormat
                        ws.Cells(insertRow, 8).NumberFormatLocal = moneyFormat
                        
                        ws.Cells(insertRow, 6).Formula = "=ROUND(D" & insertRow & "*E" & insertRow & ", 2)"
                        ws.Cells(insertRow, 6).NumberFormatLocal = moneyFormat
                        ws.Cells(insertRow, 9).Formula = "=ROUND(G" & insertRow & "*H" & insertRow & ", 2)"
                        ws.Cells(insertRow, 9).NumberFormatLocal = moneyFormat
                        ws.Cells(insertRow, 10).Formula = "=I" & insertRow & "-F" & insertRow
                        ws.Cells(insertRow, 10).NumberFormatLocal = moneyFormat
                        
                        tmciRow = tmciRow + 1
                        If firstTMCRow > 0 Then
                            firstTMCRow = firstTMCRow + 1
                            lastTMCRow = lastTMCRow + 1
                        End If
                        newRowCount = newRowCount + 1
                    End If
                End If
            Next r
            
            If newRowCount > 0 Then
                lastSMРRow = lastSMРRow + newRowCount
            End If
            
            For r = firstSMРRow To lastSMРRow
                If ws.Cells(r, 6).Formula = "" Then
                    ws.Cells(r, 6).Formula = "=ROUND(D" & r & "*E" & r & ", 2)"
                    ws.Cells(r, 6).NumberFormatLocal = moneyFormat
                    ws.Cells(r, 6).Font.Name = "Aptos Narrow"
                    ws.Cells(r, 6).Font.Size = 8
                End If
                If ws.Cells(r, 9).Formula = "" Then
                    ws.Cells(r, 9).Formula = "=ROUND(G" & r & "*H" & r & ", 2)"
                    ws.Cells(r, 9).NumberFormatLocal = moneyFormat
                    ws.Cells(r, 9).Font.Name = "Aptos Narrow"
                    ws.Cells(r, 9).Font.Size = 8
                End If
                If ws.Cells(r, 10).Formula = "" Then
                    ws.Cells(r, 10).Formula = "=I" & r & "-F" & r
                    ws.Cells(r, 10).NumberFormatLocal = moneyFormat
                    ws.Cells(r, 10).Font.Name = "Aptos Narrow"
                    ws.Cells(r, 10).Font.Size = 8
                End If
            Next r
            
            If firstSMРRow <= lastSMРRow Then
                ws.Cells(4, 6).Formula = "=SUM(F" & firstSMРRow & ":F" & lastSMРRow & ")"
                ws.Cells(4, 6).NumberFormatLocal = moneyFormat
                ws.Cells(4, 9).Formula = "=SUM(I" & firstSMРRow & ":I" & lastSMРRow & ")"
                ws.Cells(4, 9).NumberFormatLocal = moneyFormat
            End If
            
            ws.Cells(4, 10).Formula = "=I4-F4"
            ws.Cells(4, 10).NumberFormatLocal = moneyFormat
            
            If firstTMCRow <= lastTMCRow Then
                ws.Cells(tmciRow, 6).Formula = "=SUM(F" & firstTMCRow & ":F" & lastTMCRow & ")"
                ws.Cells(tmciRow, 6).NumberFormatLocal = moneyFormat
                ws.Cells(tmciRow, 9).Formula = "=SUM(I" & firstTMCRow & ":I" & lastTMCRow & ")"
                ws.Cells(tmciRow, 9).NumberFormatLocal = moneyFormat
            End If
            
            ws.Cells(tmciRow, 10).Formula = "=I" & tmciRow & "-F" & tmciRow
            ws.Cells(tmciRow, 10).NumberFormatLocal = moneyFormat
            ws.Cells(3, 6).Formula = "=F4+F" & tmciRow
            ws.Cells(3, 6).NumberFormatLocal = moneyFormat
            ws.Cells(3, 9).Formula = "=I4+I" & tmciRow
            ws.Cells(3, 9).NumberFormatLocal = moneyFormat
            ws.Cells(3, 10).Formula = "=J4+J" & tmciRow
            ws.Cells(3, 10).NumberFormatLocal = moneyFormat
        End If
    End If
    
    ' ============ ОБРАБОТКА ТМЦ ============
    
    On Error Resume Next
    Set wsТМЦ_Стало = wbСтало.Worksheets("ТМЦ уник")
    On Error GoTo 0
    
    If wsТМЦ_Стало Is Nothing Then
        MsgBox "В файле ""Стало"" нет листа ""ТМЦ уник""! Данные ТМЦ не будут загружены."
    Else
        lastRow = wsТМЦ_Стало.Cells(wsТМЦ_Стало.Rows.Count, 1).End(xlUp).Row
        
        If lastRow >= 2 Then
            newRowCount = 0
            
            For r = 2 To lastRow
                код_Стало = Trim(wsТМЦ_Стало.Cells(r, 1).Value)
                If код_Стало <> "" Then
                    foundRow = 0
                    For dataRow = firstTMCRow To lastTMCRow
                        код_БС = Trim(ws.Cells(dataRow, 1).Value)
                        If код_БС = код_Стало Then
                            foundRow = dataRow
                            Exit For
                        End If
                    Next dataRow
                    
                    If foundRow > 0 Then
                        ws.Cells(foundRow, 7).Value = wsТМЦ_Стало.Cells(r, 4).Value
                        ws.Cells(foundRow, 8).Value = wsТМЦ_Стало.Cells(r, 5).Value
                        
                        ws.Cells(foundRow, 7).NumberFormatLocal = moneyFormat
                        ws.Cells(foundRow, 8).NumberFormatLocal = moneyFormat
                        ws.Cells(foundRow, 7).Font.Name = "Aptos Narrow"
                        ws.Cells(foundRow, 8).Font.Name = "Aptos Narrow"
                        ws.Cells(foundRow, 7).Font.Size = 8
                        ws.Cells(foundRow, 8).Font.Size = 8
                        ws.Cells(foundRow, 7).HorizontalAlignment = xlRight
                        ws.Cells(foundRow, 8).HorizontalAlignment = xlRight
                        
                        ws.Cells(foundRow, 9).Formula = "=ROUND(G" & foundRow & "*H" & foundRow & ", 2)"
                        ws.Cells(foundRow, 9).NumberFormatLocal = moneyFormat
                        ws.Cells(foundRow, 10).Formula = "=I" & foundRow & "-F" & foundRow
                        ws.Cells(foundRow, 10).NumberFormatLocal = moneyFormat
                    Else
                        insertRow = tmciRow + 1
                        ws.Rows(insertRow).Insert Shift:=xlDown
                        
                        ws.Cells(insertRow, 1).Value = wsТМЦ_Стало.Cells(r, 1).Value
                        ws.Cells(insertRow, 2).Value = wsТМЦ_Стало.Cells(r, 2).Value
                        ws.Cells(insertRow, 3).Value = wsТМЦ_Стало.Cells(r, 3).Value
                        ws.Cells(insertRow, 7).Value = wsТМЦ_Стало.Cells(r, 4).Value
                        ws.Cells(insertRow, 8).Value = wsТМЦ_Стало.Cells(r, 5).Value
                        
                        With ws.Range(ws.Cells(insertRow, 1), ws.Cells(insertRow, 10))
                            .Font.Name = "Aptos Narrow"
                            .Font.Size = 8
                            .VerticalAlignment = xlCenter
                            .Interior.ColorIndex = xlNone
                        End With
                        
                        ws.Cells(insertRow, 1).HorizontalAlignment = xlCenter
                        ws.Cells(insertRow, 2).HorizontalAlignment = xlLeft
                        ws.Cells(insertRow, 2).WrapText = True
                        ws.Cells(insertRow, 3).HorizontalAlignment = xlCenter
                        ws.Cells(insertRow, 4).HorizontalAlignment = xlRight
                        ws.Cells(insertRow, 5).HorizontalAlignment = xlRight
                        ws.Cells(insertRow, 6).HorizontalAlignment = xlRight
                        ws.Cells(insertRow, 7).HorizontalAlignment = xlRight
                        ws.Cells(insertRow, 8).HorizontalAlignment = xlRight
                        ws.Cells(insertRow, 9).HorizontalAlignment = xlRight
                        ws.Cells(insertRow, 10).HorizontalAlignment = xlRight
                        
                        ws.Cells(insertRow, 7).NumberFormatLocal = moneyFormat
                        ws.Cells(insertRow, 8).NumberFormatLocal = moneyFormat
                        
                        ws.Cells(insertRow, 6).Formula = "=ROUND(D" & insertRow & "*E" & insertRow & ", 2)"
                        ws.Cells(insertRow, 6).NumberFormatLocal = moneyFormat
                        ws.Cells(insertRow, 9).Formula = "=ROUND(G" & insertRow & "*H" & insertRow & ", 2)"
                        ws.Cells(insertRow, 9).NumberFormatLocal = moneyFormat
                        ws.Cells(insertRow, 10).Formula = "=I" & insertRow & "-F" & insertRow
                        ws.Cells(insertRow, 10).NumberFormatLocal = moneyFormat
                        
                        lastTMCRow = lastTMCRow + 1
                        newRowCount = newRowCount + 1
                    End If
                End If
            Next r
            
            For r = firstTMCRow To lastTMCRow
                If ws.Cells(r, 6).Formula = "" Then
                    ws.Cells(r, 6).Formula = "=ROUND(D" & r & "*E" & r & ", 2)"
                    ws.Cells(r, 6).NumberFormatLocal = moneyFormat
                    ws.Cells(r, 6).Font.Name = "Aptos Narrow"
                    ws.Cells(r, 6).Font.Size = 8
                End If
                If ws.Cells(r, 9).Formula = "" Then
                    ws.Cells(r, 9).Formula = "=ROUND(G" & r & "*H" & r & ", 2)"
                    ws.Cells(r, 9).NumberFormatLocal = moneyFormat
                    ws.Cells(r, 9).Font.Name = "Aptos Narrow"
                    ws.Cells(r, 9).Font.Size = 8
                End If
                If ws.Cells(r, 10).Formula = "" Then
                    ws.Cells(r, 10).Formula = "=I" & r & "-F" & r
                    ws.Cells(r, 10).NumberFormatLocal = moneyFormat
                    ws.Cells(r, 10).Font.Name = "Aptos Narrow"
                    ws.Cells(r, 10).Font.Size = 8
                End If
            Next r
            
            If firstTMCRow <= lastTMCRow Then
                ws.Cells(tmciRow, 6).Formula = "=SUM(F" & firstTMCRow & ":F" & lastTMCRow & ")"
                ws.Cells(tmciRow, 6).NumberFormatLocal = moneyFormat
                ws.Cells(tmciRow, 9).Formula = "=SUM(I" & firstTMCRow & ":I" & lastTMCRow & ")"
                ws.Cells(tmciRow, 9).NumberFormatLocal = moneyFormat
            End If
            
            ws.Cells(tmciRow, 10).Formula = "=I" & tmciRow & "-F" & tmciRow
            ws.Cells(tmciRow, 10).NumberFormatLocal = moneyFormat
            ws.Cells(3, 6).Formula = "=F4+F" & tmciRow
            ws.Cells(3, 6).NumberFormatLocal = moneyFormat
            ws.Cells(3, 9).Formula = "=I4+I" & tmciRow
            ws.Cells(3, 9).NumberFormatLocal = moneyFormat
            ws.Cells(3, 10).Formula = "=J4+J" & tmciRow
            ws.Cells(3, 10).NumberFormatLocal = moneyFormat
        End If
    End If
    
    wbСтало.Close False
    Application.ScreenUpdating = True
    
    ' ============ УСТАНОВКА ГРАНИЦ ============
    
    ' Определяем последнюю непустую строку
    Dim lastDataRow As Long
    lastDataRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Устанавливаем границы для всей таблицы (колонки 1-10, строки 1 - lastDataRow)
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastDataRow, 10))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.ColorIndex = 1
    End With
    
    ' ============ ВЫВОД ИТОГОВОГО СООБЩЕНИЯ ============
    
    Dim smr6 As Double
    Dim smr9 As Double
    Dim tmci6 As Double
    Dim tmci9 As Double
    Dim total6 As Double
    Dim total9 As Double
    
    On Error Resume Next
    smr6 = ws.Cells(4, 6).Value
    smr9 = ws.Cells(4, 9).Value
    tmci6 = ws.Cells(tmciRow, 6).Value
    tmci9 = ws.Cells(tmciRow, 9).Value
    total6 = ws.Cells(3, 6).Value
    total9 = ws.Cells(3, 9).Value
    On Error GoTo 0
    
    ' Функция для форматирования числа через временную ячейку с автоматической шириной
    Dim tempCell As Range
    
    ' Создаем временную ячейку в колонке K (чтобы не мешать основной таблице)
    ' и делаем ее достаточно широкой
    ws.Columns(11).ColumnWidth = 30
    Set tempCell = ws.Cells(lastDataRow + 2, 11)
    
    ' Форматируем каждое число через временную ячейку
    Dim smr6_str As String
    Dim smr9_str As String
    Dim tmci6_str As String
    Dim tmci9_str As String
    Dim total6_str As String
    Dim total9_str As String
    
    tempCell.Value = smr6
    tempCell.NumberFormatLocal = moneyFormat
    smr6_str = tempCell.Text
    
    tempCell.Value = smr9
    tempCell.NumberFormatLocal = moneyFormat
    smr9_str = tempCell.Text
    
    tempCell.Value = tmci6
    tempCell.NumberFormatLocal = moneyFormat
    tmci6_str = tempCell.Text
    
    tempCell.Value = tmci9
    tempCell.NumberFormatLocal = moneyFormat
    tmci9_str = tempCell.Text
    
    tempCell.Value = total6
    tempCell.NumberFormatLocal = moneyFormat
    total6_str = tempCell.Text
    
    tempCell.Value = total9
    tempCell.NumberFormatLocal = moneyFormat
    total9_str = tempCell.Text
    
    ' Очищаем временную ячейку и восстанавливаем ширину колонки K
    tempCell.Clear
    ws.Columns(11).ColumnWidth = 0 ' Скрываем колонку K
    
    Dim msg As String
    msg = "Обработка завершена." & vbCrLf & vbCrLf
    msg = msg & "Проверочные суммы:" & vbCrLf & vbCrLf
    msg = msg & "БЫЛО:" & vbCrLf
    msg = msg & "  СМР: " & smr6_str & ", руб." & vbCrLf
    msg = msg & "  ТМЦ: " & tmci6_str & ", руб." & vbCrLf
    msg = msg & "  Всего: " & total6_str & ", руб." & vbCrLf & vbCrLf
    msg = msg & "СТАЛО:" & vbCrLf
    msg = msg & "  СМР: " & smr9_str & ", руб." & vbCrLf
    msg = msg & "  ТМЦ: " & tmci9_str & ", руб." & vbCrLf
    msg = msg & "  Всего: " & total9_str & ", руб."
    
    MsgBox msg, vbInformation, "Результат"
    
End Sub

