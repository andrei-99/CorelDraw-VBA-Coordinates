Sub CollectAndExportCoordinates()
    Dim doc As Document
    Set doc = Application.ActiveDocument
    
    ' Устанавливаем единицы измерения в миллиметры
    doc.Unit = cdrMillimeter
    
    Dim layer As Layer
    Set layer = doc.ActiveLayer
    
    ' Собираем координаты центров объектов
    Dim coordCollection As Collection
    Set coordCollection = New Collection
    
    Dim sh As Shape
    For Each sh In layer.Shapes
        If sh.Type <> cdrGuidesShape Then ' Исключаем направляющие
            Dim centerX As Double
            Dim centerY As Double
            
            ' Получаем координаты центра объекта
            centerX = sh.CenterX
            centerY = sh.CenterY
            
            ' Округляем до целых миллиметров по обычным правилам
            centerX = Round(centerX)
            centerY = Round(centerY)
            
            ' Сохраняем координаты и ссылку на объект
            Dim coordData(2) As Variant
            coordData(0) = centerX
            coordData(1) = centerY
            coordData(2) = sh ' Сохраняем ссылку на объект для возможного дальнейшего использования
            
            coordCollection.Add coordData
        End If
    Next sh
    
    If coordCollection.Count = 0 Then
        MsgBox "На активном слое не найдено объектов для обработки."
        Exit Sub
    End If
    
    ' Сортируем массив "змейкой" (слева направо, снизу вверх)
    Dim sortedCoords() As Variant
    sortedCoords = SortCoordinatesSnakeStyle(coordCollection)
    
    ' Нормируем координаты относительно левого нижнего объекта
    Dim normalizedCoords() As Variant
    normalizedCoords = NormalizeCoordinates(sortedCoords)
    
    ' Сохраняем в файл
    SaveCoordinatesToFile doc, normalizedCoords
    
    MsgBox "Координаты успешно экспортированы! Обработано объектов: " & coordCollection.Count
End Sub

Function SortCoordinatesSnakeStyle(coordCollection As Collection) As Variant()
    ' Преобразуем коллекцию в массив для сортировки
    Dim coords() As Variant
    ReDim coords(coordCollection.Count - 1, 2)
    
    Dim i As Long
    For i = 1 To coordCollection.Count
        Dim coordData As Variant
        coordData = coordCollection(i)
        coords(i - 1, 0) = coordData(0) ' X
        coords(i - 1, 1) = coordData(1) ' Y
        coords(i - 1, 2) = coordData(2) ' Object reference
    Next i
    
    ' Сортируем пузырьковой сортировкой по Y (снизу вверх), затем по X (слева направо)
    Dim j As Long, k As Long
    Dim tempX As Double, tempY As Double, tempObj As Object
    
    For j = 0 To UBound(coords, 1) - 1
        For k = j + 1 To UBound(coords, 1)
            ' Сначала сравниваем Y (ряды)
            If coords(k, 1) < coords(j, 1) Then
                ' Меняем местами
                tempX = coords(j, 0)
                tempY = coords(j, 1)
                tempObj = coords(j, 2)
                
                coords(j, 0) = coords(k, 0)
                coords(j, 1) = coords(k, 1)
                coords(j, 2) = coords(k, 2)
                
                coords(k, 0) = tempX
                coords(k, 1) = tempY
                coords(k, 2) = tempObj
            ElseIf coords(k, 1) = coords(j, 1) Then
                ' Если Y одинаковый, сравниваем X
                If coords(k, 0) < coords(j, 0) Then
                    ' Меняем местами
                    tempX = coords(j, 0)
                    tempY = coords(j, 1)
                    tempObj = coords(j, 2)
                    
                    coords(j, 0) = coords(k, 0)
                    coords(j, 1) = coords(k, 1)
                    coords(j, 2) = coords(k, 2)
                    
                    coords(k, 0) = tempX
                    coords(k, 1) = tempY
                    coords(k, 2) = tempObj
                End If
            End If
        Next k
    Next j
    
    SortCoordinatesSnakeStyle = coords
End Function

Function NormalizeCoordinates(coords() As Variant) As Variant()
    If UBound(coords, 1) < 0 Then
        NormalizeCoordinates = coords
        Exit Function
    End If
    
    ' Находим левый нижний объект (минимальный X среди объектов с минимальным Y)
    Dim minY As Double
    Dim minX As Double
    Dim refIndex As Long
    
    minY = coords(0, 1)
    minX = coords(0, 0)
    refIndex = 0
    
    Dim i As Long
    For i = 1 To UBound(coords, 1)
        If coords(i, 1) < minY Then
            minY = coords(i, 1)
            minX = coords(i, 0)
            refIndex = i
        ElseIf coords(i, 1) = minY And coords(i, 0) < minX Then
            minX = coords(i, 0)
            refIndex = i
        End If
    Next i
    
    ' Создаем массив с нормированными координатами
    Dim normalized() As Variant
    ReDim normalized(UBound(coords, 1), 2)
    
    For i = 0 To UBound(coords, 1)
        normalized(i, 0) = coords(i, 0) - minX ' Нормированный X
        normalized(i, 1) = coords(i, 1) - minY ' Нормированный Y
        normalized(i, 2) = coords(i, 2) ' Сохраняем ссылку на объект
    Next i
    
    NormalizeCoordinates = normalized
End Function

Sub SaveCoordinatesToFile(doc As Document, coords() As Variant)
    Dim filePath As String
    Dim fileName As String
    Dim txtFilePath As String
    
    ' Получаем путь и имя текущего файла
    filePath = doc.FilePath
    fileName = doc.FileName
    
    ' Если файл не сохранен, используем временный путь
    If filePath = "" Then
        filePath = Application.TemporaryFolderPath
        fileName = "Untitled"
    Else
        ' Убираем расширение из имени файла
        If InStr(fileName, ".") > 0 Then
            fileName = Left(fileName, InStr(fileName, ".") - 1)
        End If
    End If
    
    ' Формируем путь для текстового файла
    txtFilePath = filePath & "\" & fileName & "_coordinates.txt"
    
    ' Создаем и записываем в файл
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim txtFile As Object
    Set txtFile = fso.CreateTextFile(txtFilePath, True)
    
    ' Записываем заголовок
    txtFile.WriteLine "Координаты центров объектов (в мм)"
    txtFile.WriteLine "Отсортированы змейкой, нормированы относительно левого нижнего объекта"
    txtFile.WriteLine "Формат: X, Y"
    txtFile.WriteLine "=================================="
    
    ' Записываем координаты
    Dim i As Long
    For i = 0 To UBound(coords, 1)
        txtFile.WriteLine CStr(coords(i, 0)) & ", " & CStr(coords(i, 1))
    Next i
    
    ' Записываем итоговую информацию
    txtFile.WriteLine "=================================="
    txtFile.WriteLine "Всего объектов: " & CStr(UBound(coords, 1) + 1)
    
    txtFile.Close
    
    ' Показываем сообщение о сохранении
    MsgBox "Координаты сохранены в файл: " & vbCrLf & txtFilePath
End Sub
