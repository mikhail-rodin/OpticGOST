Attribute VB_Name = "zmxImport"
Public Type surfData 'поверхность
    id As String    'номер либо IMA/STO/OBJ
    type As String
    r As Double
    n As Double
    nShort As Double 'n для короткой длины волны
    nLong As Double 'n для большой длины волны
    D As Double
    v As Double     'коэф. Аббе
    glass As String
    diam As Double      'световой диаметр
    sag As Double       'стрелка прогиба по световому диаметру
End Type
Dim surface() As surfData
Dim stopPos As Integer 'положение диафрагмы

Public Type rayData
'индекс начинается с 1
'по адресу 0 - источник (имя файла)
    axialRayH As String 'высота осевого луча
    lowerRayH As String 'высота нижнего апертурного луча
    upperRayH As String 'высота верхнего апертурного луча
    chiefRayH As String 'высота главного луча
End Type
Public rays() As rayData
Public fieldCos As Double

Dim lineArray() As String 'сюда будем заносить строки
Public surfCounter, elementCounter As Integer 'число поверхностей, подсчитывается при импорте файла
Dim shortwave As Double 'короткая волна (для вычисления числа Аббе)
Dim longwave As Double 'длинная волна (для вычисления числа Аббе)
Public waveSelection, shortWaveSel, longWaveSel As Integer 'номер выбранной длины волны
Const rndDataStartString As String = "SURFACE DATA SUMMARY:"
Const firstSurfaceString As String = "OBJ"
Const lastSurfaceString As String = "IMA"
Const glassIndexStartString As String = "INDEX OF REFRACTION DATA:"
Const thermalStartString As String = "THERMAL COEFFICIENT OF EXPANSION DATA:"
Function FirstNotDigit(str As String) As Boolean
    Dim code As Integer
    code = Asc(Mid(str, 1, 1)) 'определим по ASCII коду первого символа
    Select Case code
        Case 48 To 57: FirstNotDigit = False
        Case Else: FirstNotDigit = True
    End Select
End Function
Function LZOStranslate(glass As String) As String
    Dim temp As String
    If InStr(1, glass, "LZ_") Then
        temp = Replace(glass, "LZ_", "")
        temp = Replace(temp, "F", "Ф")
        temp = Replace(temp, "B", "Б")
        temp = Replace(temp, "L", "Л")
        LZOStranslate = temp
    Else
        LZOStranslate = glass
    End If
End Function
Function TableParse(row As String) As String()
'на входе строчка таблицы из файла Zemax Prescription Data
'на выходе динамический массив из слов в строчке
'все пробелы удаляются
    Dim result() As String
    Dim temp(20) As String
    Dim tempLen As Integer
  '            Do
'              tempStr = lineArray(rowCounter)
'              lineArray(i) = Replace(lineArray(i), "  ", " ") 'remove multiple white spaces
'            Loop Until tempStr = lineArray(rowCounter)
  
    'ReDim result(1)
    result = Split(Trim(row), " ") 'trim to remove starting/trailing space
    'сейчас у нас пробелы записаны в массив как отдельные элементы
    'что ж, почистим массив от этого говна
    tempLen = 0
    For j = 0 To UBound(result)
        If result(j) <> "" Then
            tempLen = tempLen + 1
            temp(tempLen - 1) = result(j)
        End If
    Next j
    For j = 0 To tempLen - 1
        result(j) = temp(j)
    Next j
    ReDim Preserve result(tempLen)
'    ReDim Preserve splitLine(6)
    TableParse = result
End Function
Function FindString(text() As String, str As String) As Integer
'возвращает номер строки или 0 в случае ошибки
    Dim pos As Integer 'номер строки
    pos = 0
    Do Until pos = UBound(text) 'ищем "SURFACE DATA SUMMARY:"
        If Not InStr(text(pos), str) = 0 Then
            Exit Do
        End If
        pos = pos + 1
    Loop
    
    If pos = UBound(text) Then 'если ничего не нашли
        FindString = 0
    Else
        FindString = pos
    End If
End Function
Function FindStringBetween(text() As String, str As String, _
        ByVal searchStart As Integer, ByVal searchEnd As Integer) As Integer
    'возвращает номер строки или 0 в случае ошибки
    Dim pos As Integer 'номер строки
  
    For pos = searchStart To searchEnd
        If Not InStr(text(pos), str) = 0 Then
            Exit For
        End If
    Next pos
    
    If pos = UBound(text) Then 'если ничего не нашли
        FindStringBetween = 0
    Else
        FindStringBetween = pos
    End If
    
End Function

Sub launchUI()
rndForm.Show


'rndForm.ESKDstart.Text = "A1"
End Sub

Sub zmxPrescriptionImport()
    Dim filename, line As String
    Dim splitline() As String 'в ячейки массива заносятся строки по слову
    Dim tempArray(100) As String
    Dim rowCounter, lineCounter, rndDataPos, i, j, waveNumber, tempArrLen As Integer 'номер строки в таблице
    Dim glassIndexDataPos As Integer
    Dim tempStr As String
    Dim txtfile As Integer
    Dim wavelenght As Double

    txtfile = FreeFile() 'file handle
    
    filename = rndForm.filePath.text
    On Error Resume Next:
    Open filename For Input As txtfile
    If Err <> 0 Then
        MsgBox "Файл не найден в директории " & filename, vbCritical, "Ошибка"
        Exit Sub
    End If
    
    With rndForm.statusLabel
            .Caption = "Файл открыт."
            .ForeColor = RGB(10, 120, 10)
    End With
    
    rndForm.rndFillTableBtn.Enabled = True
    
    lineCounter = 0

    ReDim lineArray(1)
    While Not EOF(txtfile) 'чтение файла в массив, на всякий случай не даём программе выйти за пределы файла
        Line Input #txtfile, tempStr
        lineArray(lineCounter) = tempStr
        lineCounter = lineCounter + 1 'посчитаем, сколько у нас строчек
        ReDim Preserve lineArray(lineCounter)
    Wend
    Close txtfile
 
    ReDim Preserve surface(surfCounter) As surfData
    
   ' rndForm.statusLabel.Caption = "Файл импортирован. В файле " & lineCounter - 1 & " строк(и)"

    Application.ScreenUpdating = False

    rndDataPos = 0
    Do Until rndDataPos = lineCounter 'ищем "SURFACE DATA SUMMARY:"

        If Not InStr(lineArray(rndDataPos), rndDataStartString) = 0 Then
            Exit Do
        End If

        rndDataPos = rndDataPos + 1
    Loop
    
    rndForm.textBox.text = ""
    For i = 0 To lineCounter 'выводим текст в окошко
    rndForm.textBox.text = rndForm.textBox.text & i & " >> " & lineArray(i) & vbCrLf
    Next i

    'MsgBox "Строка " & rndDataPos, vbInformation, "Импорт"

'считываем SURFACE DATA SUMMARY
    j = 0
    surfCounter = 0
    i = rndDataPos + 2 'пропускаем название, пустую строку и шапку таблицы
    ReDim surface(1)
    Do Until i = lineCounter
        If Not lineArray(rowCounter) = "" Then 'пропускаем цикл если строка пустая
            splitline = TableParse(lineArray(i))
                'теперь в массиве splitLine так:
                '0 - номер п/п
                '1 - тип поверхности
                '2 - радиус
                '3 - толщина
                '4 - стекло или - !!! - св. диаметр если нет стекла - это неправильно, сейчас исправим
                '5 - световой диаметр, если есть стекло
            If Not InStr(splitline(0), "STO") = 0 Then
            'если наткунились на диафрагму
            'игнорируем её
            'и прибавляем расстояние после диафрагмы к расст. после неё
                stopPos = surfCounter
                surface(surfCounter - 1).D = surface(surfCounter - 1).D + Val(splitline(3))
            'не инкрементируем surfcounter
            'поэтому surface(surfCounter) в след. цикле перезапишется
            Else
                If FirstNotDigit(splitline(4)) = False Then 'если первый символ в поле "стеко" это цифра
                    'значит, ни хера это не стекло
                    splitline(5) = splitline(4) 'значит, там у нас св. диаметр, и мы копируем его в (5)
                    splitline(4) = ""
                End If
    
                If Not InStr(splitline(0), lastSurfaceString) = 0 Then
                    Exit Do 'Если "IMA", выходим
                End If
                
                ReDim Preserve surface(surfCounter) As surfData
                
                surface(surfCounter).id = splitline(0)
                
                surface(surfCounter).type = splitline(1)
                
                If splitline(2) = "Infinity" Then
                    surface(surfCounter).r = 0
                Else
                    surface(surfCounter).r = Val(splitline(2))
                End If
                
                If InStr(splitline(3), "Infinity") = 0 Then
                'если не бесконечность
                    surface(surfCounter).D = Val(splitline(3))
                Else
                'записываем ноль
                    surface(surfCounter).D = 0
                End If
                
                surface(surfCounter).glass = splitline(4)
                
                surface(surfCounter).diam = Val(splitline(5))
                
                surfCounter = surfCounter + 1 'посчитали ещё одну поверхность
            End If
        End If
        i = i + 1 'идём на следующую строку
        
    Loop
    
    
    glassIndexDataPos = FindString(lineArray, glassIndexStartString)

    j = 0 'заходим в первую поверхность
    For i = glassIndexDataPos + 7 To glassIndexDataPos + 7 + surfCounter
        If Not lineArray(i) = "" Then
            splitline = TableParse(lineArray(i))
            If InStr(splitline(0), "Surf") <> 0 Then 'если мы в шапке
            'заполним список длин волн
                waveNumber = 0
                shortwave = 10 'готовимся искать мин/макс длину волны
                longwave = 0
                rndForm.wavelengthList.Clear
                Do Until 4 + waveNumber = UBound(splitline)
                    If Not splitline(4 + waveNumber) = "" Then
                        wavelength = Val(splitline(4 + waveNumber)) 'String to Double
                        If wavelength < shortwave Then 'ищем самую короткую волну
                            shortwave = wavelength
                            shortWaveSel = waveNumber
                        End If
                        If wavelength > longwave Then
                            longwave = wavelength
                            longWaveSel = waveNumber
                        End If
                        rndForm.wavelengthList.AddItem (waveNumber & " - " & Left(splitline(4 + waveNumber), 5) & " мкм")
                    End If
                    waveNumber = waveNumber + 1
                Loop
                Exit For
            End If

'            If FirstNotDigit(splitline(1)) = True Then 'если стекло есть
'                surface(j).n = splitline(4 + waveSelection) 'n на 5-м месте (индекс массива 4)
'            Else 'а если стекла нет
'                surface(j).n = splitline(3 + waveSelection) 'тогда на 4-м
'            End If
            j = j + 1
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    
    rndForm.shortwaveLabel.Caption = "Короткая волна: " & shortWaveSel & " - " & shortwave & " мкм"
    rndForm.longwaveLabel.Caption = "Длинная волна: " & longWaveSel & " - " & longwave & " мкм"
    
    With rndForm.statusLabel
            .Caption = "Поверхностей: " & surfCounter & _
                ", включая плоскости предмета (IMA) и изображения (IMA)" & vbCrLf & _
                "Длин волн: " & waveNumber
            .ForeColor = RGB(10, 120, 10)
    End With
    
    
    With rndForm
        .rndFillTableBtn.Enabled = True
        .newLensSheetchk.Enabled = True
        .generateESKDchk.Enabled = True
        .generateZemaxTableChk.Enabled = True
        .lensTableChk.Enabled = True
        .createSheetChk.Enabled = True
        .ZemaxStart.Enabled = True
        .ESKDstart.Enabled = True
        .lensStart.Enabled = True
    End With
    'указываем начало второй таблицы
    rndForm.ZemaxStart.text = Range(rndForm.ESKDstart.text).Offset(surfCounter * 2 + 3, 0).Address
    rndForm.lensStart.text = Range(rndForm.ESKDstart.text).Offset(surfCounter * 3 + 3, 0).Address
    


End Sub

Public Sub rndFillTable()
    Dim i As Integer
    Dim j As Integer
    Dim ZemaxStartCell As String
    Dim ESKDstartCell As String
    Dim sheetID As String
    Dim LZOS As Boolean
    Dim rndSheet As Boolean
    Dim wsheet As Object
    Dim D As Double
    
    
    LZOS = rndForm.LZOSchk.value
    rndSheet = rndForm.createSheetChk.value
    
    If rndSheet = True Then
        Set wsheet = Application.Worksheets.Add
        wsheet.name = rndForm.sheetName.text
    Else
        Set wsheet = Application.ActiveSheet
        sheetID = wsheet.name
        rndForm.sheetName.text = sheetID
    End If
    
    ZemaxStartCell = rndForm.ZemaxStart.text
    ESKDstartCell = rndForm.ESKDstart.text
    
    Application.ScreenUpdating = False
    
    If rndForm.generateZemaxTableChk.value = True Then
        For i = 0 To surfCounter - 1 'заносим поля из структуры в ячейки
            With wsheet
                .Range(ZemaxStartCell).Offset(i, 0).value = surface(i).r
                .Range(ZemaxStartCell).Offset(i, 1).value = surface(i).D
                .Range(ZemaxStartCell).Offset(i, 2).value = surface(i).n
            End With
            
            If surface(i).v = 0 Then 'чтобы не было нулей
                wsheet.Range(ZemaxStartCell).Offset(i, 3).value = ""
            Else
                wsheet.Range(ZemaxStartCell).Offset(i, 3).value = surface(i).v
            End If
                    
            With wsheet
                .Range(ZemaxStartCell).Offset(i, 4).value = surface(i).glass
                .Range(ZemaxStartCell).Offset(i, 5).value = surface(i).diam
                .Range(ZemaxStartCell).Offset(i, 6).value = surface(i).sag
            End With
        Next i
    End If
    
    If rndForm.generateESKDchk.value = True Then
        
        With wsheet.Range(ESKDstartCell)
            .Offset(0, 2) = "ne"
            .Offset(0, 2).Characters(2, 1).Font.Subscript = True
            
            .Offset(0, 3) = "ve" 'ChrW(55349) & "e"
            .Offset(0, 3).Characters(2, 1).Font.Subscript = True
            '.Offset(0, 3).Characters(1, 1).Font.Italic = True
            
            .Offset(0, 4) = "Марка стекла"
            .Offset(0, 5) = ChrW(216) & " св."
            .Offset(0, 6) = "стрелка по " & ChrW(216) & "  св."
        End With
        
        j = 1
        For i = 1 To (surfCounter - 2) * 2 Step 2
            With wsheet.Range(ESKDstartCell)
                If i = 1 And surface(j).D = 0 Then
                'не указываем расст. до беск. удаленного предмета
                'при импорте, если бесконечность, записывается 0
                   .Offset(i, 1).value = ""
                Else
                    D = surface(j).D
                    .Offset(i, 1).value = "d" & j - 1 & " = " & Round(D, 2)
                    .Offset(i, 1).Characters(2, 2).Font.Subscript = True
                End If

                .Offset(i, 2).value = "n" & j - 1 & " = " & Round(surface(j).n, 4)
                .Offset(i, 2).Characters(2, 2).Font.Subscript = True

                If surface(j).v = 0 Then 'чтобы не было нулей
                    .Offset(i, 3).value = ""
                Else
                    .Offset(i, 3).value = Round(surface(j).v, 2)
                    .Offset(i, 3).NumberFormat = "0.00"
                End If

                If LZOS = True Then
                    .Offset(i, 4).value = LZOStranslate(surface(j).glass)
                Else
                    .Offset(i, 4).value = surface(j).glass
                End If
            End With
                j = j + 1
        Next i

        j = 1
        For i = 0 To (surfCounter - 2) * 2 + 1 Step 2
            With wsheet.Range(ESKDstartCell)
            If j >= 2 Then
                 .Offset(i, 5).value = Round(surface(j).diam, 2)
                 .Offset(i, 5).NumberFormat = "0.00"
                 
                 .Offset(i, 6).value = Round(surface(j).sag, 2)
                 .Offset(i, 6).NumberFormat = "0.00"
            End If
            
            If Not j = 1 Then 'радиус плоскости предмета не указываем
                .Offset(i, 0).value = "r" & j - 1 & " = " & Round(surface(j).r, 2)
                .Offset(i, 0).Characters(2, 2).Font.Subscript = True
            End If
            End With
                j = j + 1
        Next i
        
        With wsheet.Range(ESKDstartCell).Resize(surfCounter * 2 + 5, 7) 'для всей таблицы
            .Columns.AutoFit
            .HorizontalAlignment = XlHAlign.xlHAlignCenter
        End With
    End If

    Application.ScreenUpdating = True
    
End Sub

Public Sub glassIndexImport()
    Static j, i, glassIndexDataPos, waveNumber As Integer
    Dim splitstr() As String
    
    glassIndexDataPos = FindString(lineArray, glassIndexStartString)
    
    j = 0 'заходим в первую поверхность
    For i = glassIndexDataPos + 7 To glassIndexDataPos + 7 + surfCounter
        If Not lineArray(i) = "" Then 'если строка файла не пустая
            splitstr = TableParse(lineArray(i))
            If Not i = glassIndexDataPos + 7 + stopPos Then
            'если в строке не диафрагма
                If Not InStr(splitstr(0), "Surf") <> 0 Then 'если мы НЕ в шапке
                    If FirstNotDigit(splitstr(1)) = True Then 'если стекло есть
                        surface(j).n = Val(splitstr(4 + waveSelection)) 'n на 5-м месте (индекс массива 4)
                        surface(j).nShort = Val(splitstr(4 + shortWaveSel))
                        surface(j).nLong = Val(splitstr(4 + longWaveSel))
                    Else 'а если стекла нет
                        surface(j).n = Val(splitstr(3 + waveSelection)) 'тогда на 4-м
                        surface(j).nShort = Val(splitstr(3 + shortWaveSel))
                        surface(j).nLong = Val(splitstr(3 + longWaveSel))
                    End If
                End If
                j = j + 1
            End If
        End If
    Next i
End Sub

Public Sub CalculateAbbe()
Static i As Integer
    For i = 0 To UBound(surface)
        If surface(i).n > 1 Then 'если мы имеем дело со стеклом
            surface(i).v = (surface(i).n - 1) / (surface(i).nShort - surface(i).nLong)
        End If
    Next i
End Sub

Public Sub CalculateSag()
Static i As Integer
Static D, r As Double

    For i = 0 To UBound(surface)
        If surface(i).r = 0 Then
            surface(i).sag = 0
        Else
            D = surface(i).diam
            r = surface(i).r
            If r > 0 Then
                surface(i).sag = r - Sqr(r ^ 2 - (D / 2) ^ 2)
            Else
                surface(i).sag = r + Sqr(r ^ 2 - (D / 2) ^ 2)
            End If
        End If
    Next i
    
End Sub
Public Sub CleanUp()
    Erase lineArray
End Sub
Public Sub lensFillTable()

Dim lensStartCell As String
Dim lensSheet As Object
Dim i As Integer

    Application.ScreenUpdating = False
    
    lensStartCell = rndForm.lensStart.text
    
    If rndForm.newLensSheetchk.value = True Then
        Set lensSheet = Application.Worksheets.Add
        lensSheet.name = rndForm.lensSheetNameBox.text
    Else
        Set lensSheet = Application.ActiveSheet
        rndForm.lensSheetNameBox.text = lensSheet.name
    End If
    
    With lensSheet.Range(lensStartCell)
        .Offset(0, 0) = "№ поз. дет."
        .Offset(0, 1) = ChrW(216) & " св."
        .Offset(0, 2) = "стрелка по " & ChrW(216) & "  св."
        .Offset(0, 3) = ChrW(216) & " св."
        .Offset(0, 4) = "стрелка по " & ChrW(216) & "  св."
        .Offset(0, 5) = "толщина по оси"
    End With
    
    elementCounter = 0
    For i = 0 To surfCounter - 1
        If Not (surface(i).glass = "" Or surface(i).glass = "Glass") Then
        'если мы наткнулись на линзу И это не заголовок
            elementCounter = elementCounter + 1 'посчитаем её
            With lensSheet.Range(lensStartCell)
                .Offset(elementCounter, 0).value = elementCounter 'номер п/п
                .Offset(elementCounter, 1).value = Round(surface(i).diam, 2)
                .Offset(elementCounter, 2).value = Round(surface(i).sag, 2)
                .Offset(elementCounter, 3).value = Round(surface(i + 1).diam, 2)
                .Offset(elementCounter, 4).value = Round(surface(i + 1).sag, 2)
                .Offset(elementCounter, 5).value = Round(surface(i).D, 2)
            End With
        End If
    Next i
    
     With lensSheet.Range(lensStartCell)
        .Offset(1, 1).Resize(1 + elementCounter, 5).NumberFormat = "0.00"
        .Offset(1, 0).Resize(1 + elementCounter, 1).NumberFormat = "0"
        .Resize(1 + elementCounter, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
        .Columns.AutoFit
        .Offset(1, 1).Columns.AutoFit
        .Offset(0, 2).Columns.AutoFit
        .Offset(1, 3).Resize(1, 6).Columns.AutoFit
        .Offset(0, 4).Columns.AutoFit
        .Offset(0, 5).Columns.AutoFit
    End With
    
    Application.ScreenUpdating = True

End Sub
Sub LaunchRaytraceUI()
    fieldCos = 0
    rayForm.Show
End Sub
Function zmxRaytraceImport(filename As String) As Integer
    'возращает число поверхностей
    'возращает 0 при ошибке
    
    Dim surfcount As Integer
    Dim lineCount As Integer
    Dim fileID As Integer
    Dim i As Integer
    Dim position As Integer
    Dim tempPos As Integer
    Dim endPosition As Integer
    Dim Hy, Py As Double
    Dim field As Double 'определяется из Y-cosine для первой поверхности
    
    Static rayType As Integer
    '1 - главный
    '2 - верхний
    '3 - нижний
    '4 - апертурный
    
    Dim buffer() As String 'сюда копируем файл
    Dim parsedString() As String 'здесь храним текущую строку пословно
    
    'проверим, загружен ли уже этот файл
    With rayForm.fileList
        For i = 0 To .ListCount - 1
            If .List(i, 3) = filename Then
            'если уже есть такой файл
                If MsgBox("Файл " & Right(filename, Len(filename) - InStrRev(filename, "\")) _
                    & " уже загружен. Перезаписать?", vbOKCancel, _
                        "") = vbOK Then
                    'перезаписываем и убираем старый файл из списка
                    .RemoveItem (i)
                    Exit For
                Else
                    Exit Function
                End If
            End If
        Next i
    End With
    
    fileID = FreeFile() 'file handle
    
    On Error Resume Next:
    Open filename For Input As fileID
    If Err <> 0 Then
        MsgBox "Файл не найден в директории " & filename, vbCritical, "Ошибка"
        zmxRaytraceImport = 0
        Exit Function
    End If
    
    ReDim buffer(1)
    lineCount = 0
    While Not EOF(fileID)
        Line Input #fileID, buffer(lineCount)
        lineCount = lineCount + 1
        ReDim Preserve buffer(lineCount)
    Wend
    Close txtfile
    
    'найдем строку, где написана координата в плоскости предмета
    position = FindString(buffer, "Normalized Y Field Coord (Hy)")
    If position = 0 Then 'если не нашли
        zmxRaytraceImport = 0
        Exit Function
    End If
    'position => Normalized Y Field Coord (Hy)
    parsedString = TableParse(buffer(position))
    Hy = Val(parsedString(6))
    
    ReDim parsedString(7)
    
    'найдем строку, где написана отн. зрачковая координата
    position = FindString(buffer, "Normalized Y Pupil Coord (Py)")
    If position = 0 Then 'если не нашли
        zmxRaytraceImport = 0
        Exit Function
    End If
    'position => Normalized Y Field Coord (Hy)
    parsedString = TableParse(buffer(position))
    Py = Val(parsedString(6))
    
    If rayForm.fileList.ListCount <= 4 Then 'если список неполный
        
        'найдём начало таблицы
        tempPos = FindString(buffer, "Real Ray Trace Data:")
        position = FindStringBetween(buffer, "OBJ", tempPos, tempPos + 5) 'где-то рядом таблица
        If position = 0 Then 'если не нашли таблицу
            zmxRaytraceImport = 0
            Exit Function
        End If
        'а если нашли, то на строке position+1 у нас 1я поверхность
        'найдём конец таблицы
        endPosition = FindString(buffer, "Paraxial Ray Trace Data:") - 1
        If endPosition = 0 Then 'если не нашли таблицу
            zmxRaytraceImport = 0
            Exit Function
        End If
        
'        With rayForm
'            If .fileList.ListCount = 1 Then
'            .status.Caption = "Загрузите ещё один файл." _
'                & " Должны быть загружены файлы Raytrace для лучей: " _
'                & vbCrLf & "апертурный (Hy=0, Py=1), главный (1,0), верхний (1,1), нижний (1,-1)."
'            Else
'            .status.Caption = "Загрузите ещё " & 4 - .fileList.ListCount & " файла." _
'                & "Нужно загрузить файлы Raytrace для следующих лучей: " _
'                & "апертурный (Hy=0, Py=1), главный (1,0), верхний (1,1), нижний (1,-1)."
'            End If
'            .openBtn.Enabled = True
'        End With
        
        If Hy = 1 And Py = 0 Then 'chief
            If UBound(rays) = 1 Then 'if empty
                ReDim rays(1)
            End If
            surfcount = 0
            For i = position + 1 To endPosition
                If Not buffer(i) = "" Then
                    surfcount = surfcount + 1
                    If UBound(rays) < surfcount Then
                        ReDim Preserve rays(surfcount)
                    End If
                    parsedString = TableParse(buffer(i))
                    If i = position + 1 Then 'если мы на 1й поверхности
                    'тогда определим Y-Cosine
                        fieldCos = Val(parsedString(5))
                    End If
                    rays(surfcount).chiefRayH = parsedString(2)
                Else
                    Exit For
                End If
            Next i
            'сохраняем информацию об источнике текущей
            rays(0).chiefRayH = filename
            
            With rayForm
                .status.Caption = "В системе " & surfcount & " поверхностей. " '& .status.Caption
                With .fileList
                    .AddItem
                    .List(.ListCount - 1, 0) = "главный"
                    .List(.ListCount - 1, 1) = "Hy=1"
                    .List(.ListCount - 1, 2) = "Py=0"
                    .List(.ListCount - 1, 3) = filename
                End With
            End With
            
            zmxRaytraceImport = surfcount
            Exit Function
        End If
        
        If Hy = 1 And Py = 1 Then 'upper
            If UBound(rays) = 1 Then 'if empty
                ReDim rays(1)
            End If
            surfcount = 0
            For i = position + 1 To endPosition
                If Not buffer(i) = "" Then
                    surfcount = surfcount + 1
                    If UBound(rays) < surfcount Then
                        ReDim Preserve rays(surfcount)
                    End If
                    parsedString = TableParse(buffer(i))
                    rays(surfcount).upperRayH = parsedString(2)
                Else
                    Exit For
                End If
            Next i
            'сохраняем информацию об источнике текущей
            rays(0).upperRayH = filename
            
            With rayForm
                .status.Caption = "В системе " & surfcount & " поверхностей. " '& .status.Caption
                With .fileList
                    .AddItem
                    .List(.ListCount - 1, 0) = "верхний"
                    .List(.ListCount - 1, 1) = "Hy=1"
                    .List(.ListCount - 1, 2) = "Py=1"
                    .List(.ListCount - 1, 3) = filename
                End With
            End With
            
            zmxRaytraceImport = surfcount
            Exit Function
        End If
        
        If Hy = 1 And Py = -1 Then 'lower
            If UBound(rays) = 1 Then 'if empty
                ReDim rays(1)
            End If
            surfcount = 0
            For i = position + 1 To endPosition
                If Not buffer(i) = "" Then
                    surfcount = surfcount + 1
                    If UBound(rays) < surfcount Then
                        ReDim Preserve rays(surfcount)
                    End If
                    parsedString = TableParse(buffer(i))
                    rays(surfcount).lowerRayH = parsedString(2)
                Else
                    Exit For
                End If
            Next i
            'сохраняем информацию об источнике текущей
            rays(0).lowerRayH = filename
            
            With rayForm
                .status.Caption = "В системе " & surfcount & " поверхностей. " '& .status.Caption
                With .fileList
                    .AddItem
                    .List(.ListCount - 1, 0) = "нижний"
                    .List(.ListCount - 1, 1) = "Hy=1"
                    .List(.ListCount - 1, 2) = "Py=-1"
                    .List(.ListCount - 1, 3) = filename
                End With
            End With
            
            zmxRaytraceImport = surfcount
            Exit Function
        End If
        
        If Hy = 0 And Py = 1 Then 'axial
            If UBound(rays) = 1 Then 'if empty
                ReDim rays(1)
            End If
            surfcount = 0
            For i = position + 1 To endPosition
                If Not buffer(i) = "" Then
                    surfcount = surfcount + 1
                    If UBound(rays) < surfcount Then
                        ReDim Preserve rays(surfcount)
                    End If
                    parsedString = TableParse(buffer(i))
                    rays(surfcount).axialRayH = parsedString(2)
                Else
                    Exit For
                End If
            Next i
            'сохраняем информацию об источнике текущей
            rays(0).axialRayH = filename
            
            With rayForm
                .status.Caption = "В системе " & surfcount & " поверхностей. " '& .status.Caption
                With .fileList
                    .AddItem
                    .List(.ListCount - 1, 0) = "апертурный"
                    .List(.ListCount - 1, 1) = "Hy=0"
                    .List(.ListCount - 1, 2) = "Py=1"
                    .List(.ListCount - 1, 3) = filename
                End With
            End With
            
            zmxRaytraceImport = surfcount
            Exit Function
        End If
    Else
        With rayForm
            '.status.Caption = "Все 4 файла загружены. Удалите файл(ы), чтобы загрузить новые."
            .status.ForeColor = RGB(150, 0, 0)
            .openBtn.Enabled = False
        End With
    End If
    
End Function

Sub rayDataCleanup()
    
End Sub

Sub rayFillTable()
    
Dim rayStartCell As String
Dim headerStartCell As String
Dim raySheet As Object
Dim i As Integer
Dim deg As Integer
Dim min As Integer
Dim sec As Integer
Dim halfField As Double

halfField = ArcSin(fieldCos) * 180 / 3.1415926
deg = Fix(halfField)
min = Int((Abs(halfField - deg)) * 60) '>0
sec = Int(((Abs(halfField - deg)) * 60 _
        - Int((Abs(halfField - deg)) * 60)) * 60)
        
Application.ScreenUpdating = False
        'rayStartCell = rayForm.startCell.text
    With rayForm
        If .createSheetChk.value = True Then
            Set raySheet = Application.Worksheets.Add
            On Error Resume Next
            raySheet.name = .sheetName.text
            If Err.Number <> 0 Then
                MsgBox Err.Description
                Exit Sub
            End If
        Else
            Set raySheet = Application.ActiveSheet
            .sheetName.text = raySheet.name
        End If
        
        If .headerChk.value = True Then
            headerStartCell = .startCell.text
            rayStartCell = raySheet.Range(headerStartCell).Offset(3, 0).Address
            raySheet.Range(headerStartCell).Resize(3 + UBound(rays), 5).UnMerge
            'generate header
            With raySheet.Range(headerStartCell)
                .Resize(3 + UBound(rays), 5).UnMerge
                
                With .Offset(0, 1)
                    .Resize(1, 4).Merge
                    .MergeArea.value = "Высоты, мм"
                    .HorizontalAlignment = xlCenter
                End With
                
                With .Offset(1, 1)
                    .Resize(2, 1).Merge
                    .MergeArea.value = "Осевой луч"
                    .HorizontalAlignment = xlCenter
                    .WrapText = True
                    .MergeArea.Columns.AutoFit
                End With
                
                With .Offset(1, 2)
                    .Resize(1, 3).Merge
                    .MergeArea.value = "Наклонный луч " _
                        & ChrW(969) & " = " & deg & ChrW(176) & min & "'" & sec & "''"
                    .HorizontalAlignment = xlCenter
                End With
                
                .Offset(2, 2).value = "Нижний"
                .Offset(2, 3).value = "Главный"
                .Offset(2, 4).value = "Верхний"
                
                .Resize(3, 1).Merge
                .MergeArea.value = "Номер поверхности"
                .HorizontalAlignment = xlCenter
                .WrapText = True
                '.MergeArea.Columns.Width
                
                .Offset(0, 1).Resize(3 + UBound(rays), 4).NumberFormat = "0.00"
                .Resize(3 + UBound(rays), 1).NumberFormat = "0"
                .Resize(3 + UBound(rays), 5).HorizontalAlignment = xlCenter
            End With
        Else
            rayStartCell = .startCell.text
            With raySheet.Range(rayStartCell)
                .Resize(3 + UBound(rays), 5).UnMerge
                .Offset(0, 1).Resize(3 + UBound(rays), 4).NumberFormat = "0.00"
                .Resize(3 + UBound(rays), 1).NumberFormat = "0"
            End With
        End If
    End With
    
    For i = 1 To UBound(rays)
        With raySheet.Range(rayStartCell)
            .Offset(i - 1, 0).value = i 'номер п/п
            .Offset(i - 1, 1).value = Round(Val(rays(i).axialRayH), 2)
            .Offset(i - 1, 2).value = Round(Val(rays(i).lowerRayH), 2)
            .Offset(i - 1, 3).value = Round(Val(rays(i).chiefRayH), 2)
            .Offset(i - 1, 4).value = Round(Val(rays(i).upperRayH), 2)
        End With
    Next i
    
    Application.ScreenUpdating = True
End Sub
Public Function ArcCos(A As Double) As Double 'в радианах
  'Inverse Cosine
    On Error Resume Next
        If A = 1 Then
            ArcCos = 0
            Exit Function
        End If
        ArcCos = Atn(-A / Sqr(-A * A + 1)) + 2 * Atn(1)
    On Error GoTo 0
End Function
Public Function ArcSin(ByVal x As Double) As Double 'в радианах
    If x = 1 Then
        ArcSin = 0
        Exit Function
    Else
        ArcSin = Atn(x / Sqr(-x * x + 1))
    End If
End Function

