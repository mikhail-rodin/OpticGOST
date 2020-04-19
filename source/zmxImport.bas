Attribute VB_Name = "zmxImport"
Public Type surfData '�����������
    id As String    '����� ���� IMA/STO/OBJ
    type As String
    r As Double
    n As Double
    nShort As Double 'n ��� �������� ����� �����
    nLong As Double 'n ��� ������� ����� �����
    d As Double
    v As Double     '����. ����
    glass As String
    diam As Double      '�������� �������
    sag As Double       '������� ������� �� ��������� ��������
End Type
Dim surface() As surfData
Dim stopPos As Integer '��������� ���������

Public Type rayData
'������ ���������� � 1
'�� ������ 0 - �������� (��� �����)
    axialRayH As String '������ ������� ����
    lowerRayH As String '������ ������� ����������� ����
    upperRayH As String '������ �������� ����������� ����
    chiefRayH As String '������ �������� ����
End Type
Public rays() As rayData
Public fieldCos As Double

Dim lineArray() As String '���� ����� �������� ������
Public surfCounter, elementCounter As Integer '����� ������������, �������������� ��� ������� �����
Dim shortwave As Double '�������� ����� (��� ���������� ����� ����)
Dim longwave As Double '������� ����� (��� ���������� ����� ����)
Public waveSelection, shortWaveSel, longWaveSel As Integer '����� ��������� ����� �����
Const rndDataStartString As String = "SURFACE DATA SUMMARY:"
Const firstSurfaceString As String = "OBJ"
Const lastSurfaceString As String = "IMA"
Const glassIndexStartString As String = "INDEX OF REFRACTION DATA:"
Const thermalStartString As String = "THERMAL COEFFICIENT OF EXPANSION DATA:"



Function FirstNotDigit(str As String) As Boolean
    Static code As Integer
    code = Asc(Mid(str, 1, 1)) '��������� �� ASCII ���� ������� �������
    Select Case code
        Case 48 To 57: FirstNotDigit = False
        Case Else: FirstNotDigit = True
    End Select
End Function

Function LZOStranslate(glass As String) As String
    Static temp As String
    If InStr(1, glass, "LZ_") Then
        temp = Replace(glass, "LZ_", "")
        temp = Replace(temp, "F", "�")
        temp = Replace(temp, "B", "�")
        temp = Replace(temp, "L", "�")
        LZOStranslate = temp
    Else
        LZOStranslate = glass
    End If
End Function

Function TableParse(row As String) As String()
'�� ����� ������� ������� �� ����� Zemax Prescription Data
'�� ������ ������������ ������ �� ���� � �������
'��� ������� ���������
    Dim result() As String
    Dim temp(20) As String
    Dim tempLen As Integer
  '            Do
'              tempStr = lineArray(rowCounter)
'              lineArray(i) = Replace(lineArray(i), "  ", " ") 'remove multiple white spaces
'            Loop Until tempStr = lineArray(rowCounter)
  
    'ReDim result(1)
    result = Split(Trim(row), " ") 'trim to remove starting/trailing space
    '������ � ��� ������� �������� � ������ ��� ��������� ��������
    '��� �, �������� ������ �� ����� �����
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
'���������� ����� ������ ��� 0 � ������ ������
    Static pos As Integer '����� ������
    pos = 0
    Do Until pos = UBound(text) '���� "SURFACE DATA SUMMARY:"
        If Not InStr(text(pos), str) = 0 Then
            Exit Do
        End If
        pos = pos + 1
    Loop
    
    If pos = UBound(text) Then '���� ������ �� �����
        FindString = 0
    Else
        FindString = pos
    End If
End Function

Function FindStringBetween(text() As String, str As String, _
        ByVal searchStart As Integer, ByVal searchEnd As Integer) As Integer
    '���������� ����� ������ ��� 0 � ������ ������
    Static pos As Integer '����� ������
  
    For pos = searchStart To searchEnd
        If Not InStr(text(pos), str) = 0 Then
            Exit For
        End If
    Next pos
    
    If pos = UBound(text) Then '���� ������ �� �����
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
    Dim splitline() As String '� ������ ������� ��������� ������ �� �����
    Dim tempArray(100) As String
    Dim rowCounter, lineCounter, rndDataPos, i, j, waveNumber, tempArrLen As Integer '����� ������ � �������
    Dim glassIndexDataPos As Integer
    Dim tempStr As String
    Dim txtfile As Integer
    Dim wavelenght As Double

    txtfile = FreeFile() 'file handle
    
    filename = rndForm.filePath.text
    On Error Resume Next:
    Open filename For Input As txtfile
    If Err <> 0 Then
        MsgBox "���� �� ������ � ���������� " & filename, vbCritical, "������"
        Exit Sub
    End If
    
    With rndForm.statusLabel
            .Caption = "���� ������."
            .ForeColor = RGB(10, 120, 10)
    End With
    
    rndForm.rndFillTableBtn.Enabled = True
    
    lineCounter = 0

    ReDim lineArray(1)
    While Not EOF(txtfile) '������ ����� � ������, �� ������ ������ �� ��� ��������� ����� �� ������� �����
        Line Input #txtfile, tempStr
        lineArray(lineCounter) = tempStr
        lineCounter = lineCounter + 1 '���������, ������� � ��� �������
        ReDim Preserve lineArray(lineCounter)
    Wend
    Close txtfile
 
    ReDim Preserve surface(surfCounter) As surfData
    
   ' rndForm.statusLabel.Caption = "���� ������������. � ����� " & lineCounter - 1 & " �����(�)"

    Application.ScreenUpdating = False

    rndDataPos = 0
    Do Until rndDataPos = lineCounter '���� "SURFACE DATA SUMMARY:"

        If Not InStr(lineArray(rndDataPos), rndDataStartString) = 0 Then
            Exit Do
        End If

        rndDataPos = rndDataPos + 1
    Loop
    
    rndForm.textBox.text = ""
    For i = 0 To lineCounter '������� ����� � ������
    rndForm.textBox.text = rndForm.textBox.text & i & " >> " & lineArray(i) & vbCrLf
    Next i

    'MsgBox "������ " & rndDataPos, vbInformation, "������"

'��������� SURFACE DATA SUMMARY
    j = 0
    surfCounter = 0
    i = rndDataPos + 2 '���������� ��������, ������ ������ � ����� �������
    ReDim surface(1)
    Do Until i = lineCounter
        If Not lineArray(rowCounter) = "" Then '���������� ���� ���� ������ ������
            splitline = TableParse(lineArray(i))
                '������ � ������� splitLine ���:
                '0 - ����� �/�
                '1 - ��� �����������
                '2 - ������
                '3 - �������
                '4 - ������ ��� - !!! - ��. ������� ���� ��� ������ - ��� �����������, ������ ��������
                '5 - �������� �������, ���� ���� ������
            If Not InStr(splitline(0), "STO") = 0 Then
            '���� ����������� �� ���������
            '���������� �
            '� ���������� ���������� ����� ��������� � �����. ����� ��
                stopPos = surfCounter
                surface(surfCounter - 1).d = surface(surfCounter - 1).d + Val(splitline(3))
            '�� �������������� surfcounter
            '������� surface(surfCounter) � ����. ����� �������������
            Else
                If FirstNotDigit(splitline(4)) = False Then '���� ������ ������ � ���� "�����" ��� �����
                    '������, �� ���� ��� �� ������
                    splitline(5) = splitline(4) '������, ��� � ��� ��. �������, � �� �������� ��� � (5)
                    splitline(4) = ""
                End If
    
                If Not InStr(splitline(0), lastSurfaceString) = 0 Then
                    Exit Do '���� "IMA", �������
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
                '���� �� �������������
                    surface(surfCounter).d = Val(splitline(3))
                Else
                '���������� ����
                    surface(surfCounter).d = 0
                End If
                
                surface(surfCounter).glass = splitline(4)
                
                surface(surfCounter).diam = Val(splitline(5))
                
                surfCounter = surfCounter + 1 '��������� ��� ���� �����������
            End If
        End If
        i = i + 1 '��� �� ��������� ������
        
    Loop
    
    
    glassIndexDataPos = FindString(lineArray, glassIndexStartString)

    j = 0 '������� � ������ �����������
    For i = glassIndexDataPos + 7 To glassIndexDataPos + 7 + surfCounter
        If Not lineArray(i) = "" Then
            splitline = TableParse(lineArray(i))
            If InStr(splitline(0), "Surf") <> 0 Then '���� �� � �����
            '�������� ������ ���� ����
                waveNumber = 0
                shortwave = 10 '��������� ������ ���/���� ����� �����
                longwave = 0
                rndForm.wavelengthList.Clear
                Do Until 4 + waveNumber = UBound(splitline)
                    If Not splitline(4 + waveNumber) = "" Then
                        wavelength = Val(splitline(4 + waveNumber)) 'String to Double
                        If wavelength < shortwave Then '���� ����� �������� �����
                            shortwave = wavelength
                            shortWaveSel = waveNumber
                        End If
                        If wavelength > longwave Then
                            longwave = wavelength
                            longWaveSel = waveNumber
                        End If
                        rndForm.wavelengthList.AddItem (waveNumber & " - " & Left(splitline(4 + waveNumber), 5) & " ���")
                    End If
                    waveNumber = waveNumber + 1
                Loop
                Exit For
            End If

'            If FirstNotDigit(splitline(1)) = True Then '���� ������ ����
'                surface(j).n = splitline(4 + waveSelection) 'n �� 5-� ����� (������ ������� 4)
'            Else '� ���� ������ ���
'                surface(j).n = splitline(3 + waveSelection) '����� �� 4-�
'            End If
            j = j + 1
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    
    rndForm.shortwaveLabel.Caption = "�������� �����: " & shortWaveSel & " - " & shortwave & " ���"
    rndForm.longwaveLabel.Caption = "������� �����: " & longWaveSel & " - " & longwave & " ���"
    
    With rndForm.statusLabel
            .Caption = "������������: " & surfCounter & _
                ", ������� ��������� �������� (IMA) � ����������� (IMA)" & vbCrLf & _
                "���� ����: " & waveNumber
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
    '��������� ������ ������ �������
    rndForm.ZemaxStart.text = Range(rndForm.ESKDstart.text).Offset(surfCounter * 2 + 3, 0).Address
    rndForm.lensStart.text = Range(rndForm.ESKDstart.text).Offset(surfCounter * 3 + 3, 0).Address
    


End Sub

Public Sub rndFillTable()
    Dim i, j As Integer
    Dim ZemaxStartCell As String
    Dim ESKDstartCell As String
    Dim sheetID As String
    Dim LZOS, rndSheet As Boolean
    Dim wsheet As Object
    Static d As Double
    
    
    LZOS = rndForm.LZOSchk.Value
    rndSheet = rndForm.createSheetChk.Value
    
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
    
    If rndForm.generateZemaxTableChk.Value = True Then
        For i = 0 To surfCounter - 1 '������� ���� �� ��������� � ������
            With wsheet
                .Range(ZemaxStartCell).Offset(i, 0).Value = surface(i).r
                .Range(ZemaxStartCell).Offset(i, 1).Value = surface(i).d
                .Range(ZemaxStartCell).Offset(i, 2).Value = surface(i).n
            End With
            
            If surface(i).v = 0 Then '����� �� ���� �����
                wsheet.Range(ZemaxStartCell).Offset(i, 3).Value = ""
            Else
                wsheet.Range(ZemaxStartCell).Offset(i, 3).Value = surface(i).v
            End If
                    
            With wsheet
                .Range(ZemaxStartCell).Offset(i, 4).Value = surface(i).glass
                .Range(ZemaxStartCell).Offset(i, 5).Value = surface(i).diam
                .Range(ZemaxStartCell).Offset(i, 6).Value = surface(i).sag
            End With
        Next i
    End If
    
    If rndForm.generateESKDchk.Value = True Then
        
        With wsheet.Range(ESKDstartCell)
            .Offset(0, 2) = "ne"
            .Offset(0, 2).Characters(2, 1).Font.Subscript = True
            
            .Offset(0, 3) = "ve" 'ChrW(55349) & "e"
            .Offset(0, 3).Characters(2, 1).Font.Subscript = True
            '.Offset(0, 3).Characters(1, 1).Font.Italic = True
            
            .Offset(0, 4) = "����� ������"
            .Offset(0, 5) = ChrW(216) & " ��."
            .Offset(0, 6) = "������� �� " & ChrW(216) & "  ��."
        End With
        
        j = 1
        For i = 1 To (surfCounter - 2) * 2 Step 2
            With wsheet.Range(ESKDstartCell)
                If i = 1 And surface(j).d = 0 Then
                '�� ��������� �����. �� ����. ���������� ��������
                '��� �������, ���� �������������, ������������ 0
                   .Offset(i, 1).Value = ""
                Else
                    d = surface(j).d
                    .Offset(i, 1).Value = "d" & j - 1 & " = " & Round(d, 2)
                    .Offset(i, 1).Characters(2, 2).Font.Subscript = True
                End If

                .Offset(i, 2).Value = "n" & j - 1 & " = " & Round(surface(j).n, 4)
                .Offset(i, 2).Characters(2, 2).Font.Subscript = True

                If surface(j).v = 0 Then '����� �� ���� �����
                    .Offset(i, 3).Value = ""
                Else
                    .Offset(i, 3).Value = Round(surface(j).v, 2)
                    .Offset(i, 3).NumberFormat = "0.00"
                End If

                If LZOS = True Then
                    .Offset(i, 4).Value = LZOStranslate(surface(j).glass)
                Else
                    .Offset(i, 4).Value = surface(j).glass
                End If
            End With
                j = j + 1
        Next i

        j = 1
        For i = 0 To (surfCounter - 2) * 2 + 1 Step 2
            With wsheet.Range(ESKDstartCell)
            If j >= 2 Then
                 .Offset(i, 5).Value = Round(surface(j).diam, 2)
                 .Offset(i, 5).NumberFormat = "0.00"
                 
                 .Offset(i, 6).Value = Round(surface(j).sag, 2)
                 .Offset(i, 6).NumberFormat = "0.00"
            End If
            
            If Not j = 1 Then '������ ��������� �������� �� ���������
                .Offset(i, 0).Value = "r" & j - 1 & " = " & Round(surface(j).r, 2)
                .Offset(i, 0).Characters(2, 2).Font.Subscript = True
            End If
            End With
                j = j + 1
        Next i
        
        With wsheet.Range(ESKDstartCell).Resize(surfCounter * 2 + 5, 7) '��� ���� �������
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
    
    j = 0 '������� � ������ �����������
    For i = glassIndexDataPos + 7 To glassIndexDataPos + 7 + surfCounter
        If Not lineArray(i) = "" Then '���� ������ ����� �� ������
            splitstr = TableParse(lineArray(i))
            If Not i = glassIndexDataPos + 7 + stopPos Then
            '���� � ������ �� ���������
                If Not InStr(splitstr(0), "Surf") <> 0 Then '���� �� �� � �����
                    If FirstNotDigit(splitstr(1)) = True Then '���� ������ ����
                        surface(j).n = Val(splitstr(4 + waveSelection)) 'n �� 5-� ����� (������ ������� 4)
                        surface(j).nShort = Val(splitstr(4 + shortWaveSel))
                        surface(j).nLong = Val(splitstr(4 + longWaveSel))
                    Else '� ���� ������ ���
                        surface(j).n = Val(splitstr(3 + waveSelection)) '����� �� 4-�
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
        If surface(i).n > 1 Then '���� �� ����� ���� �� �������
            surface(i).v = (surface(i).n - 1) / (surface(i).nShort - surface(i).nLong)
        End If
    Next i
End Sub

Public Sub CalculateSag()
Static i As Integer
Static d, r As Double

    For i = 0 To UBound(surface)
        If surface(i).r = 0 Then
            surface(i).sag = 0
        Else
            d = surface(i).diam
            r = surface(i).r
            If r > 0 Then
                surface(i).sag = r - Sqr(r ^ 2 - (d / 2) ^ 2)
            Else
                surface(i).sag = r + Sqr(r ^ 2 - (d / 2) ^ 2)
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
Static i As Integer

    Application.ScreenUpdating = False
    
    lensStartCell = rndForm.lensStart.text
    
    If rndForm.newLensSheetchk.Value = True Then
        Set lensSheet = Application.Worksheets.Add
        lensSheet.name = rndForm.lensSheetNameBox.text
    Else
        Set lensSheet = Application.ActiveSheet
        rndForm.lensSheetNameBox.text = lensSheet.name
    End If
    
    With lensSheet.Range(lensStartCell)
        .Offset(0, 0) = "� ���. ���."
        .Offset(0, 1) = ChrW(216) & " ��."
        .Offset(0, 2) = "������� �� " & ChrW(216) & "  ��."
        .Offset(0, 3) = ChrW(216) & " ��."
        .Offset(0, 4) = "������� �� " & ChrW(216) & "  ��."
        .Offset(0, 5) = "������� �� ���"
    End With
    
    elementCounter = 0
    For i = 0 To surfCounter - 1
        If Not (surface(i).glass = "" Or surface(i).glass = "Glass") Then
        '���� �� ���������� �� ����� � ��� �� ���������
            elementCounter = elementCounter + 1 '��������� �
            With lensSheet.Range(lensStartCell)
                .Offset(elementCounter, 0).Value = elementCounter '����� �/�
                .Offset(elementCounter, 1).Value = Round(surface(i).diam, 2)
                .Offset(elementCounter, 2).Value = Round(surface(i).sag, 2)
                .Offset(elementCounter, 3).Value = Round(surface(i + 1).diam, 2)
                .Offset(elementCounter, 4).Value = Round(surface(i + 1).sag, 2)
                .Offset(elementCounter, 5).Value = Round(surface(i).d, 2)
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
    '��������� ����� ������������
    '��������� 0 ��� ������
    
    Static surfcount, lineCount, fileID, i As Integer
    Static position, tempPos, endPosition As Integer
    Static Hy, Py As Double
    Static field As Double '������������ �� Y-cosine ��� ������ �����������
    
    Static rayType As Integer
    '1 - �������
    '2 - �������
    '3 - ������
    '4 - ����������
    
    Dim buffer() As String '���� �������� ����
    Dim parsedString() As String '����� ������ ������� ������ ��������
    
    '��������, �������� �� ��� ���� ����
    With rayForm.fileList
        For i = 0 To .ListCount - 1
            If .List(i, 3) = filename Then
            '���� ��� ���� ����� ����
                If MsgBox("���� " & Right(filename, Len(filename) - InStrRev(filename, "\")) _
                    & " ��� ��������. ������������?", vbOKCancel, _
                        "") = vbOK Then
                    '�������������� � ������� ������ ���� �� ������
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
        MsgBox "���� �� ������ � ���������� " & filename, vbCritical, "������"
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
    
    '������ ������, ��� �������� ���������� � ��������� ��������
    position = FindString(buffer, "Normalized Y Field Coord (Hy)")
    If position = 0 Then '���� �� �����
        zmxRaytraceImport = 0
        Exit Function
    End If
    'position => Normalized Y Field Coord (Hy)
    parsedString = TableParse(buffer(position))
    Hy = Val(parsedString(6))
    
    ReDim parsedString(7)
    
    '������ ������, ��� �������� ���. ��������� ����������
    position = FindString(buffer, "Normalized Y Pupil Coord (Py)")
    If position = 0 Then '���� �� �����
        zmxRaytraceImport = 0
        Exit Function
    End If
    'position => Normalized Y Field Coord (Hy)
    parsedString = TableParse(buffer(position))
    Py = Val(parsedString(6))
    
    If rayForm.fileList.ListCount <= 4 Then '���� ������ ��������
        
        '����� ������ �������
        tempPos = FindString(buffer, "Real Ray Trace Data:")
        position = FindStringBetween(buffer, "OBJ", tempPos, tempPos + 5) '���-�� ����� �������
        If position = 0 Then '���� �� ����� �������
            zmxRaytraceImport = 0
            Exit Function
        End If
        '� ���� �����, �� �� ������ position+1 � ��� 1� �����������
        '����� ����� �������
        endPosition = FindString(buffer, "Paraxial Ray Trace Data:") - 1
        If endPosition = 0 Then '���� �� ����� �������
            zmxRaytraceImport = 0
            Exit Function
        End If
        
'        With rayForm
'            If .fileList.ListCount = 1 Then
'            .status.Caption = "��������� ��� ���� ����." _
'                & " ������ ���� ��������� ����� Raytrace ��� �����: " _
'                & vbCrLf & "���������� (Hy=0, Py=1), ������� (1,0), ������� (1,1), ������ (1,-1)."
'            Else
'            .status.Caption = "��������� ��� " & 4 - .fileList.ListCount & " �����." _
'                & "����� ��������� ����� Raytrace ��� ��������� �����: " _
'                & "���������� (Hy=0, Py=1), ������� (1,0), ������� (1,1), ������ (1,-1)."
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
                    If i = position + 1 Then '���� �� �� 1� �����������
                    '����� ��������� Y-Cosine
                        fieldCos = Val(parsedString(5))
                    End If
                    rays(surfcount).chiefRayH = parsedString(2)
                Else
                    Exit For
                End If
            Next i
            '��������� ���������� �� ��������� �������
            rays(0).chiefRayH = filename
            
            With rayForm
                .status.Caption = "� ������� " & surfcount & " ������������. " '& .status.Caption
                With .fileList
                    .AddItem
                    .List(.ListCount - 1, 0) = "�������"
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
            '��������� ���������� �� ��������� �������
            rays(0).upperRayH = filename
            
            With rayForm
                .status.Caption = "� ������� " & surfcount & " ������������. " '& .status.Caption
                With .fileList
                    .AddItem
                    .List(.ListCount - 1, 0) = "�������"
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
            '��������� ���������� �� ��������� �������
            rays(0).lowerRayH = filename
            
            With rayForm
                .status.Caption = "� ������� " & surfcount & " ������������. " '& .status.Caption
                With .fileList
                    .AddItem
                    .List(.ListCount - 1, 0) = "������"
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
            '��������� ���������� �� ��������� �������
            rays(0).axialRayH = filename
            
            With rayForm
                .status.Caption = "� ������� " & surfcount & " ������������. " '& .status.Caption
                With .fileList
                    .AddItem
                    .List(.ListCount - 1, 0) = "����������"
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
            '.status.Caption = "��� 4 ����� ���������. ������� ����(�), ����� ��������� �����."
            .status.ForeColor = RGB(150, 0, 0)
            .openBtn.Enabled = False
        End With
    End If
    
End Function

Sub rayDataCleanup()
    
End Sub

Sub rayFillTable()
    
Static rayStartCell, headerStartCell As String
Dim raySheet As Object
Static i, deg, min, sec As Integer
Static halfField As Double

halfField = ArcSin(fieldCos) * 180 / 3.1415926
deg = Fix(halfField)
min = Int((Abs(halfField - deg)) * 60) '>0
sec = Int(((Abs(halfField - deg)) * 60 _
        - Int((Abs(halfField - deg)) * 60)) * 60)

    Application.ScreenUpdating = False
    
    
        'rayStartCell = rayForm.startCell.text
    
    With rayForm
        If .createSheetChk.Value = True Then
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
        
        If .headerChk.Value = True Then
            headerStartCell = .startCell.text
            rayStartCell = raySheet.Range(headerStartCell).Offset(3, 0).Address
            raySheet.Range(headerStartCell).Resize(3 + UBound(rays), 5).UnMerge
            'generate header
            With raySheet.Range(headerStartCell)
                .Resize(3 + UBound(rays), 5).UnMerge
                
                With .Offset(0, 1)
                    .Resize(1, 4).Merge
                    .MergeArea.Value = "������, ��"
                    .HorizontalAlignment = xlCenter
                End With
                
                With .Offset(1, 1)
                    .Resize(2, 1).Merge
                    .MergeArea.Value = "������ ���"
                    .HorizontalAlignment = xlCenter
                    .WrapText = True
                End With
                
                With .Offset(1, 2)
                    .Resize(1, 3).Merge
                    .MergeArea.Value = "��������� ��� " _
                        & ChrW(969) & " = " & deg & ChrW(176) & min & "'" & sec & "''"
                    .HorizontalAlignment = xlCenter
                End With
                
                .Offset(2, 2).Value = "������"
                .Offset(2, 3).Value = "�������"
                .Offset(2, 4).Value = "�������"
                
                .Resize(3, 1).Merge
                .MergeArea.Value = "����� �����������"
                .HorizontalAlignment = xlCenter
                .WrapText = True
                
                .Offset(0, 1).Resize(3 + UBound(rays), 4).NumberFormat = "0.00"
                .Resize(3 + UBound(rays), 1).NumberFormat = "0"
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
            .Offset(i - 1, 0).Value = i '����� �/�
            .Offset(i - 1, 1).Value = Round(Val(rays(i).axialRayH), 2)
            .Offset(i - 1, 2).Value = Round(Val(rays(i).lowerRayH), 2)
            .Offset(i - 1, 3).Value = Round(Val(rays(i).chiefRayH), 2)
            .Offset(i - 1, 4).Value = Round(Val(rays(i).upperRayH), 2)
        End With
    Next i
    
    
    
    Application.ScreenUpdating = True

End Sub

Public Function ArcCos(A As Double) As Double '� ��������
  'Inverse Cosine
    On Error Resume Next
        If A = 1 Then
            ArcCos = 0
            Exit Function
        End If
        ArcCos = Atn(-A / Sqr(-A * A + 1)) + 2 * Atn(1)
    On Error GoTo 0
End Function

Public Function ArcSin(ByVal x As Double) As Double '� ��������
    If x = 1 Then
        ArcSin = 0
        Exit Function
    Else
        ArcSin = Atn(x / Sqr(-x * x + 1))
    End If
End Function

