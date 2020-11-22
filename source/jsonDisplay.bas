Attribute VB_Name = "jsonDisplay"
Option Base 0
Option Explicit
Public Sub printStatus(ByVal info As String)
    With jsonForm.status
        .Caption = .text & vbCrLf & info
    End With
End Sub
Public Sub rinseStatus()
    With jsonForm.status
        .Caption = ""
    End With
End Sub
Public Sub addWaves(ByRef lens As CLens)
    Dim i As Integer
    With jsonForm.waveList
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                Call lens.addWave(i + 1)
            End If
        Next i
    End With
End Sub
Public Sub addFields(ByRef lens As CLens)
    Dim i As Integer
    With jsonForm.fieldList
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                Call lens.addField(i + 1)
            End If
        Next i
    End With
End Sub
Public Sub refreshWaves(lens As CLens)
    Dim i As Integer
    With jsonForm.waveList
        .Clear
        For i = 0 To lens.waveCount - 1
            .AddItem
            .List(i, 0) = CStr(i + 1)
            .List(i, 1) = CStr(1000 * lens.wavelength(i + 1)) + " нм"
            .List(i, 2) = optics.SpectralLine(1000 * lens.wavelength(i + 1))
        Next i
    End With
    
    Dim selection() As Integer
    selection = lens.selectedWaves
    With jsonForm.waveSel
        .Clear
        If UBound(selection) > 0 Then
        'if something's selected
            For i = 0 To UBound(selection) - 1
                .AddItem
                .List(i, 0) = CStr(selection(i + 1))
                .List(i, 1) = CStr(1000 * lens.wavelength(selection(i + 1))) + " нм"
                .List(i, 2) = optics.SpectralLine(1000 * lens.wavelength(selection(i + 1)))
            Next i
        End If
    End With
End Sub
Public Sub delWaves(lens As CLens)
    Dim i As Integer
    Dim waveNo As Integer
    With jsonForm.waveSel
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                waveNo = Int(.List(i, 0))
                Call lens.delWave(waveNo)
            End If
        Next i
    End With
End Sub
Public Sub delFields(lens As CLens)
    Dim i As Integer
    Dim fieldNo As Integer
    With jsonForm.fieldSel
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                fieldNo = Int(.List(i, 0))
                Call lens.delField(fieldNo)
            End If
        Next i
    End With
End Sub
Public Sub refreshFields(lens As CLens)
    Dim postfix As String
    If lens.field_type = 0 Then
        postfix = ChrW(176) 'degrees dot
    Else
        postfix = " " + lens.units
    End If
    
    Dim isZeroField As Boolean
    Const eps As Double = 0.05
    Dim i As Integer
    With jsonForm.fieldList
        .Clear
        For i = 0 To lens.fieldCount - 1
            .AddItem
            .List(i, 0) = CStr(i + 1)
            .List(i, 1) = CStr(Round(lens.srcHy(i + 1), 2))
            .List(i, 2) = CStr(Round(lens.srcHx(i + 1), 2))
            .List(i, 3) = CStr(Round(lens.srcYField(i + 1), 2)) + postfix
            .List(i, 4) = CStr(Round(lens.srcXField(i + 1), 2)) + postfix
        Next i
    End With
    
    Dim selection() As Integer
    selection = lens.selectedFields
    With jsonForm.fieldSel
        .Clear
        If UBound(selection) > 0 Then
        'if something's selected
            For i = 0 To UBound(selection) - 1
                .AddItem
                .List(i, 0) = CStr(selection(i + 1))
                .List(i, 1) = CStr(Round(lens.srcHy(selection(i + 1)), 2))
                .List(i, 2) = CStr(Round(lens.srcHx(selection(i + 1)), 2))
                .List(i, 3) = CStr(Round(lens.srcYField(selection(i + 1)), 2)) + postfix
                .List(i, 4) = CStr(Round(lens.srcXField(selection(i + 1)), 2)) + postfix
            Next i
        End If
    End With
End Sub
Public Sub fillAberTable(ByRef lens As CLens, startCell As Excel.Range)
    Dim OPD As Boolean
    OPD = jsonForm.OPDchk.value
    
    Application.ScreenUpdating = False
    
    Dim chiefOffset As Integer, tangOffset As Integer, sagOffset, fullHeigth As Integer
    chiefOffset = printAxial(startCell, lens)
    tangOffset = chiefOffset + 1 + printChief(startCell.Offset(chiefOffset, 0), lens)
    sagOffset = tangOffset + 2 + printTang(startCell.Offset(tangOffset, 0), lens)
    fullHeigth = sagOffset + 1 + printSag(startCell.Offset(sagOffset, 0), lens)
    
    With startCell.Resize(fullHeigth * 2, 100)
        .HorizontalAlignment = xlCenter
    End With
    
    Application.ScreenUpdating = True
End Sub
Private Function printChief(startCell As Excel.Range, ByRef lens As CLens) As Integer
'returns printed row count
    Dim uOMEGA As String
    uOMEGA = ChrW(969)
    Dim uLAMBDA As String
    uLAMBDA = ChrW(955)
    Dim uINCR As String
    uINCR = ChrW(8710)
    
    Dim imSizeSymbol As String
    If lens.afocal Then
        imSizeSymbol = uOMEGA
    Else
        imSizeSymbol = "y"
    End If
    
    With startCell
        Dim offsetDown As Integer
        'offset of second row to fit 'lambda e'
        If lens.selectedWaveCount = 1 Then
            offsetDown = 0
        Else
            offsetDown = 1
            With .Offset(0, 2)
                .Resize(1, 7).Merge
                .MergeArea.value = uLAMBDA & waveLetter(lens, 1)
                .Characters(Start:=2, Length:=2).Font.Subscript = True
            End With
            Dim wave As Integer
            Dim field As Integer
            For wave = 2 To lens.selectedWaveCount
                With .Offset(0, 9 - 2 + wave)
                'print image sizes
                    .Resize(3, 1).Merge
                    .MergeArea.value = imSizeSymbol & "'" & waveLetter(lens, wave)
                    .Characters(Start:=3, Length:=2).Font.Subscript = True
                    .VerticalAlignment = xlCenter
                    For field = 1 To lens.selectedFieldCount
                        If lens.afocal Then
                            .Offset(field, 0).value = tools.degMinSec( _
                                tools.Deg(lens.chiefRayAngle(wave, field)))
                        Else
                            .Offset(field, 0).value = _
                                Round(lens.chiefOpVal("REAY", wave, field), 4)
                        End If
                    Next field
                End With
            Next wave
            For wave = 2 To lens.selectedWaveCount - 1
                With .Offset(0, 9 - 3 + lens.selectedWaveCount + wave)
                'print lateral chroma
                    .Resize(3, 1).Merge
                    .MergeArea.value = imSizeSymbol & "'" & waveLetter(lens, wave) & _
                        "-" & imSizeSymbol & "'" & waveLetter(lens, wave + 1)
                    .Characters(Start:=3, Length:=2).Font.Subscript = True
                    .Characters(Start:=8, Length:=2).Font.Subscript = True
                    .VerticalAlignment = xlCenter
                    For field = 1 To lens.selectedFieldCount
                        If lens.afocal Then
                            .Offset(field, 0).value = tools.degMinSec( _
                                tools.Deg(lens.chiefRayAngle(wave, field) _
                                    - lens.chiefRayAngle(wave + 1, field)))
                        Else
                            .Offset(field, 0).value = Round(lens.chiefOpVal("REAY", wave, field) _
                                - lens.chiefOpVal("REAY", wave + 1, field), 4)
                        End If
                    Next field
                End With
            Next wave
        End If
        With .Offset(offsetDown + 1, 0)
            .Resize(1, 8).NumberFormat = "0.000"
            Dim distortionVal As Double
            Dim zM As Double
            Dim zS As Double
            Dim imSize As Double
            For field = 1 To lens.selectedFieldCount
                distortionVal = lens.chiefOpVal("DISG", 0, field)
                .Offset(field, 0).value = tools.degMinSec(lens.yField(field))
                .Offset(field, 1).value = Round(lens.sP_entr(field, 1), 4)
                .Offset(field, 3).value = Round(lens.sP_exit(field, 1), 4)
                .Offset(field, 5).value = Round(distortionVal, 2)
                zM = lens.chiefOpVal("FCGT", 0, field)
                zS = lens.chiefOpVal("FCGS", 0, field)
                If lens.afocal Then
                    imSize = lens.chiefRayAngle(1, field) 'in radians
                    .Offset(field, 2).value = tools.degMinSec(tools.Deg(imSize))
                    .Offset(field, 4).value = tools.degMinSec( _
                        distortionVal * tools.Deg(imSize) / 100)
                    .Offset(field, 6).value = Round(zM * Cos(imSize), 4)
                    .Offset(field, 7).value = Round(zS * Cos(imSize), 4)
                    .Offset(field, 8).value = Round((zM - zS) * Cos(imSize), 4)
                Else
                    imSize = lens.chiefOpVal("REAY", 1, field)
                    .Offset(field, 2).value = Round(imSize, 4)
                    .Offset(field, 4).value = Round(distortionVal * imSize / 100, 4)
                    .Offset(field, 6).value = Round(zM, 3)
                    .Offset(field, 7).value = Round(zS, 3)
                    .Offset(field, 8).value = Round(zM - zS, 3)
                End If
            Next field
        End With
        With .Offset(offsetDown, 2)
            .Resize(2, 1).Merge
            .MergeArea.value = imSizeSymbol & "'"
            .VerticalAlignment = xlCenter
        End With
        With .Offset(offsetDown, 3)
            .Resize(2, 1).Merge
            .MergeArea.value = "s'P'"
            .Characters(Start:=3, Length:=2).Font.Subscript = True
            .VerticalAlignment = xlCenter
        End With
        With .Offset(offsetDown, 4)
            .Resize(1, 2).Merge
            .MergeArea.value = "Дисторсия " & uINCR & imSizeSymbol
        End With
        If lens.afocal Then
            .Offset(offsetDown + 1, 4).value = "Угл. мера"
        Else
            .Offset(offsetDown + 1, 4).value = "Лин. мера"
        End If
        .Offset(offsetDown + 1, 5).value = "%"
        With .Offset(offsetDown, 6)
            .Resize(1, 3).Merge
            If lens.afocal Then
                .MergeArea.value = "Астигматические отрезки, дптр"
            Else
                .MergeArea.value = "Астигматические отрезки, мм"
            End If
        End With
        With .Offset(offsetDown + 1, 6)
            .value = "l'm cos" & uOMEGA
            .Characters(Start:=3, Length:=1).Font.Subscript = True
        End With
        With .Offset(offsetDown + 1, 7)
            .value = "l's cos" & uOMEGA
            .Characters(Start:=3, Length:=1).Font.Subscript = True
        End With
        With .Offset(offsetDown + 1, 8)
            .value = "(l'm-l's) cos" & uOMEGA
            .Characters(Start:=4, Length:=1).Font.Subscript = True
            .Characters(Start:=8, Length:=1).Font.Subscript = True
        End With
        With .Offset(0, 0)
            .Resize(offsetDown + 2, 1).Merge
            .MergeArea.value = uOMEGA
            .VerticalAlignment = xlCenter
        End With
        With .Offset(0, 1)
            .Resize(offsetDown + 2, 1).Merge
            .MergeArea.value = "sP"
            .Characters(Start:=2, Length:=1).Font.Subscript = True
        End With
    End With
    printChief = lens.selectedFieldCount + 3
End Function
Private Function printSag(startCell As Excel.Range, ByRef lens As CLens) As Integer
'returns printed row count
    Dim uINCR As String
    uINCR = ChrW(8710)
    Dim uSIGMA As String
    uSIGMA = ChrW(963)
    Dim uPSI As String
    uPSI = ChrW(968)
    Dim uDEG As String
    uDEG = ChrW(176)
    Dim uLAMBDA As String
    uLAMBDA = ChrW(955)
    Dim uOMEGA As String
    uOMEGA = ChrW(969)
    
    Dim TAberSymbol As String, SAberSymbol As String
    If lens.afocal Then
    'note: symbols always 3 chars wide
        TAberSymbol = uINCR + uSIGMA + "'"
        SAberSymbol = uINCR + uPSI + "'"
    Else
        TAberSymbol = uINCR + "y'"
        SAberSymbol = uINCR + "x'"
    End If
    
    With startCell
        Dim wave As Integer
        Dim field As Integer
        Dim coord As Integer
        With .Offset(0, 1)
            .Resize(3, 1).Merge
            .MergeArea.value = "M"
            .VerticalAlignment = xlCenter
        End With
        With .Offset(0, 2)
            .Resize(3, 1).Merge
            .MergeArea.value = "M'"
            .VerticalAlignment = xlCenter
        End With
        With .Offset(0, 3)
            .Resize(3, 1).Merge
            .MergeArea.value = "m'"
            .VerticalAlignment = xlCenter
        End With
        With .Offset(0, 4)
            Dim vShift As Integer
            For wave = 1 To lens.selectedWaveCount
                With .Offset(0, (wave - 1) * 2)
                    .Offset(2, 0).value = SAberSymbol
                    .Offset(2, 1).value = TAberSymbol
                    For field = 1 To lens.selectedFieldCount
                        For coord = 1 To lens.SagCoordCount
                            vShift = (field - 1) * lens.SagCoordCount + coord + 2
                            If lens.afocal Then
                                .Offset(vShift, 0) = tools.degMinSec( _
                                    tools.Deg(lens.sagOpVal("ANAX", field, wave, coord)))
                                .Offset(vShift, 1) = tools.degMinSec( _
                                    tools.Deg(lens.sagOpVal("ANAY", field, wave, coord)))
                            Else
                                .Offset(vShift, 0) = _
                                    Round(lens.sagOpVal("TRAX", field, wave, coord), 4)
                                .Offset(vShift, 1) = _
                                    Round(lens.sagOpVal("TRAY", field, wave, coord), 4)
                            End If
                        Next coord
                    Next field
                    With .Offset(1, 0)
                        .Resize(1, 2).Merge
                        .MergeArea.value = uLAMBDA + waveLetter(lens, wave)
                        .Characters(Start:=2, Length:=2).Font.Subscript = True
                    End With
                End With
            Next wave
            .Resize(1, lens.selectedWaveCount * 2).Merge
            .MergeArea = "Поперечные аберрации"
        End With
        With .Offset(2, 0)
        'print values
            For field = 1 To lens.selectedFieldCount
                With .Offset((field - 1) * lens.SagCoordCount + 1, 0)
                    .Resize(lens.SagCoordCount, 1).Merge
                    .MergeArea.value = tools.degMinSec(lens.yField(field))
                    .VerticalAlignment = xlCenter
                End With
                For coord = 1 To lens.SagCoordCount
                    vShift = (field - 1) * lens.SagCoordCount + coord
                    .Offset(vShift, 1).value = Round(lens.MEntrAbsSag(field, coord), 3)
                    .Offset(vShift, 2).value = Round(lens.MExitAbsSag(field, coord), 3)
                    .Offset(vShift, 3).value = Round(lens.mExitTShiftSag(field, coord), 3)
                Next coord
            Next field
        End With
        With .Offset(0, 0)
            .Resize(3, 1).Merge
            .MergeArea.value = uOMEGA
            .VerticalAlignment = xlCenter
        End With
    End With
    printSag = lens.SagCoordCount * lens.selectedFieldCount + 3
End Function
Private Function printTang(startCell As Excel.Range, ByRef lens As CLens) As Integer
'returns printed row count
    Dim uINCR As String
    uINCR = ChrW(8710)
    Dim uSIGMA As String
    uSIGMA = ChrW(963)
    Dim uPSI As String
    uPSI = ChrW(968)
    Dim uDEG As String
    uDEG = ChrW(176)
    Dim uLAMBDA As String
    uLAMBDA = ChrW(955)
    Dim uOMEGA As String
    uOMEGA = ChrW(969)
    
    Dim rowCount As Integer
    rowCount = lens.TangCoordCount + 3
    Dim colCount As Integer
    colCount = 2 + 2 * lens.selectedWaveCount - 2
    If lens.selectedWaveCount = 1 Then colCount = 3

    Dim rowOffset As Integer
    Dim colOffset As Integer
    Dim printedRowCount As Integer
    
    Dim field As Integer
    For field = 1 To lens.selectedFieldCount
        Select Case lens.selectedFieldCount
        Case 1 To 2:
        'tables are printed in a row
            rowOffset = 0
            colOffset = (field - 1) * colCount
            printedRowCount = rowCount
        Case 3 To 4:
        'two rows of 2+1 or 2+2
            If field <= 2 Then
                rowOffset = 0
                colOffset = (field - 1) * colCount
            Else
                rowOffset = rowCount + 2
                colOffset = (field - 3) * colCount
            End If
            printedRowCount = rowCount * 2
        Case 5 To 6:
        'first row of 3, second row of 2 or 3
            If field <= 3 Then
                rowOffset = 0
                colOffset = (field - 1) * colCount
            Else
                rowOffset = rowCount + 2
                colOffset = (field - 4) * colCount
            End If
            printedRowCount = rowCount * 2
        End Select
        With startCell.Offset(rowOffset, colOffset)
            With .Offset(1, 0)
                .Resize(2, 1).Merge
                .MergeArea.value = "m'"
                .VerticalAlignment = xlCenter
            End With
            With .Offset(1, 1)
                .Resize(2, 1).Merge
                .MergeArea.value = "m'-m'гл"
                .Characters(Start:=6, Length:=2).Font.Subscript = True
                .VerticalAlignment = xlCenter
            End With
            Dim coord As Integer
            For coord = 1 To lens.TangCoordCount
                .Offset(2 + coord, 0).value = Round(lens.mAbsTang(field, coord), 3)
                .Offset(2 + coord, 1).value = Round(lens.mRelativeTang(field, coord), 3)
            Next coord
            With .Offset(1, 2)
                .Resize(1, lens.selectedWaveCount).Merge
                .MergeArea.value = uINCR & uSIGMA & "'"
            End With
            Dim wave As Integer
            For wave = 1 To lens.selectedWaveCount
            'transverse for each wave
                With .Offset(2, 1 + wave)
                    .value = uLAMBDA & waveLetter(lens, wave)
                    .Characters(Start:=2, Length:=2).Font.Subscript = True
                    For coord = 1 To lens.TangCoordCount
                        If lens.afocal Then
                            .Offset(coord, 0).value = tools.degMinSec( _
                                tools.Deg(lens.tangOpVal("ANAY", field, wave, coord)))
                        Else
                            .Offset(coord, 0).value = _
                                Round(lens.tangOpVal("TRAY", field, wave, coord), 4)
                        End If
                    Next coord
                End With
            Next wave
            Dim shortWaveTR As Double
            Dim longWaveTR As Double
            For wave = 2 To lens.selectedWaveCount - 1
            'transverse differences for each wave pair
                With .Offset(1, 2 + lens.selectedWaveCount)
                    For coord = 1 To lens.TangCoordCount
                        If lens.afocal Then
                            shortWaveTR = lens.tangOpVal("ANAY", field, wave, coord)
                            longWaveTR = lens.tangOpVal("ANAY", field, wave + 1, coord)
                            .Offset(1 + coord, 0).value = _
                                tools.degMinSec(tools.Deg(shortWaveTR - longWaveTR))
                        Else
                            shortWaveTR = lens.tangOpVal("TRAY", field, wave, coord)
                            longWaveTR = lens.tangOpVal("TRAY", field, wave + 1, coord)
                            .Offset(1 + coord, 0).value = _
                                Round(shortWaveTR - longWaveTR, 4)
                        End If
                    Next coord
                    .Resize(2, 1).Merge
                    .MergeArea.value = uINCR & uSIGMA & "'" & waveLetter(lens, wave) & _
                        "-" & uINCR & uSIGMA & waveLetter(lens, wave + 1)
                    .Characters(Start:=3, Length:=2).Font.Subscript = True
                    .Characters(Start:=9, Length:=2).Font.Subscript = True
                End With
            Next wave
            Dim fieldVal As Double, vigFactor As Double
            If lens.isXfield(lens.selectedField(field)) Then
                fieldVal = lens.xField(field)
                vigFactor = lens.vigCompressionX(field)
            Else
                fieldVal = lens.yField(field)
                vigFactor = lens.vigCompressionY(field)
            End If
            If field > 1 Then
            'first header is filled later so that merging does not disturb formatting
                With .Offset(0, 0)
                    .Resize(1, colCount).Merge
                    .MergeArea.value = uOMEGA & CStr(field) & "=" & _
                        tools.degMinSec(fieldVal) + ", k=" & _
                            CStr(Round(1 - lens.vigCompressionY(field), 2))
                    .Characters(Start:=2, Length:=1).Font.Subscript = True
                End With
            End If
        End With
    Next field
    With startCell
        .Resize(1, colCount).Merge
        If lens.isXfield(lens.selectedField(1)) Then
            fieldVal = lens.xField(1)
        Else
            fieldVal = lens.yField(1)
        End If
        .MergeArea.value = uOMEGA & "1=" & _
            tools.degMinSec(fieldVal) + ", k=" & CStr(Round(1 - lens.vigCompressionY(1), 2))
        .Characters(Start:=2, Length:=1).Font.Subscript = True
    End With
    printTang = printedRowCount
End Function
Private Function printAxial(startCell As Excel.Range, ByRef lens As CLens) As Integer
'returns number of rows printed
    Dim uETA As String
    uETA = ChrW(951)
    Dim uLAMBDA As String
    uLAMBDA = ChrW(955)
    Dim uINCR As String
    uINCR = ChrW(8710)
    Dim uSIGMA As String
    uSIGMA = ChrW(963)
    
    Dim trAbSymbol As String
    Dim unitsSymbol As String
    If lens.afocal Then
        trAbSymbol = uSIGMA
        unitsSymbol = "град"
    Else
        trAbSymbol = "y"
        unitsSymbol = "мм"
    End If
    
    With startCell
        Dim wave As Integer
        Dim coord As Integer
        Select Case lens.selectedWaveCount
            Case 1:
            'monochromatic
                With .Offset(0, 2)
                    .Resize(2, 1).Merge
                    If lens.afocal Then
                        .MergeArea.value = uINCR & "s, дптр"
                    Else
                        .MergeArea.value = uINCR & "s, мм"
                    End If
                    .VerticalAlignment = xlCenter
                End With
                With .Offset(0, 3)
                    .Resize(2, 1).Merge
                    If lens.afocal Then
                        .MergeArea.value = uSIGMA & "', град"
                    Else
                        .MergeArea.value = uSIGMA & "', мм"
                    End If
                    .VerticalAlignment = xlCenter
                End With
                With .Offset(0, 4)
                    .Resize(2, 1).Merge
                    .MergeArea.value = "Неизопланатизм " & uETA & ", %"
                    .VerticalAlignment = xlCenter
                End With
            Case 3 To 5:
                With .Offset(0, 2)
                    .Resize(1, lens.selectedWaveCount).Merge
                    If lens.afocal Then
                        .MergeArea.value = "Продольная аберрация " & uINCR & "s, дптр"
                    Else
                        .MergeArea.value = "Продольная аберрация " & uINCR & "s, мм"
                    End If
                End With
                For wave = 1 To lens.selectedWaveCount
                'longitudinal for each wave
                    With .Offset(1, 1 + wave)
                        .value = uLAMBDA & waveLetter(lens, wave)
                        .Characters(Start:=2, Length:=2).Font.Subscript = True
                        For coord = 1 To lens.axialYCoordCount
                            .Offset(coord, 0).value _
                                = Round(lens.axialYOpVal("LONA", wave, coord), 4)
                        Next coord
                    End With
                Next wave
                For wave = 2 To lens.selectedWaveCount - 1
                'longitudinal differences for each wave pair
                    With .Offset(0, lens.selectedWaveCount + wave)
                        Dim shortWaveLONA As Double
                        Dim longWaveLONA As Double
                        For coord = 1 To lens.axialYCoordCount
                            shortWaveLONA = lens.axialYOpVal("LONA", wave, coord)
                            longWaveLONA = lens.axialYOpVal("LONA", wave + 1, coord)
                            .Offset(coord + 1, 0).value = _
                                Round(shortWaveLONA - longWaveLONA, 4)
                        Next coord
                        .Resize(2, 1).Merge
                        .MergeArea.value = uINCR & "s" & waveLetter(lens, wave) & _
                            "-" & uINCR & "s" & waveLetter(lens, wave + 1)
                        .Characters(Start:=3, Length:=2).Font.Subscript = True
                        .Characters(Start:=8, Length:=2).Font.Subscript = True
                        .VerticalAlignment = xlCenter
                    End With
                Next wave
                With .Offset(0, 2 * lens.selectedWaveCount)
                    .Resize(1, lens.selectedWaveCount).Merge
                    .MergeArea.value = "Поперечная аберрация " & uINCR _
                        & trAbSymbol & "', " & unitsSymbol
                End With
                For wave = 1 To lens.selectedWaveCount
                'transverse for each wave
                    With .Offset(1, 2 * lens.selectedWaveCount - 1 + wave)
                        .value = uLAMBDA & waveLetter(lens, wave)
                        .Characters(Start:=2, Length:=2).Font.Subscript = True
                        For coord = 1 To lens.axialYCoordCount
                            If lens.afocal Then
                                .Offset(coord, 0).value = tools.degMinSec( _
                                    tools.Deg(lens.axialYOpVal("ANAY", wave, coord)))
                            Else
                                .Offset(coord, 0).value = _
                                    Round(lens.axialYOpVal("TRAY", wave, coord), 4)
                            End If
                        Next coord
                    End With
                Next wave
                For wave = 2 To lens.selectedWaveCount - 1
                'transverse differences for each wave
                    With .Offset(0, 3 * lens.selectedWaveCount + wave - 2)
                        Dim shortWaveTR As Double
                        Dim longWaveTR As Double
                        For coord = 1 To lens.axialYCoordCount
                            If lens.afocal Then
                                shortWaveTR = tools.Deg(lens.axialYOpVal("ANAY", wave, coord))
                                longWaveTR = tools.Deg(lens.axialYOpVal("ANAY", wave + 1, coord))
                                .Offset(coord + 1, 0).value = _
                                    tools.degMinSec(shortWaveTR - longWaveTR)
                            Else
                                shortWaveTR = lens.axialYOpVal("TRAY", wave, coord)
                                longWaveTR = lens.axialYOpVal("TRAY", wave + 1, coord)
                                .Offset(coord + 1, 0).value = _
                                    Round(shortWaveTR - longWaveTR, 4)
                            End If
                        Next coord
                        .Resize(2, 1).Merge
                        .MergeArea.value = uINCR & trAbSymbol & waveLetter(lens, wave) _
                            & "-" & uINCR & trAbSymbol & waveLetter(lens, wave + 1)
                        .Characters(Start:=3, Length:=2).Font.Subscript = True
                        .Characters(Start:=8, Length:=2).Font.Subscript = True
                        .VerticalAlignment = xlCenter
                    End With
                Next wave
                With .Offset(0, 4 * lens.selectedWaveCount - 2)
                    For coord = 1 To lens.axialYCoordCount
                        .Offset(1 + coord, 0).value = Round(lens.isoplanaticErrorY(coord), 4)
                    Next coord
                    .Resize(2, 1).Merge
                    .MergeArea.value = "Неизопланатизм " & uETA & ", %"
                    .VerticalAlignment = xlCenter
                End With
        End Select
        For coord = 1 To lens.axialYCoordCount
            With .Offset(coord + 1, 0)
                .value = Round(lens.m_entr_AxialY(coord), 3)
                .Offset(0, 1).value = Round(lens.m_exit_AxialY(coord), 3)
            End With
        Next coord
        With .Offset(0, 0)
            .Resize(2, 1).Merge
            .MergeArea.value = "m, мм"
            .VerticalAlignment = xlCenter
        End With
        With .Offset(0, 1)
            .Resize(2, 1).Merge
            .MergeArea.value = "m', мм"
            .VerticalAlignment = xlCenter
        End With
    End With
    printAxial = 2 + lens.axialYCoordCount
End Function
Public Function checkWaveCount(ByRef lens As CLens) As Boolean
'returns true when wavelength count is allowed
    Select Case lens.waveCount
    Case 1:
        Call rinseStatus
        Call printStatus("Выбрана 1 длина волны (монохроматическая ОС)")
        checkWaveCount = True
    Case 3:
        Call rinseStatus
        Call printStatus("Выбрана 3 длины волны (ахроматическая ОС)")
        checkWaveCount = True
    Case 4:
        Call rinseStatus
        Call printStatus("Выбрана 4 длины волны (апохроматическая ОС)")
        checkWaveCount = True
    Case Else:
        Call rinseStatus
        Call printStatus("Выбрана 1, 3 или 4 длины волны!")
        checkWaveCount = False
    End Select
End Function
Private Function waveLetter(lens As CLens, waveNo As Integer) As String
    waveLetter = optics.SpectralLine(1000 * lens.wavelength(lens.selectedWave(waveNo)))
End Function
