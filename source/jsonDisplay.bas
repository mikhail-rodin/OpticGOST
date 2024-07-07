Attribute VB_Name = "jsonDisplay"
Option Base 0
Option Explicit
Public Type TTableSel
    aberrations As Boolean
    rnd As Boolean
    parts As Boolean
End Type
Public Type TOptions
    OPD As Boolean
    anamorphic As Boolean
    mRelative As Boolean
    tgSigma As Boolean
    origFieldindices As Boolean
End Type
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
Public Sub fillTables(ByRef lens As CLens, startCell As Excel.Range, options As TOptions, sel As TTableSel)
    Dim h As Integer, temp As Integer
    Application.ScreenUpdating = False
    
    If sel.rnd Then
        h = printRND(startCell, lens)
        If sel.parts Then
            temp = printParts(startCell.Offset(0, 10), lens)
        End If
    ElseIf sel.parts Then
        h = printParts(startCell, lens)
    Else
        h = 0
    End If
    
    If sel.aberrations Then
        h = h + fillAberTable(lens, startCell.Offset(h + 1, 0), options)
    End If
    Application.ScreenUpdating = True
End Sub
Public Function fillAberTable(ByRef lens As CLens, startCell As Excel.Range, options As TOptions) As Integer
    Dim chiefOffset As Integer, tangOffset As Integer, sagOffset, fullHeight As Integer
    chiefOffset = printAxial(startCell, lens)
    tangOffset = chiefOffset + 1 + printChief(startCell.Offset(chiefOffset, 0), lens)
    sagOffset = tangOffset + 2 + printTang(startCell.Offset(tangOffset, 0), lens, options)
    fullHeight = sagOffset + 1 + printSag(startCell.Offset(sagOffset, 0), lens)
    
    With startCell.Resize(fullHeight * 2, 100)
        .HorizontalAlignment = xlCenter
    End With
    
    fillAberTable = fullHeight
End Function
Public Function fillPrescription(ByRef lens As CLens, startCell As Excel.Range, options As TOptions) As Integer
    Dim height As Integer, temp As Integer
    height = printRND(startCell, lens)
    temp = printParts(startCell.Offset(0, 10), lens)
    fillPrescription = height
End Function
Private Function printChief(startCell As Excel.Range, ByRef lens As CLens) As Integer
'returns printed row count
    Dim uOMEGA As String
    uOMEGA = ChrW(969)
    Dim uLAMBDA As String
    uLAMBDA = ChrW(955)
    Dim uINCR As String
    uINCR = ChrW(8710)
    
    Dim Xfield As Boolean
    Xfield = lens.isSelXfield(1)
    
    Dim imSizeSymbol As String
    Dim imHeightOpcode As String
    
    If lens.afocal Then
        imSizeSymbol = uOMEGA
    Else
        If Xfield Then
            imSizeSymbol = "x"
            imHeightOpcode = "REAX"
        Else
            imSizeSymbol = "y"
            imHeightOpcode = "REAY"
        End If
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
                                tools.deg(lens.chiefRayAngle(wave, field)))
                        Else
                            .Offset(field, 0).value = _
                                Round(lens.chiefOpVal(imHeightOpcode, wave, field), 4)
                        End If
                    Next field
                End With
            Next wave
            
            If lens.selectedWaveCount = 2 Then
                With .Offset(0, 9)
                'print lateral chroma
                    .Resize(3, 1).Merge
                    .MergeArea.value = imSizeSymbol & "'" & waveLetter(lens, 1) & _
                        "-" & imSizeSymbol & "'" & waveLetter(lens, 2)
                    .Characters(Start:=3, Length:=2).Font.Subscript = True
                    .Characters(Start:=8, Length:=2).Font.Subscript = True
                    .VerticalAlignment = xlCenter
                    For field = 1 To lens.selectedFieldCount
                        If lens.afocal Then
                            .Offset(field, 0).value = tools.degMinSec( _
                                tools.deg(lens.chiefRayAngle(1, field) _
                                    - lens.chiefRayAngle(2, field)))
                        Else
                            .Offset(field, 0).value = Round(lens.chiefOpVal(imHeightOpcode, 1, field) _
                                - lens.chiefOpVal(imHeightOpcode, 2, field), 4)
                        End If
                    Next field
                End With
            Else
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
                                    tools.deg(lens.chiefRayAngle(wave, field) _
                                        - lens.chiefRayAngle(wave + 1, field)))
                            Else
                                .Offset(field, 0).value = Round(lens.chiefOpVal(imHeightOpcode, wave, field) _
                                    - lens.chiefOpVal(imHeightOpcode, wave + 1, field), 4)
                            End If
                        Next field
                    End With
                Next wave
            End If
        End If
        With .Offset(offsetDown + 1, 0)
            .Resize(1, 8).NumberFormat = "0.000"
            Dim distortionVal As Double
            Dim zM As Double
            Dim zS As Double
            Dim imSize As Double
            Dim fieldVal As Double
            For field = 1 To lens.selectedFieldCount
                If lens.isSelXfield(field) Then
                    fieldVal = lens.Xfield(field)
                Else
                    fieldVal = lens.yField(field)
                End If
                distortionVal = lens.chiefOpVal("DISG", 0, field)
                .Offset(field, 0).value = tools.degMinSec(fieldVal)
                .Offset(field, 1).value = Round(lens.sP_entr(field, 1), 4)
                .Offset(field, 3).value = Round(lens.sP_exit(field, 1), 4)
                .Offset(field, 5).value = Round(distortionVal, 2)
                zM = lens.chiefOpVal("FCGT", 0, field)
                zS = lens.chiefOpVal("FCGS", 0, field)
                If lens.afocal Then
                    imSize = lens.chiefRayAngle(1, field) 'in radians
                    .Offset(field, 2).value = tools.degMinSec(tools.deg(imSize))
                    .Offset(field, 4).value = tools.degMinSec( _
                        distortionVal * tools.deg(imSize) / 100)
                    .Offset(field, 6).value = Round(zM * Cos(imSize), 4)
                    .Offset(field, 7).value = Round(zS * Cos(imSize), 4)
                    .Offset(field, 8).value = Round((zM - zS) * Cos(imSize), 4)
                Else
                    imSize = lens.chiefOpVal(imHeightOpcode, 1, field)
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
            .VerticalAlignment = xlCenter
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
    Dim TAberOpCode As String, SAberOpCode As String
    Dim Xfield As Boolean
    
    Xfield = lens.isSelXfield(1)
    ' currently you have to manually select either all X or all Y field points
    
    If lens.afocal Then
        'note: symbols always 3 chars wide
        TAberSymbol = uINCR + uSIGMA + "'"
        SAberSymbol = uINCR + uPSI + "'"
        If Xfield Then
            TAberOpCode = "ANAX"
            SAberOpCode = "ANAY"
        Else
            TAberOpCode = "ANAY"
            SAberOpCode = "ANAX"
        End If
    Else
        If Xfield Then
            TAberSymbol = uINCR + "x'"
            SAberSymbol = uINCR + "y'"
            TAberOpCode = "TRAX"
            SAberOpCode = "TRAY"
        Else
            TAberSymbol = uINCR + "y'"
            SAberSymbol = uINCR + "x'"
            TAberOpCode = "TRAY"
            SAberOpCode = "TRAX"
        End If
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
                                    tools.deg(lens.sagOpVal(SAberOpCode, field, wave, coord)))
                                .Offset(vShift, 1) = tools.degMinSec( _
                                    tools.deg(lens.sagOpVal(TAberOpCode, field, wave, coord)))
                            Else
                                .Offset(vShift, 0) = _
                                    Round(lens.sagOpVal(SAberOpCode, field, wave, coord), 4)
                                .Offset(vShift, 1) = _
                                    Round(lens.sagOpVal(TAberOpCode, field, wave, coord), 4)
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
            Dim fieldVal As Double
            For field = 1 To lens.selectedFieldCount
                If Xfield Then
                    fieldVal = lens.Xfield(field)
                Else
                    fieldVal = lens.yField(field)
                End If
                With .Offset((field - 1) * lens.SagCoordCount + 1, 0)
                    .Resize(lens.SagCoordCount, 1).Merge
                    .MergeArea.value = tools.degMinSec(fieldVal)
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
Private Function printTang(startCell As Excel.Range, ByRef lens As CLens, options As TOptions) As Integer
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
    Dim rayCoordColCount As Integer
    rayCoordColCount = 1 + tools.BoolToInt(options.mRelative) + 2 * tools.BoolToInt(options.tgSigma)
    
    Dim colCount As Integer
    
    If lens.selectedWaveCount = 2 Then
        colCount = rayCoordColCount + 3
    Else
        colCount = rayCoordColCount + 2 * lens.selectedWaveCount - 2
    End If
    
    If lens.selectedWaveCount = 1 Then colCount = 3

    Dim rowOffset As Integer 'of the whole table for a field
    Dim colOffset As Integer 'of the whole table
    'recalculated for each field (with a different formula based on field count)
    Dim printedRowCount As Integer 'depends not only on pupil coord count
    'but also on whether the tables are stacked or printed side by side
    
    Dim field As Integer
    For field = 1 To lens.selectedFieldCount
    
        Dim indField As Integer
        If options.origFieldindices Then
            indField = lens.selectedField(field)
        Else
            indField = field
        End If
    
        Dim imSizeSymb As String
        Dim transverseAbrOpCode As String 'depends on whether the plane is XZ or YZ and if the lens is afocal
        Dim fieldVal As Double, vigFactor As Double
        If lens.isSelXfield(field) Then
            fieldVal = lens.Xfield(field)
            vigFactor = lens.vigCompressionX(field)
            If lens.afocal Then
                transverseAbrOpCode = "ANAX"
                imSizeSymb = uSIGMA
            Else
                transverseAbrOpCode = "TRAX"
                imSizeSymb = "x"
            End If
        Else
            fieldVal = lens.yField(field)
            vigFactor = lens.vigCompressionY(field)
            If lens.afocal Then
                transverseAbrOpCode = "ANAY"
                imSizeSymb = uSIGMA
            Else
                transverseAbrOpCode = "TRAY"
                imSizeSymb = "y"
            End If
        End If
        
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
            Dim currentColOffset As Integer
            currentColOffset = 0
            With .Offset(1, currentColOffset)
                .Resize(2, 1).Merge
                .MergeArea.value = "m'"
                .VerticalAlignment = xlCenter
            End With
            Dim coord As Integer
            For coord = 1 To lens.TangCoordCount
                .Offset(2 + coord, currentColOffset).value = Round(lens.mAbsTang(field, coord), 3)
            Next coord
            If options.mRelative Then
                currentColOffset = currentColOffset + 1
                With .Offset(1, currentColOffset)
                    .Resize(2, 1).Merge
                    .MergeArea.value = "m'-m'гл"
                    .Characters(Start:=6, Length:=2).Font.Subscript = True
                    .VerticalAlignment = xlCenter
                End With
                For coord = 1 To lens.TangCoordCount
                    .Offset(2 + coord, currentColOffset).value = Round(lens.mRelativeTang(field, coord), 3)
                Next coord
            End If
            If options.tgSigma Then
                currentColOffset = currentColOffset + 1
                With .Offset(1, currentColOffset)
                    .Resize(2, 1).Merge
                    .MergeArea.value = "tg" + uSIGMA + "'"
                    .VerticalAlignment = xlCenter
                End With
                For coord = 1 To lens.TangCoordCount
                    .Offset(2 + coord, currentColOffset).value = Round(lens.tgSigmaImTang(field, 1, coord), 3)
                Next coord
                currentColOffset = currentColOffset + 1
                With .Offset(1, currentColOffset)
                    .Resize(2, 1).Merge
                    .MergeArea.value = uINCR + "tg" + uSIGMA + "'"
                    .VerticalAlignment = xlCenter
                End With
                Dim tgSigmaChief As Double
                tgSigmaChief = lens.tgChiefRayAngle(field, 1)
                For coord = 1 To lens.TangCoordCount
                    .Offset(2 + coord, currentColOffset).value = _
                        Round(lens.tgSigmaImTang(field, 1, coord) - tgSigmaChief, 3)
                Next coord
            End If
            currentColOffset = currentColOffset + 1
            With .Offset(1, currentColOffset)
                .Resize(1, lens.selectedWaveCount).Merge
                .MergeArea.value = uINCR & imSizeSymb & "'"
            End With
            currentColOffset = currentColOffset - 1 'a step back to fill the same cols
            Dim wave As Integer
            For wave = 1 To lens.selectedWaveCount
                currentColOffset = currentColOffset + 1
                'transverse for each wave
                With .Offset(2, currentColOffset)
                    .value = uLAMBDA & waveLetter(lens, wave)
                    .Characters(Start:=2, Length:=2).Font.Subscript = True
                    For coord = 1 To lens.TangCoordCount
                        If lens.afocal Then
                            .Offset(coord, 0).value = tools.degMinSec( _
                                tools.deg(lens.tangOpVal(transverseAbrOpCode, field, wave, coord)))
                        Else
                            .Offset(coord, 0).value = _
                                Round(lens.tangOpVal(transverseAbrOpCode, field, wave, coord), 4)
                        End If
                    Next coord
                End With
            Next wave
            Dim shortWaveTR As Double
            Dim longWaveTR As Double
            currentColOffset = currentColOffset + 1
            
            Dim nonRefWave As Integer
            If lens.selectedWaveCount = 2 Then
                nonRefWave = 1
            Else
            ' first wave is reference - we don't subtract it
                nonRefWave = 2
            End If
            
            For wave = nonRefWave To lens.selectedWaveCount - 1
            'transverse differences for each wave pair
                'With .Offset(1, 2 + lens.selectedWaveCount)
                With .Offset(1, currentColOffset)
                    For coord = 1 To lens.TangCoordCount
                        If lens.afocal Then
                            shortWaveTR = lens.tangOpVal(transverseAbrOpCode, field, wave, coord)
                            longWaveTR = lens.tangOpVal(transverseAbrOpCode, field, wave + 1, coord)
                            .Offset(1 + coord, 0).value = _
                                tools.degMinSec(tools.deg(shortWaveTR - longWaveTR))
                        Else
                            shortWaveTR = lens.tangOpVal(transverseAbrOpCode, field, wave, coord)
                            longWaveTR = lens.tangOpVal(transverseAbrOpCode, field, wave + 1, coord)
                            .Offset(1 + coord, 0).value = _
                                Round(shortWaveTR - longWaveTR, 4)
                        End If
                    Next coord
                    .Resize(2, 1).Merge
                    .MergeArea.value = uINCR & imSizeSymb & "'" & waveLetter(lens, wave) & _
                        "-" & uINCR & imSizeSymb & "'" & waveLetter(lens, wave + 1)
                    .Characters(Start:=4, Length:=2).Font.Subscript = True
                    .Characters(Start:=10, Length:=2).Font.Subscript = True
                End With
            Next wave

            If field > 1 Then
            'first header is filled later so that merging does not disturb formatting
                With .Offset(0, 0)
                    .Resize(1, colCount).Merge
                    .MergeArea.value = uOMEGA & CStr(indField) & "=" & _
                        tools.degMin(fieldVal) + ", k=" & _
                            CStr(Round(1 - lens.vigCompressionTang(field), 2))
                    .Characters(Start:=2, Length:=1).Font.Subscript = True
                End With
            End If
        End With
    Next field
    With startCell
        .Resize(1, colCount).Merge
        
        If lens.isSelXfield(1) Then
            fieldVal = lens.Xfield(1)
        Else
            fieldVal = lens.yField(1)
        End If
        
        If options.origFieldindices Then
            indField = lens.selectedField(1)
        Else
            indField = 1
        End If
        
        .MergeArea.value = uOMEGA & CStr(indField) & "=" & _
            tools.degMin(fieldVal) + ", k=" & CStr(Round(1 - lens.vigCompressionTang(1), 2))
        .Characters(Start:=2, Length:=1).Font.Subscript = True
    End With
    printTang = printedRowCount + 1
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
    
    Dim sectionXZ As Boolean
    sectionXZ = (lens.isSelXfield(lens.selectedFieldCount) _
        And Not lens.isSelYfield(lens.selectedFieldCount))
        
    Dim coordCount As Integer
    If sectionXZ Then
        coordCount = lens.axialXCoordCount
    Else
        coordCount = lens.axialYCoordCount
    End If
    
    Dim trAbSymbol As String
    Dim unitsSymbol As String
    If lens.afocal Then
        trAbSymbol = uSIGMA
        unitsSymbol = "град"
    Else
        If sectionXZ Then
            trAbSymbol = "x"
        Else
            trAbSymbol = "y"
        End If
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
                    For coord = 1 To coordCount
                        If sectionXZ Then
                            .Offset(coord, 0).value _
                                = Round(lens.axialXOpVal("LONA", 1, coord), 4)
                        Else
                            .Offset(coord, 0).value _
                                = Round(lens.axialYOpVal("LONA", 1, coord), 4)
                        End If
                    Next coord
                End With
                With .Offset(0, 3)
                    .Resize(2, 1).Merge
                    If lens.afocal Then
                        .MergeArea.value = uSIGMA & "', град"
                    Else
                        .MergeArea.value = uINCR & trAbSymbol & "', мм"
                    End If
                    .VerticalAlignment = xlCenter
                    For coord = 1 To coordCount
                        If sectionXZ Then
                            If lens.afocal Then
                                .Offset(coord, 0).value = tools.degMinSec( _
                                    tools.deg(lens.axialXOpVal("ANAX", 1, coord)))
                            Else
                                .Offset(coord, 0).value = _
                                    Round(lens.axialXOpVal("TRAX", 1, coord), 4)
                            End If
                        Else
                            If lens.afocal Then
                                .Offset(coord, 0).value = tools.degMinSec( _
                                    tools.deg(lens.axialYOpVal("ANAY", 1, coord)))
                            Else
                                .Offset(coord, 0).value = _
                                    Round(lens.axialYOpVal("TRAY", 1, coord), 4)
                            End If
                        End If
                    Next coord
                End With
                With .Offset(0, 4)
                    .Resize(2, 1).Merge
                    .MergeArea.value = "Неизопланатизм " & uETA & ", %"
                    .VerticalAlignment = xlCenter
                    For coord = 1 To coordCount
                        .Offset(coord, 0).value = Round(lens.isoplanaticErrorY(coord), 4)
                    Next coord
                End With
            Case 2 To 5:
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
                        For coord = 1 To coordCount
                            If sectionXZ Then
                                .Offset(coord, 0).value _
                                    = Round(lens.axialXOpVal("LONA", wave, coord), 4)
                            Else
                                .Offset(coord, 0).value _
                                    = Round(lens.axialYOpVal("LONA", wave, coord), 4)
                            End If
                        Next coord
                    End With
                Next wave
                For wave = 2 To lens.selectedWaveCount - 1
                'longitudinal differences for each wave pair
                    With .Offset(0, lens.selectedWaveCount + wave)
                        Dim shortWaveLONA As Double
                        Dim longWaveLONA As Double
                        For coord = 1 To coordCount
                            If sectionXZ Then
                                shortWaveLONA = lens.axialXOpVal("LONA", wave, coord)
                                longWaveLONA = lens.axialXOpVal("LONA", wave + 1, coord)
                            Else
                                shortWaveLONA = lens.axialYOpVal("LONA", wave, coord)
                                longWaveLONA = lens.axialYOpVal("LONA", wave + 1, coord)
                            End If
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
                        For coord = 1 To coordCount
                            If sectionXZ Then
                                If lens.afocal Then
                                    .Offset(coord, 0).value = tools.degMinSec( _
                                        tools.deg(lens.axialXOpVal("ANAX", wave, coord)))
                                Else
                                    .Offset(coord, 0).value = _
                                        Round(lens.axialXOpVal("TRAX", wave, coord), 4)
                                End If
                            Else
                                If lens.afocal Then
                                    .Offset(coord, 0).value = tools.degMinSec( _
                                        tools.deg(lens.axialYOpVal("ANAY", wave, coord)))
                                Else
                                    .Offset(coord, 0).value = _
                                        Round(lens.axialYOpVal("TRAY", wave, coord), 4)
                                End If
                            End If
                        Next coord
                    End With
                Next wave
                For wave = 2 To lens.selectedWaveCount - 1
                'transverse differences for each wave
                    With .Offset(0, 3 * lens.selectedWaveCount + wave - 2)
                        Dim shortWaveTR As Double
                        Dim longWaveTR As Double
                        For coord = 1 To coordCount
                            If lens.afocal Then
                                If sectionXZ Then
                                    shortWaveTR = tools.deg(lens.axialXOpVal("ANAX", wave, coord))
                                    longWaveTR = tools.deg(lens.axialXOpVal("ANAX", wave + 1, coord))
                                Else
                                    shortWaveTR = tools.deg(lens.axialYOpVal("ANAY", wave, coord))
                                    longWaveTR = tools.deg(lens.axialYOpVal("ANAY", wave + 1, coord))
                                End If
                                .Offset(coord + 1, 0).value = _
                                    tools.degMinSec(shortWaveTR - longWaveTR)
                            Else
                                If sectionXZ Then
                                    shortWaveTR = lens.axialXOpVal("TRAX", wave, coord)
                                    longWaveTR = lens.axialXOpVal("TRAX", wave + 1, coord)
                                Else
                                    shortWaveTR = lens.axialYOpVal("TRAY", wave, coord)
                                    longWaveTR = lens.axialYOpVal("TRAY", wave + 1, coord)
                                End If
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
                    For coord = 1 To coordCount
                        .Offset(1 + coord, 0).value = Round(lens.isoplanaticErrorY(coord), 4)
                    Next coord
                    .Resize(2, 1).Merge
                    .MergeArea.value = "Неизопланатизм " & uETA & ", %"
                    .VerticalAlignment = xlCenter
                End With
        End Select
        For coord = 1 To coordCount
            With .Offset(coord + 1, 0)
                If sectionXZ Then
                    .value = Round(lens.m_entr_AxialX(coord), 3)
                    .Offset(0, 1).value = Round(lens.m_exit_AxialX(coord), 3)
                Else
                    .value = Round(lens.m_entr_AxialY(coord), 3)
                    .Offset(0, 1).value = Round(lens.m_exit_AxialY(coord), 3)
                End If
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
    printAxial = 2 + coordCount
End Function
Private Function printRND(startCell As Excel.Range, ByRef lens As CLens) As Integer
    Dim srf As Integer, line As Integer
    With startCell
        .Offset(0, 2) = "ne"
        .Offset(0, 2).Characters(2, 1).Font.Subscript = True
        
        .Offset(0, 3) = "ve" 'ChrW(55349) & "e"
        .Offset(0, 3).Characters(2, 1).Font.Subscript = True
        '.Offset(0, 3).Characters(1, 1).Font.Italic = True
        
        .Offset(0, 4) = "Марка стекла"
        .Offset(0, 5) = ChrW(216) & " св."
        .Offset(0, 6) = "стрелка по " & ChrW(216) & "  св."
        
        srf = 1
        For line = 1 To (lens.surfaceCount - 2) * 2 Step 2
            If line = 1 And lens.thickness(srf) = 0 Then
            ' obj at infty => do not include
               .Offset(line, 1).value = ""
            Else
                .Offset(line, 1).value = "d" _
                    & srf - 1 & " = " & Round(lens.thickness(srf), 2)
                .Offset(line, 1).Characters(2, 2).Font.Subscript = True
            End If
    
            .Offset(line, 2).value = "n" & srf - 1 _
                & " = " & Round(lens.indexOfRefraction(srf), 4)
            .Offset(line, 2).Characters(2, 2).Font.Subscript = True
    
            If lens.abbeNumber(srf) = 0 Then
                .Offset(line, 3).value = ""
            Else
                .Offset(line, 3).value = Round(lens.abbeNumber(srf), 2)
                .Offset(line, 3).NumberFormat = "0.00"
            End If
    
            .Offset(line, 4).value = optics.LZOStranslate(lens.glass(srf))
            srf = srf + 1
        Next line
    
        srf = 1
        For line = 0 To (lens.surfaceCount - 2) * 2 + 1 Step 2
            If srf >= 2 Then
                 .Offset(line, 5).value = Round(lens.diameter(srf), 2)
                 .Offset(line, 5).NumberFormat = "0.00"
                 
                 .Offset(line, 6).value = Round(lens.sag(srf), 2)
                 .Offset(line, 6).NumberFormat = "0.00"
            End If
            
            Dim c As Double, r As Double
            c = lens.curvature(srf)
            If Not srf = 1 Then
                If Abs(c) < 0.000001 Then
                    .Offset(line, 0).value = "r" & srf - 1 & " = " & ChrW(8734)
                Else
                    r = 1 / c
                    .Offset(line, 0).value = "r" & srf - 1 & " = " & Round(r, 2)
                End If
                .Offset(line, 0).Characters(2, 2).Font.Subscript = True
            End If
            srf = srf + 1
        Next line
        
        With .Resize(lens.surfaceCount * 2 + 5, 7)
            .Columns.AutoFit
            .HorizontalAlignment = XlHAlign.xlHAlignCenter
        End With
    End With
    printRND = (lens.surfaceCount - 2) * 2 + 1
End Function
Private Function printParts(startCell As Excel.Range, ByRef lens As CLens) As Integer
    With startCell
        .Offset(0, 0) = "№ поз. дет."
        .Offset(0, 1) = ChrW(216) & " св."
        .Offset(0, 2) = "стрелка по " & ChrW(216) & "  св."
        .Offset(0, 3) = ChrW(216) & " св."
        .Offset(0, 4) = "стрелка по " & ChrW(216) & "  св."
        .Offset(0, 5) = "толщина по оси"
    
        Dim elt As Integer, srf As Integer
        elt = 0
        For srf = 1 To lens.surfaceCount - 1
            If Not lens.isAirspace(srf) Then
                elt = elt + 1
                .Offset(elt, 0).value = elt
                .Offset(elt, 1).value = Round(lens.diameter(srf), 2)
                .Offset(elt, 2).value = Round(lens.sag(srf), 2)
                .Offset(elt, 3).value = Round(lens.diameter(srf + 1), 2)
                .Offset(elt, 4).value = Round(lens.sag(srf + 1), 2)
                .Offset(elt, 5).value = Round(lens.thickness(srf), 2)
            End If
        Next srf
        
        .Offset(1, 1).Resize(1 + elt, 5).NumberFormat = "0.00"
        .Offset(1, 0).Resize(1 + elt, 1).NumberFormat = "0"
    End With
    printParts = elt + 1
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
        Call printStatus("Выбрано 3 длины волны (ахроматическая ОС)")
        checkWaveCount = True
    Case 4:
        Call rinseStatus
        Call printStatus("Выбрано 4 длины волны (апохроматическая ОС)")
        checkWaveCount = True
    Case Else:
        Call rinseStatus
        Call printStatus("Выберите 1, 2, 3 или 4 спектральные линии!")
        checkWaveCount = False
    End Select
End Function
Private Function waveLetter(lens As CLens, waveNo As Integer) As String
    waveLetter = optics.SpectralLine(1000 * lens.wavelength(lens.selectedWave(waveNo)))
End Function


