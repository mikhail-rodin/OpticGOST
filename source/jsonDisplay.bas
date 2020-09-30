Attribute VB_Name = "jsonDisplay"
Option Base 0
Option Explicit
Public Sub printInfo(ByVal info As String)
    With jsonForm.outputTB
        .text = .text & vbCrLf & info
    End With
End Sub
Public Sub printStatus(ByVal info As String)
    With jsonForm.status
        .Caption = .text & vbCrLf & info
    End With
End Sub
Public Sub rinseInfo()
    With jsonForm.outputTB
        .text = ""
    End With
End Sub
Public Sub rinseStatus()
    With jsonForm.status
        .Caption = ""
    End With
End Sub
Public Sub displayDict(dict As Scripting.Dictionary)
    'prints out dictionary contents in a window
    
    Dim Txt As String
    Dim i As Integer
    
    With dict
        printInfo ("Имя системы: " & dict.Item("name"))
        printInfo ("Число заданных длин волн: " & .Item("wavelength_count"))
        printInfo ("Основная длина волны: " & .Item("primary_wavelength"))
        printInfo ("Тип задания поля: " & .Item("field_type"))
        printInfo ("Число заданных величин поля: " & .Item("field_count"))
        'Txt = Txt & "Полное поле: " & .Item("max_field") & vbCrLf
        'Txt = Txt & "Невиньетированное поле: " & .Item("unvignetted_field") & vbCrLf
        printInfo ("Число поверхностей: " & .Item("surface_count"))
        printInfo ("Число заданных длин волн: " & .Item("wavelength_count"))
        
        printInfo ("Длины волн:")
        For i = 0 To Val(.Item("wavelength_count")) - 1
            printInfo (.Item("wavelengths")(i) & ", ")
        Next i
        printInfo ("")
        printInfo ("Апертурные характеристики:")
        With .Item("aperture_data")
            printInfo ("  Тип апертуры: " & .Item("type"))
            printInfo ("  Величина апертуры: " & .Item("value"))
            printInfo ("  Диаметр входного зрачка: " & .Item("D_im"))
            printInfo ("  Диаметр выходного зрачка: " & .Item("D_obj"))
            printInfo ("  Положение входного зрачка: " & .Item("ENPP") & " от первой поверхности")
            printInfo ("  Положение выходного зрачка: " & .Item("EXPP") & " от плоскости изображения")
        End With
        printInfo ("")
        
        printInfo ("Поверхности:")
        Dim surf As Scripting.Dictionary
        Dim radius, curvature, thickness As Double
        For Each surf In .Item("surfaces")
            curvature = Val(surf.Item("curvature"))
            If curvature = 0 Then
                radius = 0
            Else
                radius = 1 / curvature
            End If
            thickness = Val(surf.Item("thickness"))
            printInfo (surf.Item("no") & "  " & radius & "  " & thickness & surf.Item("glass"))
        Next surf
    End With
    
End Sub

Public Sub dispWaves(ByRef lens As Scripting.Dictionary)
    Dim waveArr() As Double
    waveArr = lens.Item("wavelengths")
    Dim i As Integer
    For i = 0 To UBound(waveArr)
        With jsonForm.waveList
            .AddItem
            .List(i, 0) = CStr(i)
            .List(i, 1) = CStr(1000 * waveArr(i)) + " нм"
            .List(i, 2) = optics.SpectralLine(1000 * waveArr(i))
        End With
    Next i
    
    Dim primaryWave As Integer
    primaryWave = lens.Item("primary_wavelength") - 1 'Zemax is 1-based
    With jsonForm.waveSel
        .Clear
        .AddItem
        .List(0, 0) = CStr(primaryWave)
        .List(0, 1) = CStr(1000 * waveArr(primaryWave)) + " нм"
        .List(0, 2) = optics.SpectralLine(1000 * waveArr(primaryWave))
    End With
End Sub
Public Sub dispFields(ByRef lens As Scripting.Dictionary)
    Dim fieldType As Integer
    Dim lensUnits, postfix As String
    fieldType = lens.Item("field_type")
    lensUnits = lens.Item("units")
    If fieldType = 0 Then
        postfix = ChrW(176) 'degrees dot
    Else
        postfix = " " + lensUnits
    End If
    Dim field As Variant
    Dim i As Integer
    i = 0
    For Each field In lens.Item("fields")
        With jsonForm.fieldList
            .AddItem
            .List(i, 0) = CStr(field.Item("no"))
            .List(i, 1) = CStr(Round(field.Item("Hy"), 2))
            .List(i, 2) = CStr(Round(field.Item("Hx"), 2))
            .List(i, 3) = CStr(Round(field.Item("y_field"), 2)) + postfix
            .List(i, 4) = CStr(Round(field.Item("x_field"), 2)) + postfix
            i = i + 1
        End With
    Next field
End Sub

Public Sub fillAberTable(ByRef lens As Scripting.Dictionary, sheetName As String)
    Dim afocal, anamorphic, OPD As Boolean
    
    Dim primaryWave As Integer
    
    If afocal Then
        
    Else
    
    End If
End Sub

Public Sub copyListboxItem(ByRef srcListbox As MSForms.listBox, ByRef destListbox As MSForms.listBox)
    Dim srcIndex, destIndex, col, i As Integer
    Dim isRepeated As Boolean
    destIndex = destListbox.ListCount - 1
    For srcIndex = 0 To srcListbox.ListCount - 1
        For i = 0 To destListbox.ListCount - 1
            If destListbox.List(i, 0) = srcListbox.List(srcIndex, 0) Then
                isRepeated = True
            End If
        Next i
        If srcListbox.Selected(srcIndex) And Not isRepeated Then
            With destListbox
                .AddItem
                destIndex = destIndex + 1
                For col = 0 To .ColumnCount - 1
                    .List(destIndex, col) = srcListbox.List(srcIndex, col)
                Next col
            End With
        End If
    Next srcIndex
End Sub

Public Sub delListboxItem(ByRef listBox As MSForms.listBox)
    With listBox
        Dim i As Integer
        i = 0
        While i <= listBox.ListCount - 1
            If .Selected(i) Then
                .RemoveItem (i)
            Else
                i = i + 1
            End If
        Wend
    End With
End Sub
