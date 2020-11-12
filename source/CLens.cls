VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'CLens is used as a predeclared class, effectively a singleton
'class must not be instantiated

'wave and field indices are 1-based like in Zemax
Private Const ERR_OUT_OF_RANGE As String = vbObjectError + 1024

'Note on data structs:
'Fields() is an unsorted collection in order as they appear in JSON file
'selectedFields() is an INDEX (i.e. integer) array in order as fields are shown in UI
'Waves() is an unsorted array in order as they are set in Zemax
'selectedWaves() is a sorted INDEX array: first is primary, then shortest to longest

Private Type TLens
    mainDict As Scripting.Dictionary
    selectedFields() As Integer
    '1-based
    
    selectedWaves() As Integer
    '1-based
    'sorted: first is primary wave, then short to long
    
    fieldCount As Integer
    waveCount As Integer
    anamorphic As Boolean
    afocal As Boolean
    primaryWaveNo As Integer
    surfaceCount As Integer
End Type
Private This As TLens
Private Sub Class_Initialize()
    ReDim This.selectedFields(0)
    ReDim This.selectedWaves(0)
    
    Set This.mainDict = Nothing
End Sub
Public Sub parse(json As String)
    ReDim This.selectedFields(0)
    ReDim This.selectedWaves(0)
    
    Set This.mainDict = New Scripting.Dictionary
    Set This.mainDict = hjsonParse.jsonToDict(json)
    
    This.waveCount = This.mainDict.Item("wavelength_count")
    This.fieldCount = This.mainDict.Item("fields").Count
    This.afocal = This.mainDict.Item("afocal")
    This.primaryWaveNo = This.mainDict.Item("primary_wavelength")
    This.surfaceCount = This.mainDict.Item("surface_count")
End Sub
Property Get name() As String
    name = This.mainDict.Item("name")
End Property
Property Get selectedWaveCount() As Integer
    selectedWaveCount = UBound(This.selectedWaves)
End Property
Property Get selectedFieldCount()
    selectedFieldCount = UBound(This.selectedFields)
End Property
Property Get afocal() As Boolean
    afocal = This.afocal
End Property
Property Get SagCoordCount() As Integer
    SagCoordCount = This.mainDict.Item("fields")(1).Item("sagittal").Count
End Property
Property Get TangCoordCount() As Integer
    TangCoordCount = This.mainDict.Item("fields")(1).Item("tangential").Count
End Property
Property Get primaryWaveNo() As Integer
    primaryWaveNo = This.primaryWaveNo
End Property
Property Get anamorphic() As Boolean
'TODO
    anamorphic = False
End Property
Property Get field_type() As Integer
    field_type = This.mainDict("field_type")
End Property
Property Get units() As String
    units = This.mainDict.Item("units")
End Property
Property Get fieldCount() As Integer
    fieldCount = This.fieldCount
End Property
Property Get isXfield(fieldNo As Integer) As Boolean
'returns true if y_field = 0 for a given field
    Const eps As Double = 0.05
    If Abs(Me.srcHy(fieldNo)) < eps And Abs(Me.srcHx(fieldNo)) > 2 * eps Then isXfield = True
End Property
Property Get waveCount() As Integer
    waveCount = This.waveCount
End Property
Property Get tangOpVal(opName As String, selField As Integer, selWave As Integer, coord As Integer) As Double
    tangOpVal = This.mainDict.Item("fields")(This.selectedFields(selField)) _
        .Item("tangential")(coord).Item(opName)(This.selectedWaves(selWave) - 1)
End Property
Property Get sagOpVal(opName As String, selField As Integer, selWave As Integer, coord As Integer) As Double
    sagOpVal = This.mainDict.Item("fields")(This.selectedFields(selField)) _
        .Item("sagittal")(coord).Item(opName)(This.selectedWaves(selWave) - 1)
End Property
Property Get axialXOpVal(opName As String, selWave As Integer, coord As Integer) As Double
    axialXOpVal = This.mainDict.Item("axial_x")(coord) _
        .Item(opName)(This.selectedWaves(selWave) - 1)
End Property
Property Get axialYCoordCount() As Integer
    axialYCoordCount = This.mainDict.Item("axial_y").Count
End Property
Property Get axialYOpVal(opName As String, selWave As Integer, coord As Integer) As Double
    axialYOpVal = This.mainDict.Item("axial_y")(coord) _
        .Item(opName)(This.selectedWaves(selWave) - 1)
End Property
Property Get chiefOpVal(opName As String, selWave As Integer, selField As Integer) As Double
    If selWave = 0 Then
    'this operand has only one wavelength
        chiefOpVal = This.mainDict.Item("fields")(This.selectedFields(selField)) _
            .Item("chief").Item(opName)
    Else
        chiefOpVal = This.mainDict.Item("fields")(This.selectedFields(selField)) _
            .Item("chief").Item(opName)(This.selectedWaves(selWave) - 1)
    End If
End Property
Property Get selectedWaves() As Integer()
    selectedWaves = This.selectedWaves
End Property
Property Get selectedWave(i As Integer) As Integer
    selectedWave = This.selectedWaves(i)
End Property
Property Get wavelength(waveNo As Integer) As Double
    If waveNo <= This.waveCount Then
        wavelength = This.mainDict.Item("wavelengths")(waveNo - 1)
    Else
        Err.Raise ERR_OUT_OF_RANGE, "CLens::wavelength", "array index out of range"
    End If
End Property
Property Get selectedFields() As Integer()
    selectedFields = This.selectedFields
End Property
Property Get selectedField(selFieldNo As Integer) As Integer
    selectedField = This.selectedFields(selFieldNo)
End Property
Property Get m_obj(fieldNo As Integer) As Double

End Property
Property Get xField(selField As Integer) As Double
    xField = This.mainDict.Item("fields")(This.selectedFields(selField)).Item("x_field")
End Property
Property Get srcXField(fieldNo As Integer) As Double
    srcXField = This.mainDict.Item("fields")(fieldNo).Item("x_field")
End Property
Property Get yField(selField As Integer) As Double
    yField = This.mainDict.Item("fields")(This.selectedFields(selField)).Item("y_field")
End Property
Property Get srcYField(fieldNo As Integer) As Double
    srcYField = This.mainDict.Item("fields")(fieldNo).Item("y_field")
End Property
Property Get srcHx(fieldNo As Integer) As Double
    srcHx = This.mainDict.Item("fields")(fieldNo).Item("Hx")
End Property
Property Get Hx(selField As Integer) As Double
    Hx = This.mainDict.Item("fields")(This.selectedFields(selField)).Item("Hx")
End Property
Property Get srcHy(fieldNo As Integer) As Double
    srcHy = This.mainDict.Item("fields")(fieldNo).Item("Hy")
End Property
Property Get Hy(selField As Integer) As Double
    Hy = This.mainDict.Item("fields")(This.selectedFields(selField)).Item("Hy")
End Property
Property Get vigAngle(selField As Integer) As Double
    vigAngle = This.mainDict.Item("fields")(This.selectedFields(selField)) _
        .Item("vignetting_angle")
End Property
Property Get vigCompressionX(selField As Integer) As Double
    vigCompressionX = This.mainDict.Item("fields")(This.selectedFields(selField)) _
        .Item("vignetting_compression_x")
End Property
Property Get vigCompressionY(selField As Integer) As Double
    vigCompressionY = This.mainDict.Item("fields")(This.selectedFields(selField)) _
        .Item("vignetting_compression_y")
End Property
Property Get vigDecenterX(selField As Integer) As Double
    vigDecenterX = This.mainDict.Item("fields")(This.selectedFields(selField)) _
        .Item("vignetting_decenter_x")
End Property
Property Get vigDecenterY(selField As Integer) As Double
    vigDecenterY = This.mainDict.Item("fields")(This.selectedFields(selField)) _
        .Item("vignetting_decenter_y")
End Property
Property Get curvature(surfaceNo As Integer) As Double
    curvature = This.mainDict.Item("surfaces")(surfaceNo).Item("curvature")
End Property
Property Get thickness(surfaceNo As Integer) As Double
    thickness = This.mainDict.Item("surfaces")(surfaceNo).Item("thickness")
End Property
Property Get surfaceCount() As Integer
    surfaceCount = This.surfaceCount
End Property
Property Get mRelativeTang(selField As Integer, coord As Integer) As Double
'image space m coord relative to chief ray heigth for a tangential ray fan
    Dim EXPD As Double
    EXPD = This.mainDict.Item("aperture_data").Item("D_im")
    Dim P As Double, P1 As Double
    If Me.isXfield(This.selectedFields(selField)) Then
        P = This.mainDict.Item("fields")(This.selectedFields(selField)) _
            .Item("tangential")(coord).Item("Px")
        P1 = Me.vigDecenterX(selField) + (1 - Me.vigCompressionX(selField)) * P
    Else
        P = This.mainDict.Item("fields")(This.selectedFields(selField)) _
            .Item("tangential")(coord).Item("Py")
        P1 = Me.vigDecenterY(selField) + (1 - Me.vigCompressionY(selField)) * P
    End If
    mRelativeTang = EXPD * P1 / 2
End Property
Property Get mRelativeSag(selField As Integer, coord As Integer) As Double
'image space m coord relative to chief ray heigth for a tangential ray fan
    Dim EXPD As Double
    EXPD = This.mainDict.Item("aperture_data").Item("D_im")
    Dim P As Double, P1 As Double
    If Me.isXfield(This.selectedFields(selField)) Then
        P = This.mainDict.Item("fields")(This.selectedFields(selField)) _
            .Item("tangential")(coord).Item("Px")
        P1 = Me.vigDecenterX(selField) + (1 - Me.vigCompressionX(selField)) * P
    Else
        P = This.mainDict.Item("fields")(This.selectedFields(selField)) _
            .Item("tangential")(coord).Item("Py")
        P1 = Me.vigDecenterY(selField) + (1 - Me.vigCompressionY(selField)) * P
    End If
    mRelativeSag = EXPD * P1 / 2
End Property
Property Get m_entr_AxialY(coord As Integer) As Double
    Dim ENPD As Double
    ENPD = This.mainDict.Item("aperture_data").Item("D_obj")
    Dim P As Double
    P = This.mainDict.Item("axial_y")(coord).Item("Py")
    m_entr_AxialY = ENPD * P / 2
End Property
Property Get m_exit_AxialY(coord As Integer) As Double
    Dim EXPD As Double
    EXPD = This.mainDict.Item("aperture_data").Item("D_im")
    Dim P As Double
    P = This.mainDict.Item("axial_y")(coord).Item("Py")
    m_exit_AxialY = EXPD * P / 2
End Property
Property Get MEntrAbsSag(selField As Integer, coord As Integer) As Double
'image space m coord relative to chief ray heigth for a tangential ray fan
    Dim ENPD As Double
    ENPD = This.mainDict.Item("aperture_data").Item("D_obj")
    Dim P As Double, P1 As Double
    If Me.isXfield(This.selectedFields(selField)) Then
        P = This.mainDict.Item("fields")(This.selectedFields(selField)) _
            .Item("sagittal")(coord).Item("Py")
        P1 = Me.vigDecenterY(selField) + (1 - Me.vigCompressionY(selField)) * P
    Else
        P = This.mainDict.Item("fields")(This.selectedFields(selField)) _
            .Item("sagittal")(coord).Item("Px")
        P1 = Me.vigDecenterX(selField) + (1 - Me.vigCompressionX(selField)) * P
    End If
    MEntrAbsSag = ENPD * P1 / 2
End Property
Property Get mExitTShiftSag(selField As Integer, coord As Integer) As Double
'tangential exit pupil m coord of a sagittal ray fan
    mExitTShiftSag = chiefRayYShiftEXP(selField)
End Property
Property Get MExitAbsSag(selField As Integer, coord As Integer) As Double
'image space m coord relative to chief ray heigth for a tangential ray fan
    Dim EXPD As Double
    EXPD = This.mainDict.Item("aperture_data").Item("D_im")
    Dim P As Double, P1 As Double
    If Me.isXfield(This.selectedFields(selField)) Then
        P = This.mainDict.Item("fields")(This.selectedFields(selField)) _
            .Item("sagittal")(coord).Item("Py")
        P1 = Me.vigDecenterY(selField) + (1 - Me.vigCompressionY(selField)) * P
    Else
        P = This.mainDict.Item("fields")(This.selectedFields(selField)) _
            .Item("sagittal")(coord).Item("Px")
        P1 = Me.vigDecenterX(selField) + (1 - Me.vigCompressionX(selField)) * P
    End If
    MExitAbsSag = EXPD * P1 / 2
End Property
Property Get mAbsTang(selField As Integer, coord As Integer)
    mAbsTang = Me.mRelativeTang(selField, coord) + chiefRayYShiftEXP(selField)
End Property
Property Get mAbsSag(selField As Integer, coord As Integer)
    mAbsTang = Me.mRelativeSag(selField, coord)
End Property
Private Function chiefRayYShiftEXP(selField As Integer) As Double
'Y coord of real chief ray on the exit pupil
    Dim sP_exit_real As Double, sP_exit_paraxial As Double, realAng As Double
    realAng = Me.chiefOpVal("RANG", 1, selField)
    'since Zemax measures EXPP from image surf (+ is to right), we'll substract last d
    sP_exit_paraxial = Me.thickness(This.surfaceCount - 1) + This.mainDict.Item("aperture_data").Item("EXPP")
    sP_exit_real = Me.sP_exit(selField, 1)
    chiefRayYShiftEXP = (sP_exit_real - sP_exit_paraxial) * Tan(realAng)
End Function
Property Get sP_entr(selField As Integer, selWave As Integer) As Double
'entrance pupil pos relative to first surface
    Dim h1 As Double, ang1 As Double
    h1 = This.mainDict.Item("fields")(This.selectedFields(selField)) _
        .Item("chief").Item("h_1")(This.selectedWaves(selWave) - 1)
    ang1 = This.mainDict.Item("fields")(This.selectedFields(selField)) _
        .Item("chief").Item("entr_RANG")(This.selectedWaves(selWave) - 1)
    sP_entr = -h1 / Tan(ang1)
End Property
Property Get sP_exit(selField As Integer, selWave As Integer) As Double
'entrance pupil pos relative to first surface
    Dim h_q As Double, ang_q As Double
    h_q = This.mainDict.Item("fields")(This.selectedFields(selField)) _
        .Item("chief").Item("h_q")(This.selectedWaves(selWave) - 1)
    ang_q = This.mainDict.Item("fields")(This.selectedFields(selField)) _
        .Item("chief").Item("RANG")(This.selectedWaves(selWave) - 1)
    sP_exit = -h_q / Tan(ang_q)
End Property
Public Sub addField(fieldNo As Integer)
    If fieldNo <= This.fieldCount Then
        Dim isRepeated As Boolean
        Dim field As Variant
        For Each field In This.selectedFields
            If field = fieldNo Then isRepeated = True
        Next field
        If Not isRepeated Then
            Dim uBnd As Integer
            uBnd = UBound(This.selectedFields)
            ReDim Preserve This.selectedFields(uBnd + 1)
            This.selectedFields(uBnd + 1) = fieldNo
        End If
    Else
        'error
    End If
End Sub
Public Sub delField(fieldNo As Integer)
    Dim temp() As Integer
    Dim i As Integer
    Dim fld As Variant
    i = -1
    For Each fld In This.selectedFields
        If fld <> fieldNo Then
            i = i + 1
            ReDim Preserve temp(i)
            temp(i) = fld
        End If
    Next fld
    ReDim This.selectedFields(i)
    This.selectedFields = temp
End Sub
Public Sub addWave(waveNo As Integer)
    If waveNo <= This.waveCount Then
        Dim isRepeated As Boolean
        Dim wave As Variant
        For Each wave In This.selectedWaves
            If wave = waveNo Then isRepeated = True
        Next wave
        If Not isRepeated Then
            Dim uBnd As Integer
            uBnd = UBound(This.selectedWaves)
            ReDim Preserve This.selectedWaves(uBnd + 1)
            This.selectedWaves(uBnd + 1) = waveNo
        End If
        Call sortSelWaves
    Else
        'error
    End If
End Sub
Public Sub delWave(waveNo As Integer)
    Dim temp() As Integer
    Dim i As Integer
    Dim wave As Variant
    i = -1
    For Each wave In This.selectedWaves
        If wave <> waveNo Then
            i = i + 1
            ReDim Preserve temp(i)
            temp(i) = wave
        End If
    Next wave
    ReDim This.selectedWaves(i)
    This.selectedWaves = temp
    Call sortSelWaves
End Sub
Private Sub sortSelWaves()
    If UBound(This.selectedWaves) > 1 Then
    'else there's nothing to sort
        Dim temp() As Integer
        Dim iTemp As Integer
        Dim iSel As Integer
        Dim containsPrimary As Boolean
        iTemp = 0
        'now we'll get a copy of selected wave array
        'without the zero at index 0 and the primary wave
        For iSel = 1 To UBound(This.selectedWaves)
            If This.selectedWaves(iSel) = This.primaryWaveNo Then
                containsPrimary = True
            Else
                ReDim Preserve temp(iTemp)
                temp(iTemp) = This.selectedWaves(iSel)
                iTemp = iTemp + 1
            End If
        Next iSel
        'so now we've got an array of secondary waves
        Call tools.QuickSort(temp, 0, UBound(temp)) 'sort ascending
        'TODO: sort by wavelength
        If containsPrimary Then
            This.selectedWaves(1) = This.primaryWaveNo
            For iSel = 2 To UBound(This.selectedWaves)
                This.selectedWaves(iSel) = temp(iSel - 2)
            Next iSel
        Else
            For iSel = 1 To UBound(This.selectedWaves)
                This.selectedWaves(iSel) = temp(iSel - 1)
            Next iSel
        End If
    End If
End Sub
