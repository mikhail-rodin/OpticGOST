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

Private Type TLens
    mainDict As Scripting.Dictionary
    selectedFields() As Integer
    selectedWaves() As Integer
    fieldCount As Integer
    waveCount As Integer
    anamorphic As Boolean
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
    This.fieldCount = This.mainDict.Item("field_count")
End Sub
Property Get name()
    name = This.mainDict.Item("name")
End Property
Property Get field_type()
    field_type = This.mainDict("field_type")
End Property
Property Get units()
    units = This.mainDict.Item("units")
End Property
Property Get fieldCount()
    fieldCount = This.fieldCount
End Property
Property Get waveCount()
    waveCount = This.waveCount
End Property
Property Get axialOpVal(opName As String, Px As Double, Py As Double)
'returns Zemax operand value
    Dim axObj As Variant
    For Each axObj In This.mainDict.Item("axial")
        If axObj.Item("Px") = Px And axObj.Item("Py") = Py Then
            axialOpVal = axObj.Item(opName)
        End If
    Next axObj
End Property
Property Get axialDicts() As Collection
'returns axialObjs for Px/Py values relevant to documentation
    Dim axObj As Variant
    Dim res As Collection
    Set res = New Collection
    If anamorphic Then
        For Each axObj In This.mainDict.Item("axial")
            If axObj.Item("Px") = 0 Then
                res.Add axObj
            End If
        Next axObj
    Else
        res = This.mainDict.Item("axial")
    End If
    axialDicts = res
End Property
Property Get tangDicts(fieldNo As Integer) As Collection
'return a collection of tangential aberration dicts
    If fieldNo <= This.fieldCount Then
        tangDicts = This.mainDict.Item("fields")(fieldNo).Item("tangential")
    Else
        Err.Raise ERR_OUT_OF_RANGE, "CLens::tangDicts", "array index out of range"
    End If
End Property
Property Get sagDicts(fieldNo As Integer) As Collection
'returns a collection of sagittal aberration dicts
    If fieldNo <= This.fieldCount Then
        sagDicts = This.mainDict.Item("fields")(fieldNo).Item("sagittal")
    Else
        Err.Raise ERR_OUT_OF_RANGE, "CLens::sagDicts", "array index out of range"
    End If
End Property
Property Get selectedWaves() As Integer()
    selectedWaves = This.selectedWaves
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
Property Get m_obj(fieldNo As Integer) As Double

End Property
Property Get xField(fieldNo As Integer) As Double
    If fieldNo <= This.fieldCount Then
        xField = This.mainDict.Item("fields")(fieldNo).Item("x_field")
    Else
        Err.Raise ERR_OUT_OF_RANGE, "CLens::xField", "array index out of range"
    End If
End Property
Property Get yField(fieldNo As Integer) As Double
    If fieldNo <= This.fieldCount Then
        yField = This.mainDict.Item("fields")(fieldNo).Item("y_field")
    Else
        Err.Raise ERR_OUT_OF_RANGE, "CLens::yField", "array index out of range"
    End If
End Property
Property Get Hx(fieldNo As Integer) As Double
    If fieldNo <= This.fieldCount Then
        Hx = This.mainDict.Item("fields")(fieldNo).Item("Hx")
    Else
        Err.Raise ERR_OUT_OF_RANGE, "CLens::Hx", "array index out of range"
    End If
End Property
Property Get Hy(fieldNo As Integer) As Double
    If fieldNo <= This.fieldCount Then
        Hy = This.mainDict.Item("fields")(fieldNo).Item("Hy")
    Else
        Err.Raise ERR_OUT_OF_RANGE, "CLens::Hy", "array index out of range"
    End If
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
    Else
        'error
    End If
End Sub
Public Sub delWave(waveNo As Integer)
    Dim temp() As Integer
    Dim i As Integer
    Dim wave As Variant
    i = -1
    For Each wave In This.selectedFields
        If wave <> waveNo Then
            i = i + 1
            ReDim Preserve temp(i)
            temp(i) = wave
        End If
    Next wave
    ReDim This.selectedWaves(i)
    This.selectedWaves = temp
End Sub
