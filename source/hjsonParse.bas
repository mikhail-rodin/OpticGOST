Attribute VB_Name = "hjsonParse"

Option Explicit

Private Function parseOneLevel(ByVal jsonString As String) As Object
'outputs a dictionary
'doesn't parse arrays: parseArray(str) is used for that
    Private position As Long 'no of characters in string
    Static key, value As String
    Static currentChar As String
    
    Static arrayFlag As Boolean
    '1 if we're inside an array
    Static arrayCounter As Long
    
    Static charType As Integer
    Const tUNRECOGNISED As Integer = 0
    Const tCOMMENT As Integer = 1
    Const tKEY As Integer = 2
    Const tVAL As Integer = 3
    Const tARR As Integer = 4
    Const tOBJ As Integer = 5
    Const tPENDING As Integer = 6

    Const arrayBASE as Integer = 1
    
    Dim mainDict As Scripting.Dictionary
    Set mainDict = New Scripting.Dictionary
        
    jsonString = Replace(jsonString, " ", "") 'remove spaces
    'now every char in jsonString means something
    ': = value start
    '# = comment start
    '{ } = object start/end
    '[ ] = array start/end
    'CrLf and Lf = expect new key or object or array member
    'comma = next member
    position = arrayBASE
    key = ""
    value = ""
    charType = tUNRECOGNISED
    arrayFlag = false

    Do While position < Len(jsonString)
        currentChar = Mid(jsonString, position, 1)
        Select Case currentChar
            Case "#"
                charType = tCOMMENT
            Case ":"
                value = "" 'rinse value
                charType = tVAL
            Case "{"
                value = currentChar 'rinse value and input {
                charType = tOBJ
            Case "["
                'if we're parsing an array
                value = currentChar 'rinse value and input [
                charType = tARR
            Case "}" Or "]"
                value = value & currentChar 'add the closing }/] to value
                mainDict.Add key, value
                charType = tPENDING
            Case vbCrLf Or vbCr Or vbLf Or ","
                Select Case charType
                    Case 1: charType = tPENDING 'comment=>pending
                    Case 3 'value=>add member to dict
                        mainDict.Add key, value
                        key = ""
                        value = ""
                    'Case 4 Or 5 'array or object
                        'change nothing
                    'Case 6 'pending
                        'change nothing
                End Select
            Case ""
             'error
            Case Else 'it's a simple string
                Select Case charType
                'chartype = what was the previously input symbol
                    Case 0 'unrecognised
                    Case 1 'comment
                        'it's a comment, do nothing
                    Case 2 'key
                        key = key & currentChar
                    Case 3 To 5 'value/object/array
                        value = value & currentChar
                    Case 6 'pending => it's a new key
                        key = currentChar 'rinse and beging inputting a key
                End Select
        End Select
        position = position + 1
    Loop
    
    'if nothing was found
    If charType = tUNRECOGNISED Then
        Err.Raise vbObjectError + 1100, , "error parsing JSON text file"
    End If

    If arrayFlag Then
        Set parseOneLevel = Nothing
    Else
        Set parseOneLevel = mainDict
    End If
End Function

Private Function parseArray(ByVal jsonString As String) As String()
'outputs a dictionary
    Private position As Long 'no of characters in string
    Static key, value As String
    Static currentChar As String
    
    Static arrayFlag As Boolean
    '1 if we're inside an array
    Static arrayCounter As Long
    
    Static charType As Integer
    Const tUNRECOGNISED As Integer = 0
    Const tCOMMENT As Integer = 1
    Const tKEY As Integer = 2
    Const tVAL As Integer = 3
    Const tARR As Integer = 4
    Const tOBJ As Integer = 5
    Const tPENDING As Integer = 6

    Const arrayBASE as Integer = 1 'VBA strings are 1-based
    
    Dim mainArray() As String
    
    jsonString = Replace(jsonString, " ", "") 'remove spaces
    'now every char in jsonString means something
    ': = value start
    '# = comment start
    '{ } = object start/end
    '[ ] = array start/end
    'CrLf and Lf = expect new key or object or array member
    'comma = next member
    position = arrayBASE 
    value = ""
    'nestLevel = 0
    charType = tUNRECOGNISED
    arrayFlag = false
    Redim mainArray(1)
    arrayCounter = arrayBASE

    Do While position < Len(jsonString)
        currentChar = Mid(jsonString, position, 1)
        Select Case currentChar
            Case "#"
                charType = tCOMMENT
            Case ":"
                value = "" 'rinse value
                charType = tVAL
            Case "{"
                value = currentChar 'rinse value and input {
                charType = tOBJ
            Case "["
                'if we're parsing an array
                If charType = tUNRECOGNISED Then
                    arrayFlag = true
                End If
                value = currentChar 'rinse value and input {
                charType = tARR
            Case "}" Or "]"
                value = value & currentChar 'add the closing }/] to value
                charType = tPENDING
            Case vbCrLf Or vbCr Or vbLf Or ","
                Select Case charType
                    Case 1: charType = tPENDING 'comment=>pending
                    Case 3 'value=>
                        key = ""
                        value = ""
                    Case 4 Or 5 'array or object
                        'change nothing or
                        If arrayFlag Then
                        'add value to the mainArray
                            Redim Preserve mainArray(arrayCounter) 
                            mainArray(arrayCounter) = value
                            arrayCounter  = arrayCounter + 1
                        End If
                    Case 6 'pending
                        'change nothing
                End Select
            'Case ""
             'error
            Case Else 'it's a simple string
                Select Case charType
                'chartype = what was the previously input symbol
                    Case 0 'unrecognised
                    Case 1 'comment
                        'it's a comment, do nothing
                    Case 2 'key
                        key = key & currentChar
                    Case 3 To 5 'value/object/array
                        If arrayFlag Then

                        Else
                            value = value & currentChar
                        End If
                    Case 6 'pending => it's a new key
                        key = currentChar 'rinse and beging inputting a key
                End Select
        End Select
        position = position + 1
    Loop
    
    'if nothing was found
    If charType = tUNRECOGNISED Then
        Err.Raise vbObjectError + 1100, , "error parsing JSON text file"
    End If

    If arrayFlag Then
        parseOneLevel = mainArray
    End If
End Function

Public Function readTextToString(filePath As String) As String
    Static fileID As Integer
    Static buffer As String
    
    fileID = FreeFile()
    
    On Error Resume Next:
    Open filePath For Input As fileID
    If Err <> 0 Then
        MsgBox "���� �� ������ � ���������� " & filePath, vbCritical, "������"
        Exit Function
    End If
    
    readTextToString = ""
    While Not EOF(fileID)
        Line Input #fileID, buffer
        readTextToString = readTextToString & buffer
    Wend
    Close filePath
End Function

Public Function jsonToDict(jsonContents As String) As Scripting.Dictionary
    Dim outputDict as Scripting.Dictionary
    Set outputDict = New Scripting.Dictionary

    Const BASE as Integer = 1

    Dim i As Integer
    Dim wave, surf, field, coord As Integer

    Dim wavelength_count As Integer
    Dim primary_wavelength As Integer
    Dim field_type As Integer
    Dim field_count As Integer
    Dim max_field As Double
    Dim unvignetted_field As Double
    Dim surface_count As Integer
    Dim Py_coord_count As Integer

    Dim wavelengths() As Double
    
    Dim axialUnparsed As String
    Dim aperture_dataUnparsed As String
    Dim fieldsUnparsed As String
    Dim chiefUnparsed As String
    Dim image_sizeUnparsed() As String
    Dim wavelengthsUnparsed As String
    Dim surfacesUnparsed As String

    Dim tempStr As String
    Dim tempArray() As String
    Dim tempDict As Scripting.Dictionary

    Dim fieldDict() As Scripting.Dictionary
    Dim apertureDict As Scripting.Dictionary
    Dim axialAberDict() As Scripting.Dictionary
    Dim aberDict() As Scripting.Dictionary
    Dim chiefAberDict As Scripting.Dictionary
    Dim imSizeDict() As Scripting.Dictionary
    Dim maxFieldDict As Scripting.Dictionary
    Dim unvigFieldDict As Scripting.Dictionary
    Dim surfaceDict() As Scripting.Dictionary
    Dim maxFieldDict() As Scripting.Dictionary
    Dim unvigFieldDict() As Scripting.Dictionary

    outputDict = parseOneLevel(jsonContents)
    
    With outputDict
        wavelength_count=Int(.Item("wavelength_count"))
        primary_wavelength=Int(.Item("primary_wavelength"))
        'field_type=Int(.Item("field_type"))
        'field_count=Int(.Item("field_count"))
        'max_field=.Item("max_field")
        'unvignetted_field=.Item("unvignetted_field")
        surface_count=Int(.Item("surface_count"))
        Py_coord_count=Int(.Item("Py_coord_count"))

        wavelengthsUnparsed=.Item("wavelengths")
        fieldsUnparsed=.Item("fields")
        axialUnparsed=.Item("axial")
        chiefUnparsed=.Item("chief")
        surfacesUnparsed=.Item("surfaces")

        Redim wavelengths(wavelength_count)
        For wave = BASE To wavelength_count 
            wavelengths(wave)=parseArray(wavelengthsUnparsed)(wave)
        Next wave 
        Set .Item("wavelengths")=wavelengths

        Redim fieldDict(field_count)
        For field = BASE To field_count 
            Set fieldDict(field)= _
                parseOneLevel(withoutOuterBrackets(parseArray(fieldsUnparsed)(field)))
        Next field
        Set .Item("fields")=fieldDict

        Redim surfaceDict(surface_count)
        For surf = BASE To surface_count 
            Set surfaceDict(surf)= _ 
                parseOneLevel(withoutOuterBrackets(parseArray(surfacesUnparsed)(surf)))
        Next surf
        Set .Item("surfaces")=surfaceDict

        Redim axialAberDict(Py_coord_count)
        For coord=BASE To Py_coord_count
        'add axial aberrations for each Py
            Set axialAberDict(coord)=_
                parseOneLevel(withoutOuterBrackets(parseArray(axialUnparsed)(coord))) 
            
            'add an array of aberObjects for each wave to "aberrations" key
            Redim aberDict(wavelength_count)
            Redim tempArray(wavelength_count)
            tempArray=parseArray(axialAberDict(coord).Item("aberrations"))
            For wave=BASE To wavelength_count
                aberDict(wave)=_
                    parseOneLevel(withoutOuterBrackets(tempArray(wave)))
            Next wave
            Set axialAberDict(coord).Item("aberrations")=aberDict
        Next i 
        Set .Item("axial")=axialAberDict

        'create a temp dict with two records:
        'unvignetted_field and max_field
        Set tempDict = parseOneLevel(.Item("chief"))
        Set maxFieldDict=(tempDict.Item("max_field"))
        Set unvigFieldDict=(tempDict.Item("unvignetted_field"))
        Redim chiefAberDict(wavelength_count)
        Redim tempArray(wavelength_count)

        'populate a chiefAberObject first with max field image size data
        tempArray=parseArray(maxFieldDict.Item("image_size"))
        For wave=BASE To wavelength_count 
            'get image size value for max field at this wave
            tempArray(wave)=withoutOuterBrackets(tempArray(wave))
            chiefAberDict(wave)=parseOneLevel(tempArray(wave))
        Next wave
        Set maxFieldDict.Item("image_size")=chiefAberDict

        tempArray=parseArray(unvigFieldDict.Item("image_size"))
        For wave=BASE To wavelength_count 
            'get image size value for max field at this wave
            tempArray(wave)=withoutOuterBrackets(tempArray(wave))
            chiefAberDict(wave)=parseOneLevel(tempArray(wave))
        Next wave
        Set unvigFieldDict.Item("image_size")=chiefAberDict

        Set .Item("chief") = tempDict 'add max_field and unvignetted_field (and other) entries
        'add correct data for these keys
        Set .Item("chief").Item("max_field") = maxFieldDict
        Set .Item("chief").Item("unvignetted_field") = unvigFieldDict

    End With

    Set jsonToDict = outputDict
    Set outputDict = Nothing
End Function

Public Sub printInfo(info As String)
    With jsonForm.outputTB
        .text = .text & vbCrLf & info
    End With
End Sub

Private Function withoutOuterBrackets(str as String) As String
    'remove outer {}:
    If Left(str, 1)="{" Then
        str=Right(str, Len(str)-1) 
    End If
    If Right(str, 1)="}" Then
        str=Left(str, Len(str)-1)
    End If
    withoutOuterBrackets=str
End Function