Attribute VB_Name = "hjsonParse"

Option Explicit

Private Function parseOneLevel(ByVal jsonString As String) As Object
'outputs a dictionary
'doesn't parse arrays: parseArray(str) is used for that
    Static position As Long 'no of characters in string
    Static key, value As String
    Static currentChar As String
    
    Static arrayFlag As Boolean
    Static objectFlag As Boolean
    '1 if we're inside an array
    Static arrayCounter As Long
    
    Static charType As Integer
    Const tUNRECOGNISED As Integer = 0
    Const tCOMMENT As Integer = 1
    Const tKEY As Integer = 2
    Const tVAL As Integer = 3
    Const tARR As Integer = 4
    Const tOBJ As Integer = 5
    Const tPENDING As Integer = 6 'means waiting for a key

    Const arrayBASE As Integer = 1
    
    Static nestLevel As Integer
    'only the top-level object is parsed
    'this var is used to prevent the function
    'from exiting on a nested "]" or "}" instead of the top-level "]"/"}"
    'top level is even number of brackets
    
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
    arrayFlag = False
    nestLevel = 0
    Do While position <= Len(jsonString)
        currentChar = Mid(jsonString, position, 1)
        Select Case currentChar
            Case "#"
                If nestLevel = 0 Then
                'if we're not inside a nested obj, treat "#" as usual
                    If charType = tVAL Or charType = tOBJ Or charType = tARR Then
                    'if preceeded on a line by an entry
                    'add entry, otherwise it'll be discarded
                        mainDict.Item(key) = value
                        key = ""
                        value = ""
                    End If
                    charType = tCOMMENT
                Else 'if inside a nested obj
                'just keep adding it all as a value
                    charType = tVAL
                    value = value & currentChar
                End If
            Case ":"
                If nestLevel = 0 Then
                'if we're not inside a nested obj, treat ":" as usual
                    value = "" 'rinse value
                Else
                    value = value & currentChar
                End If
                charType = tVAL 'in both cases
            Case "{"
                nestLevel = nestLevel + 1
                If nestLevel = 0 Then
                'if we're not inside a nested obj, treat "#" as usual
                    value = currentChar 'rinse value and input {
                    charType = tOBJ
                Else 'if inside a nested obj
                'just keep adding it all as a value
                    charType = tVAL
                    value = value & currentChar
                End If
            Case "["
                nestLevel = nestLevel + 1
                If nestLevel = 0 Then
                'if we're not inside a nested obj, treat "#" as usual
                    value = currentChar 'rinse value and input [
                    charType = tARR
                Else 'if inside a nested obj
                'just keep adding it all as a value
                    charType = tVAL
                    value = value & currentChar
                End If
            Case "}", "]"
                nestLevel = nestLevel - 1
                If nestLevel = 0 Then
                    value = value & currentChar 'add the closing }/] to value
                    mainDict.Item(key) = value
                    charType = tPENDING
                Else 'if inside a nested obj
                'just keep adding it all as a value
                    charType = tVAL
                    value = value & currentChar
                End If
            Case vbCrLf, vbCr, vbLf, ","
                If nestLevel = 0 Then
                'if we're not inside a nested obj, treat CrLf as usual
                    Select Case charType
                        Case 1:
                            If Not currentChar = "," Then
                                charType = tPENDING 'comment=>pending on newline
                            End If
                        Case 3 'value=>add member to dict
                            mainDict.Item(key) = value
                            key = ""
                            value = ""
                            charType = tPENDING
                        'Case 4 Or 5 'array or object
                            'change nothing
                        'Case 6 'pending
                            'change nothing
                    End Select
                Else 'if inside a nested obj
                'just keep adding it all as a value
                    charType = tVAL
                    value = value & currentChar
                End If
            Case ""
             'error
            Case Else 'it's a simple string
                Select Case charType
                'chartype = what was the previously input symbol
                    Case 0 'unrecognised
                    'beginning of a file
                        key = currentChar 'rinse and start writing
                        charType = tKEY
                    Case 1 'comment
                        'it's a comment, do nothing
                    Case 2 'key
                        key = key & currentChar
                    Case 3 To 5 'value/object/array
                        value = value & currentChar
                    Case 6 'pending => it's a new key
                        key = currentChar 'rinse and beging inputting a key
                        charType = tKEY
                End Select
        End Select
        position = position + 1
    Loop
    
    'if nothing was found
    If charType = tUNRECOGNISED Then
        Err.Raise vbObjectError + 1100, , "not recognized as JSON"
    End If
    
    If nestLevel <> 0 Then
        Err.Raise vbObjectError + 1101, , "missing } or ] at entry with key " & key
        Set parseOneLevel = Nothing
        Exit Function
    End If

    If arrayFlag Then
        Set parseOneLevel = Nothing
    Else
        Set parseOneLevel = mainDict
    End If
End Function

Private Function parseArray(ByVal jsonString As String) As String()
'outputs a dictionary
    Static position As Long 'no of characters in string
    Static currentChar As String
    Static arrayCounter As Long
    Static commentFlag As Boolean
    
    Const arrayBASE As Integer = 0
    
    Static nestLevel As Integer
    'only the top-level array is parsed
    'this var is used to prevent the function
    'from exiting on a nested "]" instead of the top-level "]"
    'top level is even number of brackets
    
    Dim mainArray() As String
    
    jsonString = Replace(jsonString, " ", "") 'remove spaces
    'now every char in jsonString means something
    ': = value start
    '# = comment start
    '{ } = object start/end
    '[ ] = array start/end
    'CrLf and Lf = expect new key or object or array member
    'comma = next member
    
    If Left(jsonString, 1) = "[" Then
        jsonString = Right(jsonString, Len(jsonString) - 1)
    End If
    currentChar = Left(jsonString, 1)
    If currentChar = vbCr Or currentChar = vbLf Or currentChar = vbCrLf Then
        jsonString = Right(jsonString, Len(jsonString) - 1)
    End If
    If Right(jsonString, 1) = "," Then
        jsonString = Left(jsonString, Len(jsonString) - 1)
    End If
    If Right(jsonString, 1) = "]" Then
        jsonString = Left(jsonString, Len(jsonString) - 1)
    End If
    
    position = 1
    ReDim mainArray(arrayBASE)
    arrayCounter = arrayBASE
    nestLevel = 0
    commentFlag = False
    Do While position <= Len(jsonString)
        currentChar = Mid(jsonString, position, 1)
        'Debug.Assert Not (currentChar = "}")
        Select Case currentChar
            Case "#"
                If nestLevel = 0 Then
                    commentFlag = True
                    'now subsequent chars on the line won't get read into the array
                End If
            Case "{", "["
                nestLevel = nestLevel + 1
                mainArray(arrayCounter) = _
                            mainArray(arrayCounter) & currentChar
            Case "}", "]"
                nestLevel = nestLevel - 1
                mainArray(arrayCounter) = _
                            mainArray(arrayCounter) & currentChar
                                 
            Case ","
            'member added, go to next index
                If nestLevel = 0 Then
                    arrayCounter = arrayCounter + 1
                    ReDim Preserve mainArray(arrayCounter)
                    'next char will be input at a new index
                Else
                    mainArray(arrayCounter) = _
                            mainArray(arrayCounter) & currentChar
                End If
            Case vbCr, vbLf, vbCrLf
                commentFlag = False
                If nestLevel <> 0 Then
                    mainArray(arrayCounter) = _
                            mainArray(arrayCounter) & currentChar
                'preserve newlines in nested objects
                'discard them in simple values
                End If
            Case Else 'it's a simple string
                If commentFlag = False Then
                    mainArray(arrayCounter) = _
                        mainArray(arrayCounter) & currentChar
                End If
        End Select
        position = position + 1
    Loop
    
    If nestLevel <> 0 Then
        Err.Raise vbObjectError + 1101, , "closing ] or } not found at index " & arrayCounter
        Exit Function
    Else
        parseArray = mainArray
    End If
End Function

Public Function readTextToString(ByVal filePath As String) As String
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
        readTextToString = readTextToString & vbLf & buffer
    Wend
    Close filePath
End Function

Public Function jsonToDict(ByVal jsonContents As String) As Scripting.Dictionary
    Dim outputDict As Scripting.Dictionary
    Set outputDict = New Scripting.Dictionary

    Const BASE As Integer = 0

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

    Dim fieldDict As Scripting.Dictionary
    Dim fieldDicts As Collection
    
    Dim apertureDict As Scripting.Dictionary
    
    Dim axialAberDict As Scripting.Dictionary
    Dim axialAberDicts As Collection
    
    Dim aberDict As Scripting.Dictionary
    Dim aberDicts As Collection
    
    Dim chiefAberDict As Scripting.Dictionary
    Dim max_imSizeDicts As Collection
    Dim unvig_imSizeDicts As Collection
    
    Dim imSizeDict As Scripting.Dictionary
    Dim imSizeDicts As Collection
    
    Dim maxFieldDict As Scripting.Dictionary
    
    Dim unvigFieldDict As Scripting.Dictionary
    
    Dim surfaceDict As Scripting.Dictionary
    Dim surfaceDicts As Collection
    
    Set outputDict = parseOneLevel(jsonContents)
    
    With outputDict
        wavelength_count = Int(.Item("wavelength_count"))
        primary_wavelength = Int(.Item("primary_wave_no"))
        field_type = Int(.Item("field_type"))
        field_count = Int(.Item("field_count"))
        max_field = Val(.Item("max_field"))
        unvignetted_field = Val(.Item("unvignetted_field"))
        surface_count = Int(.Item("surface_count"))
        Py_coord_count = Int(.Item("Py_coord_count"))

        wavelengthsUnparsed = .Item("wavelengths")
        fieldsUnparsed = .Item("fields")
        axialUnparsed = .Item("axial")
        chiefUnparsed = .Item("chief")
        surfacesUnparsed = .Item("surfaces")
        
        tempArray = parseArray(wavelengthsUnparsed)
        ReDim wavelengths(wavelength_count - 1)
        For wave = 0 To (wavelength_count - 1)
            wavelengths(wave) = Val(tempArray(wave))
        Next wave
        .Item("wavelengths") = wavelengths
        'wavelenght array added
        
        Set fieldDicts = New Collection
        For field = BASE To field_count - 1
            Set fieldDict = New Scripting.Dictionary
            Set fieldDict = _
                parseOneLevel(withoutOuterBrackets(parseArray(fieldsUnparsed)(field)))
            fieldDicts.Add fieldDict
        Next field
        Set .Item("fields") = fieldDicts
        'fields dict added
        
        Set surfaceDicts = New Collection
        For surf = BASE To surface_count - 1
            Set surfaceDict = New Scripting.Dictionary
            Set surfaceDict = _
                parseOneLevel(withoutOuterBrackets(parseArray(surfacesUnparsed)(surf)))
            surfaceDicts.Add surfaceDict
        Next surf
        Set .Item("surfaces") = surfaceDicts
        'surfaces dict added
        
        Set axialAberDicts = New Collection
        For coord = BASE To Py_coord_count - 1
        'add axial aberrations for each Py
            Set axialAberDict = New Scripting.Dictionary
            Set axialAberDict = _
                parseOneLevel(withoutOuterBrackets(parseArray(axialUnparsed)(coord)))
            
            'add an array of aberObjects for each wave to "aberrations" key
            ReDim tempArray(wavelength_count)
            tempArray = parseArray(axialAberDict(coord).Item("aberrations"))
            For wave = BASE To wavelength_count - 1
                Set aberDict = New Scripting.Dictionary
                Set aberDict = _
                    parseOneLevel(withoutOuterBrackets(tempArray(wave)))
                aberDicts.Add aberDict
            Next wave
            Set axialAberDict.Item("aberrations") = aberDicts
            axialAberDicts.Add axialAberDict
        Next coord
        Set .Item("axial") = axialAberDicts
        'axial aberration dict added

        Set chiefAberDict = New Scripting.Dictionary
        Set chiefAberDict = _
            parseOneLevel(withoutOuterBrackets(.Item("chief")))
        Set .Item("chief") = chiefAberDict
        Set maxFieldDict = New Scripting.Dictionary
        Set maxFieldDict = _
            parseOneLevel(withoutOuterBrackets(chiefAberDict.Item("max_field")))
        Set .Item("max_field") = maxFieldDict
        Set unvigFieldDict = New Scripting.Dictionary
        Set unvigFieldDict = _
            parseOneLevel(withoutOuterBrackets(chiefAberDict.Item("unvignetted_field")))
        Set .Item("unvignetted_field") = unvigFieldDict

        ReDim tempArray(wavelength_count)
        Set max_imSizeDicts = New Collection
        'populate a chiefAberObject first with max field image size data
        tempArray = parseArray(maxFieldDict.Item("image_size"))
        For wave = BASE To wavelength_count - 1
            Set imSizeDict = New Scripting.Dictionary
            'get image size value for max field at this wave
            Set imSizeDict = parseOneLevel(withoutOuterBrackets(tempArray(wave)))
            max_imSizeDicts.Add imSizeDict
        Next wave
        Set maxFieldDict.Item("image_size") = max_imSizeDicts
        
        Set unvig_imSizeDicts = New Collection
        tempArray = parseArray(unvigFieldDict.Item("image_size"))
        For wave = BASE To wavelength_count - 1
            Set imSizeDict = New Scripting.Dictionary
            'get image size value for max field at this wave
            Set imSizeDict = parseOneLevel(withoutOuterBrackets(tempArray(wave)))
            unvig_imSizeDicts.Add imSizeDict
        Next wave
        Set unvigFieldDict.Item("image_size") = unvig_imSizeDicts
        
        Set .Item("chief").Item("max_field") = maxFieldDict
        Set .Item("chief").Item("unvignetted_field") = unvigFieldDict

    End With

    Set jsonToDict = outputDict
    Set outputDict = Nothing
End Function

Public Sub printInfo(ByVal info As String)
    With jsonForm.outputTB
        .text = .text & vbCrLf & info
    End With
End Sub

Public Sub rinseInfo()
    With jsonForm.outputTB
        .text = ""
    End With
End Sub

Private Function withoutOuterBrackets(ByVal str As String) As String
    'remove outer {}:
    Dim firstChar As String
    firstChar = Left(str, 1)
    If firstChar = "{" Then
        str = Right(str, Len(str) - 1)
    End If
    firstChar = Left(str, 1) 'renew
    If firstChar = vbCr Or firstChar = vbLf Or firstChar = vbCrLf Then
        str = Right(str, Len(str) - 1)
    End If
    If Right(str, 1) = "," Then
        str = Left(str, Len(str) - 1)
    End If
    If Right(str, 1) = "}" Then
        str = Left(str, Len(str) - 1)
    End If
    withoutOuterBrackets = str
End Function


