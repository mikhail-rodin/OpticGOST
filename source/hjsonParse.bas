Attribute VB_Name = "hjsonParse"
Option Base 0
Option Explicit

Private Function parseOneLevel(ByVal jsonString As String) As Object
'outputs a dictionary
'doesn't parse arrays: parseArray(str) is used for that
    Dim position As Long 'no of characters in string
    Dim key, value As String
    Dim currentChar As String
    
    Dim arrayFlag As Boolean
    Dim objectFlag As Boolean
    '1 if we're inside an array
    Dim arrayCounter As Long
    
    Dim charType As Integer
    Const tUNRECOGNISED As Integer = 0
    Const tCOMMENT As Integer = 1
    Const tKEY As Integer = 2
    Const tVAL As Integer = 3
    Const tARR As Integer = 4
    Const tOBJ As Integer = 5
    Const tPENDING As Integer = 6 'means waiting for a key

    Const arrayBASE As Integer = 1
    
    Dim nestLevel As Integer
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
        'Err.Raise vbObjectError + 1100, , "not recognized as JSON"
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
    Dim position As Long 'no of characters in string
    Dim currentChar As String
    Dim arrayCounter As Long
    Dim commentFlag As Boolean
    
    Const arrayBASE As Integer = 0
    
    Dim nestLevel As Integer
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
    Dim fileID As Integer
    Dim buffer As String
    
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
'creates a tree of nested dictionary from a json file
    
    'TODO: fix max_field
    
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

    Dim tempStr As String
      
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
        
        Dim allWavesStr() As String
        Dim allWaves() As Double
        allWavesStr = delEmptyLines(parseArray(.Item("wavelengths")))
        ReDim allWaves(wavelength_count - 1)
        For i = 0 To (wavelength_count - 1)
            allWaves(i) = Val(allWavesStr(i))
        Next i
        .Item("wavelengths") = allWaves
        'wavelenght array added
        
        Dim fieldDict As Scripting.Dictionary
        Dim fieldDicts As Collection
        Set fieldDicts = New Collection
        Dim fieldsArr() As String
        fieldsArr = delEmptyLines(parseArray(.Item("fields")))
        Dim fieldStr As Variant
        For Each fieldStr In fieldsArr
            Set fieldDict = New Scripting.Dictionary
            Set fieldDict = _
                parseOneLevel(withoutOuterBrackets(fieldStr))
            With fieldDict
                Dim chiefDict As Scripting.Dictionary
                Set chiefDict = New Scripting.Dictionary
                Set chiefDict = parseOneLevel(withoutOuterBrackets(.Item("chief")))
                With chiefDict
                    .Item("REAX") = delEmptyLines(parseArray(.Item("REAX")))
                    .Item("REAY") = delEmptyLines(parseArray(.Item("REAY")))
                End With
                Set .Item("chief") = chiefDict
                
                Dim tangDicts As Collection
                Set tangDicts = New Collection
                Dim tangArr() As String
                tangArr = delEmptyLines(parseArray(.Item("tangential")))
                Dim tangObjUnparsed As Variant
                For Each tangObjUnparsed In tangArr
                    Dim tangDict As Scripting.Dictionary
                    Set tangDict = New Scripting.Dictionary
                    Set tangDict = parseOneLevel(withoutOuterBrackets(tangObjUnparsed))
                    With tangDict
                        .Item("TRAX") = delEmptyLines(parseArray(.Item("TRAX")))
                        .Item("TRAY") = delEmptyLines(parseArray(.Item("TRAY")))
                        .Item("ANAX") = delEmptyLines(parseArray(.Item("ANAX")))
                        .Item("ANAY") = delEmptyLines(parseArray(.Item("ANAY")))
                    End With
                    tangDicts.Add tangDict
                Next tangObjUnparsed
                Set .Item("tangential") = tangDicts
                
                Dim sagDicts As Collection
                Set sagDicts = New Collection
                Dim sagArr() As String
                sagArr = delEmptyLines(parseArray(.Item("sagittal")))
                Dim sagObjUnparsed As Variant
                For Each sagObjUnparsed In sagArr
                    Dim sagDict As Scripting.Dictionary
                    Set sagDict = New Scripting.Dictionary
                    Set sagDict = parseOneLevel(withoutOuterBrackets(sagObjUnparsed))
                    With sagDict
                        .Item("TRAX") = delEmptyLines(parseArray(.Item("TRAX")))
                        .Item("TRAY") = delEmptyLines(parseArray(.Item("TRAY")))
                        .Item("ANAX") = delEmptyLines(parseArray(.Item("ANAX")))
                        .Item("ANAY") = delEmptyLines(parseArray(.Item("ANAY")))
                    End With
                    sagDicts.Add sagDict
                Next sagObjUnparsed
                Set .Item("sagittal") = sagDicts
            End With
            fieldDicts.Add fieldDict
        Next fieldStr
        Set .Item("fields") = fieldDicts
        'fields dict added
        
        Dim surfaceDict As Scripting.Dictionary
        Dim surfaceDicts As Collection
        Set surfaceDicts = New Collection
        For surf = BASE To surface_count - 1
            Set surfaceDict = New Scripting.Dictionary
            Set surfaceDict = _
                parseOneLevel(withoutOuterBrackets(parseArray(.Item("surfaces"))(surf)))
            surfaceDicts.Add surfaceDict
        Next surf
        Set .Item("surfaces") = surfaceDicts
        'surfaces dict added
        
        Dim apertureDict As Scripting.Dictionary
        Set apertureDict = New Scripting.Dictionary
        Set apertureDict = parseOneLevel(withoutOuterBrackets(.Item("aperture_data")))
        Set .Item("aperture_data") = apertureDict
        'aperture data dict added
        
        Dim axialDicts As Collection
        Set axialDicts = New Collection
        Dim axialArr() As String
        axialArr = delEmptyLines(parseArray(.Item("axial")))
        Dim axialObjUnparsed As Variant
        For Each axialObjUnparsed In axialArr
            Dim axialDict As Scripting.Dictionary
            Set axialDict = New Scripting.Dictionary
            Set axialDict = parseOneLevel(withoutOuterBrackets(axialObjUnparsed))
            With axialDict
                .Item("TRAX") = delEmptyLines(parseArray(.Item("TRAX")))
                .Item("TRAY") = delEmptyLines(parseArray(.Item("TRAY")))
                .Item("LONA") = delEmptyLines(parseArray(.Item("LONA")))
            End With
            axialDicts.Add axialDict
        Next axialObjUnparsed
        Set .Item("axial") = axialDicts
        
        Dim maxAberDict As Scripting.Dictionary
        Set maxAberDict = New Scripting.Dictionary
        'Set .Item("maximum") = parseOneLevel(withoutOuterBrackets(.Item("maximum"))
        
    End With
    Set jsonToDict = outputDict
End Function
Function delEmptyLines(strArr() As String) As String()
    Dim outArr() As String
    Dim i As Integer
    For i = 0 To UBound(strArr)
        If Replace(strArr(i), " ", "") <> "" Then
            ReDim Preserve outArr(i)
            outArr(i) = strArr(i)
        Else
            Exit For
        End If
    Next i
    delEmptyLines = outArr
End Function

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





