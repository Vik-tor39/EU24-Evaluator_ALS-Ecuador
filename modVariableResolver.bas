Attribute VB_Name = "modVariableResolver"
Option Explicit

Public Function ResolveVariableName( _
    ByVal resultVariableName As String, _
    ByVal repository As clsNormativaRepository) _
    As clsVariableResolution

    Dim result As clsVariableResolution

    '1. Coincidencia exacta
    Set result = ResolveExact( _
        resultVariableName, _
        repository)

    If result.status <> VariableResolution_NotFound Then
        Set ResolveVariableName = result
        Exit Function
    End If

    '2. Mapeo explícito
    Set result = ResolveExplicitMapping( _
        resultVariableName, _
        repository)

    If result.status <> VariableResolution_NotFound Then
        Set ResolveVariableName = result
        Exit Function
    End If

    '3. Coincidencia normalizada
    Set result = ResolveNormalized( _
        resultVariableName, _
        repository)

    Set ResolveVariableName = result

End Function

Private Function ResolveExact( _
    ByVal resultVariableName As String, _
    ByVal repository As clsNormativaRepository) _
    As clsVariableResolution

    Dim variables As Collection
    Dim item As Variant

    Dim variableID As String
    Dim variableName As String

    Dim matchCount As Long

    Set variables = repository.GetVariables()

    For Each item In variables

        variableID = CStr(item(0))
        variableName = CStr(item(1))

        If StrComp( _
            Trim$(resultVariableName), _
            Trim$(variableName), _
            vbTextCompare) = 0 Then

            matchCount = matchCount + 1

        End If

    Next item

    If matchCount = 0 Then

        Set ResolveExact = _
            CreateVariableResolution( _
                VariableResolution_NotFound, _
                vbNullString, _
                resultVariableName, _
                vbNullString, _
                VariableMatch_None)

        Exit Function

    End If

    If matchCount > 1 Then

        Set ResolveExact = _
            CreateVariableResolution( _
                VariableResolution_Ambiguous, _
                vbNullString, _
                resultVariableName, _
                vbNullString, _
                VariableMatch_Exact)

        Exit Function

    End If

    For Each item In variables

        variableID = CStr(item(0))
        variableName = CStr(item(1))

        If StrComp( _
            Trim$(resultVariableName), _
            Trim$(variableName), _
            vbTextCompare) = 0 Then

            Set ResolveExact = _
                CreateVariableResolution( _
                    VariableResolution_Resolved, _
                    variableID, _
                    resultVariableName, _
                    variableName, _
                    VariableMatch_Exact)

            Exit Function

        End If

    Next item

End Function

Private Function ResolveNormalized( _
    ByVal resultVariableName As String, _
    ByVal repository As clsNormativaRepository) _
    As clsVariableResolution

    Dim variables As Collection
    Dim item As Variant

    Dim targetName As String
    Dim variableName As String
    Dim variableID As String

    Dim matchCount As Long
    Dim matchedID As String
    Dim matchedName As String

    targetName = NormalizeVariableName(resultVariableName)

    Set variables = repository.GetVariables()

    For Each item In variables

        variableID = CStr(item(0))
        variableName = CStr(item(1))

        If targetName = NormalizeVariableName(variableName) Then

            matchCount = matchCount + 1
            matchedID = variableID
            matchedName = variableName

        End If

    Next item

    If matchCount = 0 Then

        Set ResolveNormalized = _
            CreateVariableResolution( _
                VariableResolution_NotFound, _
                vbNullString, _
                resultVariableName, _
                vbNullString, _
                VariableMatch_None)

        Exit Function

    End If

    If matchCount > 1 Then

        Set ResolveNormalized = _
            CreateVariableResolution( _
                VariableResolution_Ambiguous, _
                vbNullString, _
                resultVariableName, _
                vbNullString, _
                VariableMatch_Normalized)

        Exit Function

    End If

    Set ResolveNormalized = _
        CreateVariableResolution( _
            VariableResolution_Resolved, _
            matchedID, _
            resultVariableName, _
            matchedName, _
            VariableMatch_Normalized)

End Function

Private Function NormalizeVariableName( _
    ByVal value As String) _
    As String

    Dim text As String

    text = Trim$(value)
    text = LCase$(text)

    text = Replace(text, "á", "a")
    text = Replace(text, "é", "e")
    text = Replace(text, "í", "i")
    text = Replace(text, "ó", "o")
    text = Replace(text, "ú", "u")
    text = Replace(text, "ü", "u")
    text = Replace(text, "ñ", "n")

    Do While InStr(text, "  ") > 0
        text = Replace(text, "  ", " ")
    Loop

    NormalizeVariableName = text

End Function

Private Function ResolveExplicitMapping( _
    ByVal resultVariableName As String, _
    ByVal repository As clsNormativaRepository) _
    As clsVariableResolution

    Dim mappings As Collection
    Dim item As Variant

    Dim mappingVariableID As String
    Dim mappingResultName As String
    Dim mappingNormativeName As String

    Dim matchCount As Long
    Dim matchedID As String
    Dim matchedNormativeName As String

    Set mappings = repository.GetVariableMappings()

    For Each item In mappings

        mappingVariableID = CStr(item(0))
        mappingResultName = CStr(item(1))
        mappingNormativeName = CStr(item(2))

        If StrComp( _
            Trim$(resultVariableName), _
            Trim$(mappingResultName), _
            vbTextCompare) = 0 Then

            matchCount = matchCount + 1
            matchedID = mappingVariableID
            matchedNormativeName = mappingNormativeName

        End If

    Next item

    If matchCount = 0 Then

        Set ResolveExplicitMapping = _
            CreateVariableResolution( _
                VariableResolution_NotFound, _
                vbNullString, _
                resultVariableName, _
                vbNullString, _
                VariableMatch_None)

        Exit Function

    End If

    If matchCount > 1 Then

        Set ResolveExplicitMapping = _
            CreateVariableResolution( _
                VariableResolution_Ambiguous, _
                vbNullString, _
                resultVariableName, _
                vbNullString, _
                VariableMatch_ExplicitMapping)

        Exit Function

    End If

    Set ResolveExplicitMapping = _
        CreateVariableResolution( _
            VariableResolution_Resolved, _
            matchedID, _
            resultVariableName, _
            matchedNormativeName, _
            VariableMatch_ExplicitMapping)

End Function
