Attribute VB_Name = "modApplicabilityChecker"

Option Explicit

Public Function CheckApplicability( _
    ByVal resultInput As clsResultInput, _
    ByVal limits As clsNormativeLimit) _
    As clsApplicabilityResult

    Dim rawValue As String
    Dim rawUncertainty As String
    Dim operatorOperand As String

    rawValue = Trim$(resultInput.rawValue)
    rawUncertainty = UCase$(Trim$(resultInput.rawUncertainty))

    ' 1. Validar que existan límites definidos.
    If Not HasDefinedLimits(limits) Then
        Set CheckApplicability = CreateApplicabilityResult( _
            False, _
            ApplicabilityReason_LimitesNoDefinidos, _
            0, _
            0, _
            EvaluationMode_NotApplicable)
        Exit Function
    End If

    ' 2. Identificar resultados con operador.
    If ContainsOperator(rawValue) Then

        ' Solo se admite el caso especial: < X con incertidumbre NE.
        If Left$(rawValue, 1) = "<" And rawUncertainty = "NE" Then

            ' Extraer el operando numérico X.
            operatorOperand = Trim$(mID$(rawValue, 2))

            If Not IsNumericValue(operatorOperand) Then
                Set CheckApplicability = CreateApplicabilityResult( _
                    False, _
                    ApplicabilityReason_ResultadoNoNumerico, _
                    0, _
                    0, _
                    EvaluationMode_NotApplicable)
                Exit Function
            End If

            ' Resultado aplicable mediante la regla especial.
            Set CheckApplicability = CreateApplicabilityResult( _
                True, _
                ApplicabilityReason_None, _
                ToNumericValue(operatorOperand), _
                0, _
                EvaluationMode_OperatorLessThan)

            Exit Function

        End If

        ' Los demás resultados con operador siguen sin ser aplicables.
        Set CheckApplicability = CreateApplicabilityResult( _
            False, _
            ApplicabilityReason_ResultadoConOperador, _
            0, _
            0, _
            EvaluationMode_NotApplicable)

        Exit Function
    End If

    ' 3. Validar el valor numérico convencional.
    If Not IsNumericValue(rawValue) Then
        Set CheckApplicability = CreateApplicabilityResult( _
            False, _
            ApplicabilityReason_ResultadoNoNumerico, _
            0, _
            0, _
            EvaluationMode_NotApplicable)
        Exit Function
    End If

    ' 4. Validar la incertidumbre numérica convencional.
    If Not IsNumericValue(rawUncertainty) Then
        Set CheckApplicability = CreateApplicabilityResult( _
            False, _
            ApplicabilityReason_IncertidumbreNoDisponible, _
            0, _
            0, _
            EvaluationMode_NotApplicable)
        Exit Function
    End If

    ' 5. Resultado numérico convencional: aplicar EU24.
    Set CheckApplicability = CreateApplicabilityResult( _
        True, _
        ApplicabilityReason_None, _
        ToNumericValue(rawValue), _
        ToNumericValue(rawUncertainty), _
        EvaluationMode_EU24)

End Function


Private Function HasDefinedLimits( _
    ByVal limits As clsNormativeLimit) As Boolean

    HasDefinedLimits = _
        limits.HasLIE Or _
        limits.HasLSE

End Function


Private Function ContainsOperator( _
    ByVal value As String) As Boolean

    Dim text As String

    text = Trim$(value)

    ContainsOperator = _
        Left$(text, 1) = "<" Or _
        Left$(text, 1) = ">"

End Function


Private Function IsNumericValue( _
    ByVal value As String) As Boolean

    Dim text As String

    text = Trim$(value)

    If Len(text) = 0 Then
        IsNumericValue = False
        Exit Function
    End If

    IsNumericValue = IsNumeric(text)

End Function


Private Function ToNumericValue( _
    ByVal value As String) As Double

    Dim cleanValue As Variant

    cleanValue = Trim$(value)
    cleanValue = Replace( _
        cleanValue, _
        ".", _
        Application.International(xlDecimalSeparator))

    ToNumericValue = CDbl(cleanValue)

End Function

