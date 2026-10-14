Attribute VB_Name = "modOperatorEvaluator"

Option Explicit

Public Function EvaluateOperatorLessThan( _
    ByVal operand As Double, _
    ByVal limits As clsNormativeLimit) _
    As clsEvaluationResult

    ' Validar que existan límites definidos.
    If Not limits.HasLIE And Not limits.HasLSE Then

        Set EvaluateOperatorLessThan = CreateEvaluationResult( _
            EvaluationStatus_NotApplicable, _
            EvaluationValue_NoAplica, _
            EvaluationReason_None)

        Exit Function

    End If

    ' Evaluación con ambos límites.
    If limits.HasLIE And limits.HasLSE Then

        Set EvaluateOperatorLessThan = EvaluateBothLimits( _
            operand, _
            limits)

        Exit Function

    End If

    ' Evaluación con límite inferior únicamente.
    If limits.HasLIE Then

        Set EvaluateOperatorLessThan = EvaluateLowerLimit( _
            operand, _
            limits.LIE)

        Exit Function

    End If

    ' Evaluación con límite superior únicamente.
    If limits.HasLSE Then

        Set EvaluateOperatorLessThan = EvaluateUpperLimit( _
            operand, _
            limits.LSE)

        Exit Function

    End If

End Function


Private Function EvaluateUpperLimit( _
    ByVal operand As Double, _
    ByVal upperLimit As Double) _
    As clsEvaluationResult

    If operand <= upperLimit Then

        Set EvaluateUpperLimit = CreateEvaluationResult( _
            EvaluationStatus_Evaluated, _
            EvaluationValue_Cumple, _
            EvaluationReason_None)

    Else

        Set EvaluateUpperLimit = CreateEvaluationResult( _
            EvaluationStatus_NotDeterminable, _
            EvaluationValue_NoDeterminado, _
            EvaluationReason_None)

    End If

End Function


Private Function EvaluateLowerLimit( _
    ByVal operand As Double, _
    ByVal lowerLimit As Double) _
    As clsEvaluationResult

    If operand <= lowerLimit Then

        Set EvaluateLowerLimit = CreateEvaluationResult( _
            EvaluationStatus_Evaluated, _
            EvaluationValue_NoCumple, _
            EvaluationReason_None)

    Else

        Set EvaluateLowerLimit = CreateEvaluationResult( _
            EvaluationStatus_NotDeterminable, _
            EvaluationValue_NoDeterminado, _
            EvaluationReason_None)

    End If

End Function


Private Function EvaluateBothLimits( _
    ByVal operand As Double, _
    ByVal limits As clsNormativeLimit) _
    As clsEvaluationResult

    If operand <= limits.LIE Then

        Set EvaluateBothLimits = CreateEvaluationResult( _
            EvaluationStatus_Evaluated, _
            EvaluationValue_NoCumple, _
            EvaluationReason_None)

    ElseIf operand <= limits.LSE Then

        Set EvaluateBothLimits = CreateEvaluationResult( _
            EvaluationStatus_Evaluated, _
            EvaluationValue_Cumple, _
            EvaluationReason_None)

    Else

        Set EvaluateBothLimits = CreateEvaluationResult( _
            EvaluationStatus_NotDeterminable, _
            EvaluationValue_NoDeterminado, _
            EvaluationReason_None)

    End If

End Function

