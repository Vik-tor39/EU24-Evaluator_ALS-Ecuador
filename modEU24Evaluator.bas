Attribute VB_Name = "modEU24Evaluator"
Option Explicit

Public Function EvaluateEU24( _
    ByVal resultValue As Double, _
    ByVal uncertainty As Double, _
    ByVal limits As clsNormativeLimit) _
    As clsEvaluationResult

    If limits.HasLIE And limits.HasLSE Then

        Set EvaluateEU24 = EvaluateBothLimits( _
            resultValue, _
            uncertainty, _
            limits)

        Exit Function

    End If

    If limits.HasLIE Then

        Set EvaluateEU24 = EvaluateLowerLimit( _
            resultValue, _
            uncertainty, _
            limits.LIE)

        Exit Function

    End If

    If limits.HasLSE Then

        Set EvaluateEU24 = EvaluateUpperLimit( _
            resultValue, _
            uncertainty, _
            limits.LSE)

        Exit Function

    End If

    Set EvaluateEU24 = CreateEvaluationResult( _
        EvaluationStatus_NotApplicable, _
        EvaluationValue_NoAplica, _
        EvaluationReason_None)

End Function


Private Function EvaluateUpperLimit( _
    ByVal resultValue As Double, _
    ByVal uncertainty As Double, _
    ByVal upperLimit As Double) _
    As clsEvaluationResult

    Dim conditionalLower As Double
    Dim conditionalUpper As Double

    conditionalLower = upperLimit - uncertainty
    conditionalUpper = upperLimit + uncertainty

    If resultValue < conditionalLower Then

        Set EvaluateUpperLimit = CreateEvaluationResult( _
            EvaluationStatus_Evaluated, _
            EvaluationValue_Cumple, _
            EvaluationReason_None)

    ElseIf resultValue <= upperLimit Then

        Set EvaluateUpperLimit = CreateEvaluationResult( _
            EvaluationStatus_Evaluated, _
            EvaluationValue_CumpleCondicionado, _
            EvaluationReason_None)

    ElseIf resultValue <= conditionalUpper Then

        Set EvaluateUpperLimit = CreateEvaluationResult( _
            EvaluationStatus_Evaluated, _
            EvaluationValue_NoCumpleCondicionado, _
            EvaluationReason_None)

    Else

        Set EvaluateUpperLimit = CreateEvaluationResult( _
            EvaluationStatus_Evaluated, _
            EvaluationValue_NoCumple, _
            EvaluationReason_None)

    End If

End Function


Private Function EvaluateLowerLimit( _
    ByVal resultValue As Double, _
    ByVal uncertainty As Double, _
    ByVal lowerLimit As Double) _
    As clsEvaluationResult

    Dim conditionalLower As Double
    Dim conditionalUpper As Double

    conditionalLower = lowerLimit - uncertainty
    conditionalUpper = lowerLimit + uncertainty

    If resultValue < conditionalLower Then

        Set EvaluateLowerLimit = CreateEvaluationResult( _
            EvaluationStatus_Evaluated, _
            EvaluationValue_NoCumple, _
            EvaluationReason_None)

    ElseIf resultValue < lowerLimit Then

        Set EvaluateLowerLimit = CreateEvaluationResult( _
            EvaluationStatus_Evaluated, _
            EvaluationValue_NoCumpleCondicionado, _
            EvaluationReason_None)

    ElseIf resultValue <= conditionalUpper Then

        Set EvaluateLowerLimit = CreateEvaluationResult( _
            EvaluationStatus_Evaluated, _
            EvaluationValue_CumpleCondicionado, _
            EvaluationReason_None)

    Else

        Set EvaluateLowerLimit = CreateEvaluationResult( _
            EvaluationStatus_Evaluated, _
            EvaluationValue_Cumple, _
            EvaluationReason_None)

    End If

End Function


Private Function EvaluateBothLimits( _
    ByVal resultValue As Double, _
    ByVal uncertainty As Double, _
    ByVal limits As clsNormativeLimit) _
    As clsEvaluationResult

    Dim lowerConditionalLimit As Double
    Dim upperConditionalLimit As Double

    lowerConditionalLimit = limits.LIE + uncertainty
    upperConditionalLimit = limits.LSE - uncertainty

    If lowerConditionalLimit >= upperConditionalLimit Then

        Set EvaluateBothLimits = CreateEvaluationResult( _
            EvaluationStatus_CriteriaConflict, _
            EvaluationValue_NoAplica, _
            EvaluationReason_ZonasIncompatibles)

        Exit Function

    End If

    If resultValue < limits.LIE - uncertainty Then

        Set EvaluateBothLimits = CreateEvaluationResult( _
            EvaluationStatus_Evaluated, _
            EvaluationValue_NoCumple, _
            EvaluationReason_None)

    ElseIf resultValue < limits.LIE Then

        Set EvaluateBothLimits = CreateEvaluationResult( _
            EvaluationStatus_Evaluated, _
            EvaluationValue_NoCumpleCondicionado, _
            EvaluationReason_None)

    ElseIf resultValue <= lowerConditionalLimit Then

        Set EvaluateBothLimits = CreateEvaluationResult( _
            EvaluationStatus_Evaluated, _
            EvaluationValue_CumpleCondicionado, _
            EvaluationReason_None)

    ElseIf resultValue < upperConditionalLimit Then

        Set EvaluateBothLimits = CreateEvaluationResult( _
            EvaluationStatus_Evaluated, _
            EvaluationValue_Cumple, _
            EvaluationReason_None)

    ElseIf resultValue <= limits.LSE Then

        Set EvaluateBothLimits = CreateEvaluationResult( _
            EvaluationStatus_Evaluated, _
            EvaluationValue_CumpleCondicionado, _
            EvaluationReason_None)

    ElseIf resultValue <= limits.LSE + uncertainty Then

        Set EvaluateBothLimits = CreateEvaluationResult( _
            EvaluationStatus_Evaluated, _
            EvaluationValue_NoCumpleCondicionado, _
            EvaluationReason_None)

    Else

        Set EvaluateBothLimits = CreateEvaluationResult( _
            EvaluationStatus_Evaluated, _
            EvaluationValue_NoCumple, _
            EvaluationReason_None)

    End If

End Function

