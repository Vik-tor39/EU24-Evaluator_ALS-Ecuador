VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEvaluationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mNormativaRepository As clsNormativaRepository


Friend Sub Initialize( _
    ByVal normativaRepository As clsNormativaRepository)

    Set mNormativaRepository = normativaRepository

End Sub


Public Function Evaluate( _
    ByVal resultInput As clsResultInput, _
    ByVal normativaID As String) _
    As clsEvaluationRecord

    Dim resolution As clsVariableResolution
    Dim limits As clsNormativeLimit
    Dim applicability As clsApplicabilityResult
    Dim evaluation As clsEvaluationResult

    '--------------------------------------------------
    ' 1. Resolver variable
    '--------------------------------------------------

    Set resolution = _
        ResolveVariable(resultInput.variableName)

    If resolution.status <> VariableResolution_Resolved Then

        Set evaluation = _
            CreateEvaluationResult( _
                EvaluationStatus_NotApplicable, _
                EvaluationValue_NoAplica, _
                GetResolutionReason(resolution.status))

        Set Evaluate = _
            CreateEvaluationRecord( _
                resultInput, _
                resolution, _
                Nothing, _
                Nothing, _
                evaluation)

        Exit Function

    End If

    '--------------------------------------------------
    ' 2. Verificar que la variable esté contemplada
    '    por la normativa seleccionada
    '--------------------------------------------------

    If Not mNormativaRepository. _
        IsVariableConfiguredForNormativa( _
            normativaID, _
            resolution.variableID) Then

        Set evaluation = _
            CreateEvaluationResult( _
                EvaluationStatus_NotApplicable, _
                EvaluationValue_NoAplica, _
                EvaluationReason_VariableNoContempladaNormativa)

        Set Evaluate = _
            CreateEvaluationRecord( _
                resultInput, _
                resolution, _
                Nothing, _
                Nothing, _
                evaluation)

        Exit Function

    End If

    '--------------------------------------------------
    ' 3. Obtener límites
    '--------------------------------------------------

    Set limits = _
        mNormativaRepository.GetLimit( _
            normativaID, _
            resolution.variableID)

    If limits Is Nothing Then

        Set evaluation = _
            CreateEvaluationResult( _
                EvaluationStatus_NotApplicable, _
                EvaluationValue_NoAplica, _
                EvaluationReason_LimitesNoEncontrados)

        Set Evaluate = _
            CreateEvaluationRecord( _
                resultInput, _
                resolution, _
                Nothing, _
                Nothing, _
                evaluation)

        Exit Function

    End If

    '--------------------------------------------------
    ' 4. Validar consistencia de límites
    '--------------------------------------------------

    If HasInconsistentLimits(limits) Then

        Set evaluation = _
            CreateEvaluationResult( _
                EvaluationStatus_CriteriaConflict, _
                EvaluationValue_NoAplica, _
                EvaluationReason_LimitesInconsistentes)

        Set Evaluate = _
            CreateEvaluationRecord( _
                resultInput, _
                resolution, _
                limits, _
                Nothing, _
                evaluation)

        Exit Function

    End If

    '--------------------------------------------------
    ' 5. Verificar aplicabilidad del resultado
    '--------------------------------------------------

    Set applicability = _
        CheckApplicability(resultInput, limits)

    If Not applicability.isApplicable Then

        Set evaluation = _
            CreateEvaluationResult( _
                EvaluationStatus_NotApplicable, _
                EvaluationValue_NoAplica, _
                GetApplicabilityReason( _
                    applicability.reason))

        Set Evaluate = _
            CreateEvaluationRecord( _
                resultInput, _
                resolution, _
                limits, _
                applicability, _
                evaluation)

        Exit Function

    End If

    '--------------------------------------------------
    ' 6. Seleccionar y ejecutar el evaluador
    '--------------------------------------------------

    Select Case applicability.modoEvaluacion

        Case EvaluationMode_EU24

            Set evaluation = EvaluateEU24( _
                applicability.numericValue, _
                applicability.numericUncertainty, _
                limits)

        Case EvaluationMode_OperatorLessThan

            Set evaluation = EvaluateOperatorLessThan( _
                applicability.numericValue, _
                limits)

        Case EvaluationMode_NotApplicable

            Set evaluation = CreateEvaluationResult( _
                EvaluationStatus_NotApplicable, _
                EvaluationValue_NoAplica, _
                EvaluationReason_None)

        Case Else

            ' Defensa ante un modo de evaluación desconocido.
            Set evaluation = CreateEvaluationResult( _
                EvaluationStatus_NotApplicable, _
                EvaluationValue_NoAplica, _
                EvaluationReason_None)

    End Select

    '--------------------------------------------------
    ' 7. Construir registro completo
    '--------------------------------------------------

    Set Evaluate = _
        CreateEvaluationRecord( _
            resultInput, _
            resolution, _
            limits, _
            applicability, _
            evaluation)

End Function


Private Function ResolveVariable( _
    ByVal variableName As String) _
    As clsVariableResolution

    Set ResolveVariable = ResolveVariableName( _
        variableName, _
        mNormativaRepository)

End Function


Private Function GetResolutionReason( _
    ByVal status As VariableResolutionStatus) _
    As EvaluationReason

    Select Case status

        Case VariableResolution_NotFound

            GetResolutionReason = _
                EvaluationReason_VariableNoEncontrada

        Case VariableResolution_Ambiguous

            GetResolutionReason = _
                EvaluationReason_VariableAmbigua

        Case Else

            GetResolutionReason = _
                EvaluationReason_None

    End Select

End Function

Private Function GetApplicabilityReason( _
    ByVal reason As ApplicabilityReason) _
    As EvaluationReason

    Select Case reason

        Case ApplicabilityReason_ResultadoNoNumerico

            GetApplicabilityReason = _
                EvaluationReason_ResultadoNoNumerico

        Case ApplicabilityReason_ResultadoConOperador

            GetApplicabilityReason = _
                EvaluationReason_ResultadoConOperador

        Case ApplicabilityReason_IncertidumbreNoDisponible

            GetApplicabilityReason = _
                EvaluationReason_IncertidumbreNoDisponible

        Case ApplicabilityReason_LimitesNoDefinidos

            GetApplicabilityReason = _
                EvaluationReason_LimitesNoEncontrados

        Case Else

            GetApplicabilityReason = _
                EvaluationReason_None

    End Select

End Function

