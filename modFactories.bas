Attribute VB_Name = "modFactories"
Option Explicit

Public Function CreateVariableResolution( _
    ByVal status As VariableResolutionStatus, _
    ByVal variableID As String, _
    ByVal resultName As String, _
    ByVal normativeName As String, _
    ByVal matchType As VariableMatchType) _
    As clsVariableResolution

    Dim result As clsVariableResolution
    Set result = New clsVariableResolution

    result.Initialize _
        status, _
        variableID, _
        resultName, _
        normativeName, _
        matchType

    Set CreateVariableResolution = result

End Function


Public Function CreateApplicabilityResult( _
    ByVal isApplicable As Boolean, _
    ByVal reason As ApplicabilityReason, _
    ByVal numericValue As Double, _
    ByVal numericUncertainty As Double, _
    ByVal modoEvaluacion As EvaluationMode) _
    As clsApplicabilityResult

    Dim result As clsApplicabilityResult
    Set result = New clsApplicabilityResult

    result.Initialize _
        isApplicable, _
        reason, _
        numericValue, _
        numericUncertainty, _
        modoEvaluacion

    Set CreateApplicabilityResult = result

End Function


Public Function CreateEvaluationResult( _
    ByVal status As EvaluationStatus, _
    ByVal value As evaluationValue, _
    ByVal reason As EvaluationReason) _
    As clsEvaluationResult

    Dim result As clsEvaluationResult
    Set result = New clsEvaluationResult

    result.Initialize _
        status, _
        value, _
        reason

    Set CreateEvaluationResult = result

End Function


Public Function CreateEvaluationRecord( _
    ByVal resultInput As clsResultInput, _
    ByVal variableResolution As clsVariableResolution, _
    ByVal normativeLimit As clsNormativeLimit, _
    ByVal applicability As clsApplicabilityResult, _
    ByVal evaluation As clsEvaluationResult) _
    As clsEvaluationRecord

    Dim result As clsEvaluationRecord
    Set result = New clsEvaluationRecord

    result.Initialize _
        resultInput, _
        variableResolution, _
        normativeLimit, _
        applicability, _
        evaluation

    Set CreateEvaluationRecord = result

End Function


Public Function CreateInputValidationResult( _
    ByVal isValid As Boolean, _
    ByVal reason As InputValidationReason, _
    ByRef columns As ColumnReferences) _
    As clsInputValidationResult

    Dim result As clsInputValidationResult
    Set result = New clsInputValidationResult

    result.Initialize _
        isValid, _
        reason, _
        columns

    Set CreateInputValidationResult = result

End Function

Public Function CreateResultInput( _
    ByVal variableName As String, _
    ByVal rawValue As String, _
    ByVal rawUncertainty As String) _
    As clsResultInput
    
    Dim result As clsResultInput
    Set result = New clsResultInput
    
    result.Initialize _
        variableName, _
        rawValue, _
        rawUncertainty
        
    Set CreateResultInput = result
        
End Function
