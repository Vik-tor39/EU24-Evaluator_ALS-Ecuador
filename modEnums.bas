Attribute VB_Name = "modEnums"
Option Explicit

'============================================================
' Estado general de una evaluación
'============================================================
Public Enum EvaluationStatus
    EvaluationStatus_Evaluated = 1
    EvaluationStatus_NotApplicable = 2
    EvaluationStatus_NotDeterminable = 3
    EvaluationStatus_CriteriaConflict = 4
End Enum


'============================================================
' Resultado de conformidad que verá el usuario
'============================================================
Public Enum evaluationValue
    EvaluationValue_Cumple = 1
    EvaluationValue_CumpleCondicionado = 2
    EvaluationValue_NoCumpleCondicionado = 3
    EvaluationValue_NoCumple = 4
    EvaluationValue_NoAplica = 5
    EvaluationValue_NoDeterminado = 6
End Enum


'============================================================
' Motivo de un resultado de evaluación
'============================================================
Public Enum EvaluationReason
    EvaluationReason_None = 0
    EvaluationReason_ZonasIncompatibles = 1
    EvaluationReason_VariableNoEncontrada = 2
    EvaluationReason_VariableAmbigua = 3
    EvaluationReason_LimitesNoEncontrados = 4
    EvaluationReason_ResultadoNoNumerico = 5
    EvaluationReason_ResultadoConOperador = 6
    EvaluationReason_IncertidumbreNoDisponible = 7
    EvaluationReason_VariableNoContempladaNormativa = 8
    EvaluationReason_LimitesInconsistentes = 9
End Enum


'============================================================
' Motivos por los que un resultado no puede ser evaluado
'============================================================
Public Enum ApplicabilityReason
    ApplicabilityReason_None = 0
    ApplicabilityReason_ResultadoNoNumerico = 1
    ApplicabilityReason_ResultadoConOperador = 2
    ApplicabilityReason_IncertidumbreNoDisponible = 3
    ApplicabilityReason_LimitesNoDefinidos = 4
End Enum


'============================================================
' Estado de resolución de una variable
'============================================================
Public Enum VariableResolutionStatus
    VariableResolution_Resolved = 1
    VariableResolution_NotFound = 2
    VariableResolution_Ambiguous = 3
End Enum


'============================================================
' Forma en que se encontró la correspondencia
'============================================================
Public Enum VariableMatchType
    VariableMatch_None = 0
    VariableMatch_Exact = 1
    VariableMatch_Normalized = 2
    VariableMatch_ExplicitMapping = 3
End Enum


'============================================================
' Motivos de invalidación de la entrada
'============================================================
Public Enum InputValidationReason
    InputValidation_None = 0
    InputValidation_NoSelection = 1
    InputValidation_NoData = 2
    InputValidation_VariableColumnMissing = 3
    InputValidation_ValueColumnMissing = 4
    InputValidation_UncertaintyColumnMissing = 5
End Enum


'============================================================
' Modos de evaluación de la entrada
'============================================================
Public Enum EvaluationMode
    EvaluationMode_NotApplicable = 0
    EvaluationMode_EU24 = 1
    EvaluationMode_OperatorLessThan = 2
End Enum
