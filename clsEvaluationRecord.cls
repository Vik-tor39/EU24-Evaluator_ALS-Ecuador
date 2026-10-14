VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEvaluationRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mInput As clsResultInput
Private mVariableResolution As clsVariableResolution
Private mNormativeLimit As clsNormativeLimit
Private mApplicability As clsApplicabilityResult
Private mEvaluation As clsEvaluationResult

Public Property Get resultInput() As clsResultInput
    Set resultInput = mInput
End Property


Public Property Get variableResolution() As clsVariableResolution
    Set variableResolution = mVariableResolution
End Property


Public Property Get normativeLimit() As clsNormativeLimit
    Set normativeLimit = mNormativeLimit
End Property


Public Property Get applicability() As clsApplicabilityResult
    Set applicability = mApplicability
End Property


Public Property Get evaluation() As clsEvaluationResult
    Set evaluation = mEvaluation
End Property

Friend Sub Initialize( _
    ByVal resultInput As clsResultInput, _
    ByVal variableResolution As clsVariableResolution, _
    ByVal normativeLimit As clsNormativeLimit, _
    ByVal applicability As clsApplicabilityResult, _
    ByVal evaluation As clsEvaluationResult)

    Set mInput = resultInput
    Set mVariableResolution = variableResolution
    Set mNormativeLimit = normativeLimit
    Set mApplicability = applicability
    Set mEvaluation = evaluation

End Sub

