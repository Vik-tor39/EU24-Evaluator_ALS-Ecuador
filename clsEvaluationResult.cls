VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEvaluationResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mStatus As EvaluationStatus
Private mValue As evaluationValue
Private mReason As EvaluationReason


Public Property Get status() As EvaluationStatus
    status = mStatus
End Property


Public Property Get value() As evaluationValue
    value = mValue
End Property


Public Property Get reason() As EvaluationReason
    reason = mReason
End Property

Friend Sub Initialize( _
    ByVal status As EvaluationStatus, _
    ByVal value As evaluationValue, _
    ByVal reason As EvaluationReason)

    mStatus = status
    mValue = value
    mReason = reason

End Sub

