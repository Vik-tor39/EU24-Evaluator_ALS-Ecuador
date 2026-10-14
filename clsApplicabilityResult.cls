VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsApplicabilityResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private mIsApplicable As Boolean
Private mReason As ApplicabilityReason
Private mNumericValue As Double
Private mNumericUncertainty As Double
Private mEvaluationMode As EvaluationMode


Public Property Get isApplicable() As Boolean
    isApplicable = mIsApplicable
End Property


Public Property Get reason() As ApplicabilityReason
    reason = mReason
End Property


Public Property Get numericValue() As Double
    numericValue = mNumericValue
End Property


Public Property Get numericUncertainty() As Double
    numericUncertainty = mNumericUncertainty
End Property


Public Property Get modoEvaluacion() As EvaluationMode
    modoEvaluacion = mEvaluationMode
End Property


Friend Sub Initialize( _
    ByVal isApplicable As Boolean, _
    ByVal reason As ApplicabilityReason, _
    ByVal numericValue As Double, _
    ByVal numericUncertainty As Double, _
    ByVal modoEvaluacion As EvaluationMode)

    mIsApplicable = isApplicable
    mReason = reason
    mNumericValue = numericValue
    mNumericUncertainty = numericUncertainty
    mEvaluationMode = modoEvaluacion

End Sub

