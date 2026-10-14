VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVariableResolution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mStatus As VariableResolutionStatus
Private mVariableID As String
Private mResultName As String
Private mNormativeName As String
Private mMatchType As VariableMatchType


Public Property Get status() As VariableResolutionStatus
    status = mStatus
End Property


Public Property Get variableID() As String
    variableID = mVariableID
End Property


Public Property Get resultName() As String
    resultName = mResultName
End Property


Public Property Get normativeName() As String
    normativeName = mNormativeName
End Property


Public Property Get matchType() As VariableMatchType
    matchType = mMatchType
End Property

Friend Sub Initialize( _
    ByVal status As VariableResolutionStatus, _
    ByVal variableID As String, _
    ByVal resultName As String, _
    ByVal normativeName As String, _
    ByVal matchType As VariableMatchType)

    mStatus = status
    mVariableID = variableID
    mResultName = resultName
    mNormativeName = normativeName
    mMatchType = matchType

End Sub
