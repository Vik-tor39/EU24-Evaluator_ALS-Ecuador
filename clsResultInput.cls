VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsResultInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Modelo de dominio para entidad Valores en crudo

Option Explicit

Private mVariableName As String 'Nombre de var
Private mRawValue As String 'Resultado
Private mRawUncertainty As String 'Incertidumbre


Public Property Get variableName() As String
    variableName = mVariableName
End Property

Public Property Let variableName(ByVal value As String)
    mVariableName = value
End Property


Public Property Get rawValue() As String
    rawValue = mRawValue
End Property

Public Property Let rawValue(ByVal value As String)
    mRawValue = value
End Property


Public Property Get rawUncertainty() As String
    rawUncertainty = mRawUncertainty
End Property

Public Property Let rawUncertainty(ByVal value As String)
    mRawUncertainty = value
End Property

Friend Sub Initialize( _
    ByVal variableName As String, _
    ByVal rawValue As String, _
    ByVal rawUncertainty As String)

    mVariableName = variableName
    mRawValue = rawValue
    mRawUncertainty = rawUncertainty

End Sub
