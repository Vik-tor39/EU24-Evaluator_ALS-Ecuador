VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNormativeLimit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Modelo de dominio para entidad Limite.

Option Explicit

Private mNormativaID As String 'Id Normativa
Private mVariableID As String 'Id Variable

Private mHasLIE As Boolean 'Tiene LIE? Garantiza diferenciar entre valores y contenido vacio: 0 != Empty
Private mLIE As Double 'LIE

Private mHasLSE As Boolean 'Idem mHasLIE
Private mLSE As Double ' LSE


Public Property Get normativaID() As String
    normativaID = mNormativaID
End Property

Public Property Let normativaID(ByVal value As String)
    mNormativaID = value
End Property


Public Property Get variableID() As String
    variableID = mVariableID
End Property

Public Property Let variableID(ByVal value As String)
    mVariableID = value
End Property


Public Property Get HasLIE() As Boolean
    HasLIE = mHasLIE
End Property

Public Property Let HasLIE(ByVal value As Boolean)
    mHasLIE = value
End Property


Public Property Get LIE() As Double
    LIE = mLIE
End Property

Public Property Let LIE(ByVal value As Double)
    mLIE = value
End Property


Public Property Get HasLSE() As Boolean
    HasLSE = mHasLSE
End Property

Public Property Let HasLSE(ByVal value As Boolean)
    mHasLSE = value
End Property


Public Property Get LSE() As Double
    LSE = mLSE
End Property

Public Property Let LSE(ByVal value As Double)
    mLSE = value
End Property
