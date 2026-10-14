VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNormative"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Modelo de dominio para entidad Normativa

Option Explicit

Private mID As String 'Id
Private mName As String 'Nombre
Private mActive As Boolean 'Es activa?


Public Property Get ID() As String
    ID = mID
End Property

Public Property Let ID(ByVal value As String)
    mID = value
End Property


Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal value As String)
    mName = value
End Property


Public Property Get Active() As Boolean
    Active = mActive
End Property

Public Property Let Active(ByVal value As Boolean)
    mActive = value
End Property
