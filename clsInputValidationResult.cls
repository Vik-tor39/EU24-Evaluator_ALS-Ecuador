VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInputValidationResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mIsValid As Boolean
Private mReason As InputValidationReason
Private mColumns As ColumnReferences


Public Property Get isValid() As Boolean
    isValid = mIsValid
End Property


Public Property Get reason() As InputValidationReason
    reason = mReason
End Property


Public Property Get columns() As ColumnReferences
    columns = mColumns
End Property

Friend Sub Initialize( _
    ByVal isValid As Boolean, _
    ByVal reason As InputValidationReason, _
    ByRef columns As ColumnReferences)

    mIsValid = isValid
    mReason = reason
    mColumns = columns

End Sub
