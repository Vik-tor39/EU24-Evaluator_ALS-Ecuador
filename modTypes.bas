Attribute VB_Name = "modTypes"
Option Explicit

'A tomar en cuenta:
'- TODA UDT debe ser refenciada como ByRef por otros procedures

Public Type ColumnReferences
    VariableColumn As Long
    ValueColumn As Long
    UncertaintyColumn As Long
End Type
