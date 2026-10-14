Attribute VB_Name = "modInputReader"
Option Explicit

Public Function ReadInputs( _
    ByVal inputRange As Range, _
    ByRef columns As ColumnReferences) _
    As Collection

    Dim result As Collection
    Dim rowIndex As Long

    Dim variableName As String
    Dim rawValue As String
    Dim rawUncertainty As String

    Dim resultInput As clsResultInput

    Set result = New Collection

    '--------------------------------------------------
    ' Las columnas se reciben como posiciones relativas
    ' dentro del rango seleccionado.
    '--------------------------------------------------

    For rowIndex = 2 To inputRange.Rows.Count

        variableName = _
            CStr(inputRange.Cells(rowIndex, columns.VariableColumn).value)

        rawValue = _
            CStr(inputRange.Cells(rowIndex, columns.ValueColumn).value)

        rawUncertainty = _
            CStr(inputRange.Cells( _
                rowIndex, _
                columns.UncertaintyColumn).value)

        Set resultInput = CreateResultInput( _
            variableName, _
            rawValue, _
            rawUncertainty)

        result.Add resultInput

    Next rowIndex

    Set ReadInputs = result

End Function

