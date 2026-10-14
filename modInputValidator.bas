Attribute VB_Name = "modInputValidator"
Option Explicit

Public Function ValidateInputRange( _
    ByVal inputRange As Range) _
    As clsInputValidationResult

    Dim columns As ColumnReferences

    '--------------------------------------------------
    ' 1. Validar selección
    '--------------------------------------------------
    If inputRange Is Nothing Then
        Set ValidateInputRange = _
            CreateInputValidationResult( _
                False, _
                InputValidation_NoSelection, _
                columns)

        Exit Function
    End If

    '--------------------------------------------------
    ' 2. Validar que existan filas de datos
    '--------------------------------------------------
    If inputRange.Rows.Count < 2 Then
        Set ValidateInputRange = _
            CreateInputValidationResult( _
                False, _
                InputValidation_NoData, _
                columns)

        Exit Function
    End If

    '--------------------------------------------------
    ' 3. Buscar columna de variable
    '--------------------------------------------------
    columns.VariableColumn = _
        FindColumn(inputRange, "Nombre")

    If columns.VariableColumn = 0 Then
        Set ValidateInputRange = _
            CreateInputValidationResult( _
                False, _
                InputValidation_VariableColumnMissing, _
                columns)

        Exit Function
    End If

    '--------------------------------------------------
    ' 4. Buscar columna Valor
    '--------------------------------------------------
    columns.ValueColumn = _
        FindColumn(inputRange, "Valor")

    If columns.ValueColumn = 0 Then
        Set ValidateInputRange = _
            CreateInputValidationResult( _
                False, _
                InputValidation_ValueColumnMissing, _
                columns)

        Exit Function
    End If

    '--------------------------------------------------
    ' 5. Buscar columna Incertidumbre
    '--------------------------------------------------
    columns.UncertaintyColumn = _
        FindColumn(inputRange, "Incertidumbre")

    If columns.UncertaintyColumn = 0 Then
        Set ValidateInputRange = _
            CreateInputValidationResult( _
                False, _
                InputValidation_UncertaintyColumnMissing, _
                columns)

        Exit Function
    End If

    '--------------------------------------------------
    ' 6. Entrada válida
    '--------------------------------------------------
    Set ValidateInputRange = _
        CreateInputValidationResult( _
            True, _
            InputValidation_None, _
            columns)

End Function


Private Function FindColumn( _
    ByVal inputRange As Range, _
    ByVal columnName As String) _
    As Long

    Dim columnIndex As Long
    Dim headerValue As String

    For columnIndex = 1 To inputRange.columns.Count

        headerValue = _
            Trim$(CStr(inputRange.Cells(1, columnIndex).value))

        If StrComp( _
            headerValue, _
            columnName, _
            vbTextCompare) = 0 Then

            FindColumn = columnIndex
            Exit Function

        End If

    Next columnIndex

    FindColumn = 0

End Function

