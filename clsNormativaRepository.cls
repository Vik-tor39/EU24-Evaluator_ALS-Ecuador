VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNormativaRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mWorkbook As Workbook

Private mNormativasTable As ListObject
Private mVariablesTable As ListObject
Private mLimitesTable As ListObject


Friend Sub Initialize( _
    ByVal configurationWorkbook As Workbook)

    Set mWorkbook = configurationWorkbook

    Set mNormativasTable = _
        GetTable("Normativas")

    Set mVariablesTable = _
        GetTable("Variables")

    Set mLimitesTable = _
        GetTable("Limites")

End Sub


Public Function GetNormativa( _
    ByVal normativaID As String) _
    As clsNormative

    Dim row As ListRow
    Dim idValue As String

    For Each row In mNormativasTable.ListRows

        idValue = CStr( _
            row.Range.Cells(1, _
                mNormativasTable.ListColumns("ID").Index).value)

        If StrComp( _
            Trim$(idValue), _
            Trim$(normativaID), _
            vbTextCompare) = 0 Then

            Dim normativa As clsNormative

            Set normativa = New clsNormative

            normativa.ID = idValue

            normativa.Name = CStr( _
                row.Range.Cells(1, _
                    mNormativasTable.ListColumns("Nombre").Index).value)

            normativa.Active = CBool( _
                row.Range.Cells(1, _
                    mNormativasTable.ListColumns("Activa").Index).value)

            Set GetNormativa = normativa

            Exit Function

        End If

    Next row

End Function


Public Function GetLimit( _
    ByVal normativaID As String, _
    ByVal variableID As String) _
    As clsNormativeLimit

    Dim row As ListRow

    Dim normativaValue As String
    Dim variableValue As String

    For Each row In mLimitesTable.ListRows

        normativaValue = CStr( _
            row.Range.Cells(1, _
                mLimitesTable.ListColumns("NormativaID").Index).value)

        variableValue = CStr( _
            row.Range.Cells(1, _
                mLimitesTable.ListColumns("VariableID").Index).value)

        If StrComp( _
            Trim$(normativaValue), _
            Trim$(normativaID), _
            vbTextCompare) = 0 _
        And _
           StrComp( _
               Trim$(variableValue), _
               Trim$(variableID), _
               vbTextCompare) = 0 Then

            Set GetLimit = BuildNormativeLimit(row)

            Exit Function

        End If

    Next row

End Function


Public Function GetVariables() As Collection

    Dim result As Collection
    Dim row As ListRow

    Dim variableID As String
    Dim variableName As String

    Set result = New Collection

    For Each row In mVariablesTable.ListRows

        variableID = CStr( _
            row.Range.Cells(1, _
                mVariablesTable.ListColumns("ID").Index).value)

        variableName = CStr( _
            row.Range.Cells(1, _
                mVariablesTable.ListColumns("Nombre").Index).value)

        result.Add Array(variableID, variableName)

    Next row

    Set GetVariables = result

End Function


Public Function GetVariableMappings() As Collection

    Dim result As Collection
    Dim ws As Worksheet
    Dim table As ListObject
    Dim row As ListRow

    Dim variableID As String
    Dim resultName As String
    Dim normativeName As String

    Set result = New Collection

    Set table = GetTable("VariableMap")

    For Each row In table.ListRows

        variableID = CStr( _
            row.Range.Cells(1, _
                table.ListColumns("VariableID").Index).value)

        resultName = CStr( _
            row.Range.Cells(1, _
                table.ListColumns("NombreResultado").Index).value)

        normativeName = CStr( _
            row.Range.Cells(1, _
                table.ListColumns("NombreNormativa").Index).value)

        result.Add Array( _
            variableID, _
            resultName, _
            normativeName)

    Next row

    Set GetVariableMappings = result

End Function


Private Function BuildNormativeLimit( _
    ByVal row As ListRow) _
    As clsNormativeLimit

    Dim result As clsNormativeLimit

    Set result = New clsNormativeLimit

    result.normativaID = CStr( _
        row.Range.Cells(1, _
            mLimitesTable.ListColumns("NormativaID").Index).value)

    result.variableID = CStr( _
        row.Range.Cells(1, _
            mLimitesTable.ListColumns("VariableID").Index).value)

    SetLimitValues result, row

    Set BuildNormativeLimit = result

End Function


Private Sub SetLimitValues( _
    ByVal result As clsNormativeLimit, _
    ByVal row As ListRow)

    Dim lieValue As Variant
    Dim lseValue As Variant

    lieValue = row.Range.Cells(1, _
        mLimitesTable.ListColumns("LIE").Index).value

    lseValue = row.Range.Cells(1, _
        mLimitesTable.ListColumns("LSE").Index).value

    If Len(Trim$(CStr(lieValue))) > 0 Then

        result.HasLIE = True
        result.LIE = CDbl(lieValue)

    Else

        result.HasLIE = False

    End If

    If Len(Trim$(CStr(lseValue))) > 0 Then

        result.HasLSE = True
        result.LSE = CDbl(lseValue)

    Else

        result.HasLSE = False

    End If

End Sub


Private Function GetTable( _
    ByVal tableName As String) _
    As ListObject

    Dim ws As Worksheet
    Dim table As ListObject

    For Each ws In mWorkbook.Worksheets

        For Each table In ws.ListObjects

            If StrComp( _
                table.Name, _
                tableName, _
                vbTextCompare) = 0 Then

                Set GetTable = table

                'Debug.Print table.Name

                Exit Function

            End If

        Next table

    Next ws

    Err.Raise _
        vbObjectError + 1000, _
        "clsNormativaRepository", _
        "No se encontró la tabla de configuración '" _
        & tableName & "'."

End Function

Public Function IsVariableConfiguredForNormativa( _
    ByVal normativaID As String, _
    ByVal variableID As String) _
    As Boolean

    Dim row As ListRow

    Dim normativaValue As String
    Dim variableValue As String

    For Each row In mLimitesTable.ListRows

        normativaValue = _
            CStr(row.Range.Cells( _
                1, _
                mLimitesTable.ListColumns("NormativaID").Index).value)

        variableValue = _
            CStr(row.Range.Cells( _
                1, _
                mLimitesTable.ListColumns("VariableID").Index).value)

        If StrComp( _
            normativaValue, _
            normativaID, _
            vbTextCompare) = 0 Then

            If StrComp( _
                variableValue, _
                variableID, _
                vbTextCompare) = 0 Then

                IsVariableConfiguredForNormativa = True
                Exit Function

            End If

        End If

    Next row

    IsVariableConfiguredForNormativa = False

End Function


