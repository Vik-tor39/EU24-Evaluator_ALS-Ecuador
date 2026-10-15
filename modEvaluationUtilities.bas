Attribute VB_Name = "modEvaluationUtilities"
Option Explicit

Public Function GetEvaluationDestination( _
    ByVal inputRange As Range) _
    As Range

    If inputRange Is Nothing Then
        Err.Raise _
            vbObjectError + 1300, _
            "modEvaluationUtilities", _
            "El rango de entrada no puede ser Nothing."
    End If

    Set GetEvaluationDestination = _
        inputRange.Worksheet.Cells( _
            inputRange.row, _
            inputRange.Column + inputRange.columns.Count)

End Function

Public Sub ValidateEvaluationDestination( _
    ByVal destination As Range)

    Const RESULT_COLUMN_COUNT As Long = 5

    If destination Is Nothing Then
        Err.Raise _
            vbObjectError + 1301, _
            "modEvaluationUtilities", _
            "El destino no puede ser Nothing."
    End If

    If destination.Column + RESULT_COLUMN_COUNT - 1 > _
        destination.Worksheet.columns.Count Then

        Err.Raise _
            vbObjectError + 1302, _
            "modEvaluationUtilities", _
            "No existe espacio suficiente para colocar " & _
            "el bloque de resultados."
    End If

End Sub

