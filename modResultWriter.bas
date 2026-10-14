Attribute VB_Name = "modResultWriter"
Option Explicit

Public Sub WriteResults( _
    ByVal inputRange As Range, _
    ByVal records As Collection, _
    ByVal normativa As clsNormative, _
    ByVal destination As Range)

    Dim outputRange As Range
    Dim outputData() As Variant

    Dim rowCount As Long
    Dim sourceColumnCount As Long
    Dim outputColumnCount As Long

    Dim rowIndex As Long
    Dim columnIndex As Long

    Dim record As clsEvaluationRecord

    '--------------------------------------------------
    ' Validaciones básicas del contrato
    '--------------------------------------------------

    If inputRange Is Nothing Then
        Err.Raise vbObjectError + 1000, _
                  "modResultWriter", _
                  "El rango de entrada no puede ser Nothing."
    End If

    If records Is Nothing Then
        Err.Raise vbObjectError + 1001, _
                  "modResultWriter", _
                  "La colección de evaluaciones no puede ser Nothing."
    End If

    If normativa Is Nothing Then
        Err.Raise vbObjectError + 1002, _
                  "modResultWriter", _
                  "La normativa no puede ser Nothing."
    End If

    If destination Is Nothing Then
        Err.Raise vbObjectError + 1003, _
                  "modResultWriter", _
                  "El destino no puede ser Nothing."
    End If

    '--------------------------------------------------
    ' La colección debe tener un registro por fila
    ' de datos.
    '--------------------------------------------------

    rowCount = inputRange.Rows.Count - 1

    If records.Count <> rowCount Then
        Err.Raise vbObjectError + 1004, _
                  "modResultWriter", _
                  "La cantidad de evaluaciones no coincide " & _
                  "con la cantidad de filas de datos."
    End If

    sourceColumnCount = inputRange.columns.Count

    '--------------------------------------------------
    ' Columnas adicionales:
    '
    ' Normativa
    ' LIE
    ' LSE
    ' Evaluación
    ' Motivo
    '--------------------------------------------------

    outputColumnCount = sourceColumnCount + 5

    ReDim outputData( _
        1 To rowCount + 1, _
        1 To outputColumnCount)

    '--------------------------------------------------
    ' Copiar encabezados originales
    '--------------------------------------------------

    For columnIndex = 1 To sourceColumnCount

        outputData(1, columnIndex) = _
            inputRange.Cells(1, columnIndex).value

    Next columnIndex

    '--------------------------------------------------
    ' Encabezados adicionales
    '--------------------------------------------------

    outputData(1, sourceColumnCount + 1) = "Normativa"
    outputData(1, sourceColumnCount + 2) = "LIE"
    outputData(1, sourceColumnCount + 3) = "LSE"
    outputData(1, sourceColumnCount + 4) = "Evaluación"
    outputData(1, sourceColumnCount + 5) = "Motivo"

    '--------------------------------------------------
    ' Datos
    '--------------------------------------------------

    For rowIndex = 1 To rowCount

        Set record = records(rowIndex)

        '----------------------------------------------
        ' Datos originales
        '----------------------------------------------

        For columnIndex = 1 To sourceColumnCount

            outputData(rowIndex + 1, columnIndex) = _
                inputRange.Cells( _
                    rowIndex + 1, _
                    columnIndex).value

        Next columnIndex

        '----------------------------------------------
        ' Normativa
        '----------------------------------------------

        outputData( _
            rowIndex + 1, _
            sourceColumnCount + 1) = normativa.Name

        '----------------------------------------------
        ' Límites
        '----------------------------------------------

        WriteLimitValues _
            outputData, _
            rowIndex + 1, _
            sourceColumnCount + 2, _
            record

        '----------------------------------------------
        ' Evaluación
        '----------------------------------------------

        outputData( _
            rowIndex + 1, _
            sourceColumnCount + 4) = _
                GetEvaluationValueText( _
                    record.evaluation.value)

        outputData( _
            rowIndex + 1, _
            sourceColumnCount + 5) = _
                GetEvaluationReasonText( _
                    record.evaluation.reason)

    Next rowIndex

    '--------------------------------------------------
    ' Escribir todo de una sola vez
    '--------------------------------------------------

    Set outputRange = destination.Resize( _
        rowCount + 1, _
        outputColumnCount)

    outputRange.value = outputData

End Sub

Private Sub WriteLimitValues( _
    ByRef outputData() As Variant, _
    ByVal rowIndex As Long, _
    ByVal firstLimitColumn As Long, _
    ByVal record As clsEvaluationRecord)

    If record.normativeLimit Is Nothing Then

        outputData(rowIndex, firstLimitColumn) = vbNullString
        outputData(rowIndex, firstLimitColumn + 1) = vbNullString

        Exit Sub

    End If

    If record.normativeLimit.HasLIE Then

        outputData(rowIndex, firstLimitColumn) = _
            record.normativeLimit.LIE

    Else

        outputData(rowIndex, firstLimitColumn) = _
            vbNullString

    End If

    If record.normativeLimit.HasLSE Then

        outputData(rowIndex, firstLimitColumn + 1) = _
            record.normativeLimit.LSE

    Else

        outputData(rowIndex, firstLimitColumn + 1) = _
            vbNullString

    End If

End Sub


Private Function GetEvaluationValueText( _
    ByVal evaluationValue As evaluationValue) _
    As String

    Select Case evaluationValue

        Case EvaluationValue_Cumple

            GetEvaluationValueText = "CUMPLE"

        Case EvaluationValue_CumpleCondicionado

            GetEvaluationValueText = _
                "CUMPLE CONDICIONADO"

        Case EvaluationValue_NoCumpleCondicionado

            GetEvaluationValueText = _
                "NO CUMPLE CONDICIONADO"

        Case EvaluationValue_NoCumple

            GetEvaluationValueText = "NO CUMPLE"

        Case EvaluationValue_NoDeterminado

            GetEvaluationValueText = "NO DETERMINADO"

        Case EvaluationValue_NoAplica

            GetEvaluationValueText = "NO APLICA"

        Case Else

            GetEvaluationValueText = "NO APLICA"

    End Select

End Function


Private Function GetEvaluationReasonText( _
    ByVal reason As EvaluationReason) _
    As String

    Select Case reason

        Case EvaluationReason_None

            GetEvaluationReasonText = vbNullString

        Case EvaluationReason_ZonasIncompatibles

            GetEvaluationReasonText = _
                "ZONAS INCOMPATIBLES"

        Case EvaluationReason_VariableNoEncontrada

            GetEvaluationReasonText = _
                "VARIABLE NO ENCONTRADA"

        Case EvaluationReason_VariableAmbigua

            GetEvaluationReasonText = _
                "VARIABLE AMBIGUA"

        Case EvaluationReason_LimitesNoEncontrados

            GetEvaluationReasonText = _
                "LÍMITES NO ENCONTRADOS"
                
        Case EvaluationReason_VariableNoContempladaNormativa
            
            GetEvaluationReasonText = _
                "VARIABLE NO CONTEMPLADA EN LA NORMATIVA"

        Case EvaluationReason_ResultadoNoNumerico

            GetEvaluationReasonText = _
                "RESULTADO NO NUMÉRICO"

        Case EvaluationReason_ResultadoConOperador

            GetEvaluationReasonText = _
                "RESULTADO CON OPERADOR"

        Case EvaluationReason_IncertidumbreNoDisponible
    
            GetEvaluationReasonText = _
                "INCERTIDUMBRE NO DISPONIBLE"
                
        Case EvaluationReason_LimitesInconsistentes
    
            GetEvaluationReasonText = _
                "LÍMITES NORMATIVOS INCONSISTENTES"

        Case Else

            GetEvaluationReasonText = _
                "RAZÓN NO DEFINIDA"

    End Select

End Function


