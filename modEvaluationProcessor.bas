Attribute VB_Name = "modEvaluationProcessor"
Option Explicit

Public Sub ProcessEvaluation( _
    ByVal inputRange As Range, _
    ByVal normativaID As String, _
    ByVal repository As clsNormativaRepository, _
    ByVal destination As Range)

    Dim validation As clsInputValidationResult
    Dim inputs As Collection
    Dim records As Collection

    Dim service As clsEvaluationService
    Dim normativa As clsNormative

    Dim resultInput As clsResultInput
    Dim record As clsEvaluationRecord

    Dim i As Long

    '--------------------------------------------------
    ' 1. Validar dependencias
    '--------------------------------------------------

    If repository Is Nothing Then
        Err.Raise vbObjectError + 1100, _
                  "modEvaluationProcessor", _
                  "El repositorio no puede ser Nothing."
    End If

    If inputRange Is Nothing Then
        Err.Raise vbObjectError + 1101, _
                  "modEvaluationProcessor", _
                  "El rango de entrada no puede ser Nothing."
    End If

    If destination Is Nothing Then
        Err.Raise vbObjectError + 1102, _
                  "modEvaluationProcessor", _
                  "El destino no puede ser Nothing."
    End If

    If Len(Trim$(normativaID)) = 0 Then
        Err.Raise vbObjectError + 1103, _
                  "modEvaluationProcessor", _
                  "La normativa es obligatoria."
    End If

    '--------------------------------------------------
    ' 2. Obtener normativa
    '--------------------------------------------------

    Set normativa = _
        repository.GetNormativa(normativaID)

    If normativa Is Nothing Then
        Err.Raise vbObjectError + 1104, _
                  "modEvaluationProcessor", _
                  "No se encontró la normativa: " & normativaID
    End If

    '--------------------------------------------------
    ' 3. Validar estructura de entrada
    '--------------------------------------------------

    Set validation = _
        ValidateInputRange(inputRange)

    If Not validation.isValid Then

        Err.Raise vbObjectError + 1105, _
                  "modEvaluationProcessor", _
                  "La entrada no es válida. Razón: " & _
                  CStr(validation.reason)

    End If

    '--------------------------------------------------
    ' 4. Leer datos
    '--------------------------------------------------

    Set inputs = _
        ReadInputs( _
            inputRange, _
            validation.columns)

    '--------------------------------------------------
    ' 5. Crear servicio de evaluación
    '--------------------------------------------------

    Set service = New clsEvaluationService

    ' El repositorio es una dependencia del servicio.
    ' Se inyecta mediante Initialize.
    Call InitializeEvaluationService( _
        service, _
        repository)

    '--------------------------------------------------
    ' 6. Evaluar cada entrada
    '--------------------------------------------------

    Set records = New Collection

    For i = 1 To inputs.Count

        Set resultInput = inputs(i)

        Set record = _
            service.Evaluate( _
                resultInput, _
                normativaID)

        records.Add record

    Next i

    '--------------------------------------------------
    ' 7. Escribir resultados
    '--------------------------------------------------

    WriteResults _
        inputRange, _
        records, _
        normativa, _
        destination

End Sub


Private Sub InitializeEvaluationService( _
    ByVal service As clsEvaluationService, _
    ByVal repository As clsNormativaRepository)

    ' La inicialización Friend de la clase
    ' se mantiene encapsulada en este módulo.
    Call service.Initialize(repository)

End Sub

