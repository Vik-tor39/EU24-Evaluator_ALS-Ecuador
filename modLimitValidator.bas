Attribute VB_Name = "modLimitValidator"

Option Explicit

'============================================================
' Validación de consistencia de límites normativos
'============================================================

Public Function HasInconsistentLimits( _
    ByVal limits As clsNormativeLimit) _
    As Boolean

    '--------------------------------------------------
    ' 1. Comprobar que exista el objeto
    '--------------------------------------------------

    If limits Is Nothing Then
        HasInconsistentLimits = False
        Exit Function
    End If

    '--------------------------------------------------
    ' 2. Solo se pueden comparar los límites cuando
    '    ambos están definidos.
    '--------------------------------------------------

    If Not limits.HasLIE Or Not limits.HasLSE Then
        HasInconsistentLimits = False
        Exit Function
    End If

    '--------------------------------------------------
    ' 3. Detectar inconsistencia:
    '    el límite inferior supera al superior.
    '--------------------------------------------------

    HasInconsistentLimits = (limits.LIE > limits.LSE)

End Function


