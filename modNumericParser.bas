Attribute VB_Name = "modNumericParser"
Option Explicit

Public Function TryParseNumericValue( _
    ByVal rawValue As String, _
    ByRef numericValue As Double) _
    As Boolean

    Dim text As String
    Dim normalizedText As String

    text = Trim$(rawValue)

    If Len(text) = 0 Then
        Exit Function
    End If

    If Not IsValidNumericFormat(text) Then
        Exit Function
    End If

    normalizedText = Replace(text, ",", ".")

    numericValue = Val(normalizedText)

    TryParseNumericValue = True

End Function

Private Function IsValidNumericFormat( _
    ByVal text As String) _
    As Boolean

    Dim i As Long
    Dim currentChar As String

    Dim separatorCount As Long
    Dim digitCount As Long
    Dim digitsBeforeSeparator As Long
    Dim digitsAfterSeparator As Long

    For i = 1 To Len(text)

        currentChar = mID$(text, i, 1)

        Select Case currentChar

            Case "0" To "9"

                digitCount = digitCount + 1

                If separatorCount = 0 Then
                    digitsBeforeSeparator = _
                        digitsBeforeSeparator + 1
                Else
                    digitsAfterSeparator = _
                        digitsAfterSeparator + 1
                End If

            Case ".", ","

                separatorCount = separatorCount + 1

                If separatorCount > 1 Then
                    Exit Function
                End If

                If i = 1 Then
                    Exit Function
                End If

                If i = Len(text) Then
                    Exit Function
                End If

            Case "+", "-"

                If i <> 1 Then
                    Exit Function
                End If

            Case Else

                Exit Function

        End Select

    Next i

    If digitCount = 0 Then
        Exit Function
    End If

    If separatorCount = 1 Then

        If digitsBeforeSeparator = 0 Then
            Exit Function
        End If

        If digitsAfterSeparator = 0 Then
            Exit Function
        End If

    End If

    IsValidNumericFormat = True

End Function

