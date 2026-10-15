VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEvaluation 
   Caption         =   "Evaluación Normativa"
   ClientHeight    =   4296
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6156
   OleObjectBlob   =   "frmEvaluation.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmEvaluation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mInputRange As Range
Private mRepository As clsNormativaRepository

Private Sub UserForm_Initialize()

    txtInputRange.value = vbNullString
    txtDestination.value = vbNullString

    Set mRepository = New clsNormativaRepository
    mRepository.Initialize ThisWorkbook

    LoadActiveNormatives

End Sub

Private Sub btnSelectRange_Click()

    Dim selectedRange As Range
    Dim destination As Range

    On Error GoTo SelectionError

    Me.Hide

    On Error Resume Next

    Set selectedRange = Application.InputBox( _
        Prompt:="Seleccione el rango de datos a evaluar.", _
        Title:="Rango de datos", _
        Type:=8)

    On Error GoTo SelectionError

    If selectedRange Is Nothing Then

        Me.Show
        Exit Sub

    End If

    Set mInputRange = selectedRange

    txtInputRange.value = _
        mInputRange.Address(False, False)

    Set destination = _
        GetEvaluationDestination(mInputRange)

    ValidateEvaluationDestination destination

    txtDestination.value = _
        destination.Address(False, False)

    Me.Show

    Exit Sub

SelectionError:

    Me.Show

    MsgBox _
        "No fue posible seleccionar el rango." & _
        vbCrLf & vbCrLf & _
        Err.Description, _
        vbExclamation, _
        "Evaluación Normativa"

End Sub

Private Sub LoadActiveNormatives()

    Dim normativas As Collection
    Dim normativa As clsNormative
    Dim i As Long

    Set normativas = mRepository.GetNormativas

    cboNormativa.Clear

    cboNormativa.ColumnCount = 2
    cboNormativa.BoundColumn = 1
    cboNormativa.ColumnWidths = "150 pt;0 pt"

    For i = 1 To normativas.Count

        Set normativa = normativas(i)

        If normativa.Active Then

            cboNormativa.AddItem normativa.Name

            cboNormativa.List( _
                cboNormativa.ListCount - 1, _
                1) = normativa.ID

        End If

    Next i

End Sub

Private Function GetSelectedNormativaID() As String

    If cboNormativa.ListIndex < 0 Then
        GetSelectedNormativaID = vbNullString
        Exit Function
    End If

    GetSelectedNormativaID = _
        CStr(cboNormativa.Column(1, cboNormativa.ListIndex))

End Function

Private Sub btnEvaluate_Click()

    Dim normativaID As String
    Dim destination As Range

    If mInputRange Is Nothing Then

        MsgBox _
            "Debe seleccionar un rango de datos.", _
            vbExclamation, _
            "Evaluación Normativa"

        Exit Sub

    End If

    normativaID = GetSelectedNormativaID()

    If Len(Trim$(normativaID)) = 0 Then

        MsgBox _
            "Debe seleccionar una normativa.", _
            vbExclamation, _
            "Evaluación Normativa"

        Exit Sub

    End If

    Set destination = _
        GetEvaluationDestination(mInputRange)

    ValidateEvaluationDestination destination

    On Error GoTo EvaluationError

    ProcessEvaluation _
        mInputRange, _
        normativaID, _
        mRepository, _
        destination

    MsgBox _
        "La evaluación normativa se completó correctamente.", _
        vbInformation, _
        "Evaluación Normativa"

    Unload Me

    Exit Sub

EvaluationError:

    MsgBox _
        "No fue posible completar la evaluación." & _
        vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "Evaluación Normativa"

End Sub

Private Sub btnCancel_Click()

    Unload Me

End Sub

' Private Sub cboNormativa_Change()

'    Debug.Print "Normativa seleccionada: " & _
                cboNormativa.value

'    Debug.Print "ID: " & _
                GetSelectedNormativaID

'End Sub
