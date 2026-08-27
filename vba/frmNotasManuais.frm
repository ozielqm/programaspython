Attribute VB_Name = "frmNotasManuais"
Attribute VB_Base = "0{E5BF804E-0836-47C0-96F8-A9B4F2C10F3C}{9406A3AC-41C8-40F9-8B8A-17AB5BE7DE0F}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private yr As clsYearRecord
Public Salvo As Boolean

Private Function ChavesNotas() As Variant
    Dim bnccRows As Variant, diverRows As Variant, i As Integer
    bnccRows = BnccRowsConc()
    diverRows = DiverRowsConc()
    Dim result(1 To 14) As String
    For i = LBound(bnccRows) To UBound(bnccRows)
        result(i - LBound(bnccRows) + 1) = CStr(bnccRows(i))
    Next i
    Dim base As Integer
    base = UBound(bnccRows) - LBound(bnccRows) + 1
    For i = LBound(diverRows) To UBound(diverRows)
        result(base + i - LBound(diverRows) + 1) = CStr(diverRows(i))
    Next i
    ChavesNotas = result
End Function

Public Sub Preparar(ByVal alvo As clsYearRecord, ByVal anoLabel As String)
    Set yr = alvo
    Salvo = False
    Me.Caption = "Notas manuais - " & anoLabel
    Dim k As Variant
    k = ChavesNotas()
    Dim i As Integer
    For i = 1 To 14
        Dim txt As Object
        Set txt = Me.Controls("txtNota" & i)
        If yr.ManualNotas.Exists(k(i)) Then
            txt.Value = yr.ManualNotas(k(i))
        Else
            txt.Value = ""
        End If
    Next i
End Sub

Private Sub cmdSalvarNotas_Click()
    Dim k As Variant
    k = ChavesNotas()
    Dim i As Integer
    For i = 1 To 14
        Dim txt As Object
        Set txt = Me.Controls("txtNota" & i)
        Dim v As String
        v = Trim$(txt.Value)
        If v <> "" Then
            yr.ManualNotas(k(i)) = v
        ElseIf yr.ManualNotas.Exists(k(i)) Then
            yr.ManualNotas.Remove k(i)
        End If
    Next i
    Salvo = True
    Me.Hide
End Sub

Private Sub cmdCancelarNotas_Click()
    Salvo = False
    Me.Hide
End Sub

