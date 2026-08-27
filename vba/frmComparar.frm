Attribute VB_Name = "frmComparar"
Attribute VB_Base = "0{CE5059FC-A612-4E78-8BF7-91FF6E731EAF}{A509EED4-4A91-477C-9B66-65BDF86CF201}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "Comparar alunos por ano"
    txtAnoA.Value = CStr(LatestYear())
    txtAnoB.Value = CStr(LatestYear() - 1)
End Sub

Private Sub CarregarLista(ByVal txtAno As Object, ByVal lst As Object)
    Dim ano As Long
    ano = CLng(val(txtAno.Value))
    If ano < 2000 Then
        MsgBox "Digite um ano valido.", vbExclamation
        Exit Sub
    End If
    lst.Clear
    Dim alunos As Collection
    Set alunos = ListarAlunosDoAno(ano)
    Dim r As clsMatriculaRow
    For Each r In alunos
        lst.AddItem r.nome
        Dim idx As Long
        idx = lst.ListCount - 1
        lst.List(idx, 1) = r.ra
        lst.List(idx, 2) = r.turma
        lst.List(idx, 3) = r.Serie
        lst.List(idx, 4) = r.Situacao
    Next r
    Me.Caption = "Comparar alunos por ano (" & alunos.Count & " alunos carregados)"
End Sub

Private Sub cmdCarregarA_Click()
    CarregarLista txtAnoA, lstAlunosA
End Sub

Private Sub cmdCarregarB_Click()
    CarregarLista txtAnoB, lstAlunosB
End Sub

Private Sub cmdFecharComparar_Click()
    Unload Me
End Sub

