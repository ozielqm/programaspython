Attribute VB_Name = "frmHistorico"
Attribute VB_Base = "0{B69D746B-BB09-4A2E-95A7-3C8E3D57CDF6}{3AFC3E94-C063-4A42-A58D-7D3DD0B97814}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private gTl As clsTimeline


Private Sub Label1_Click()

End Sub

Private Sub txtLog_Change()

End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "HISTORICO DE CONCLUINTES (5 ANO)"
    lblStatus.Caption = "Pasta de trabalho: " & BaseDir()
    txtLog.Value = ""
    lblAlunoEncontrado.Caption = ""
    cmdGerar.Enabled = False
    txtAnoConclusao.Value = CStr(LatestYear() - 1)
    LimparTabelaAnos
End Sub

Private Sub LimparTabelaAnos()
    Dim i As Integer
    For i = 1 To 5
        Me.Controls("lblAno" & i).Caption = i & "o ano:"
        Me.Controls("txtAno" & i).Value = ""
        Me.Controls("lblSit" & i).Caption = ""
        Me.Controls("txtEmailAno" & i).Visible = False
        Me.Controls("txtEmailAno" & i).Value = ""
        Me.Controls("cmdPedirAno" & i).Visible = False
        Me.Controls("cmdNotasAno" & i).Visible = False
        Me.Controls("txtCargaAno" & i).Value = ""
    Next i
End Sub

Private Sub AtualizarStatusLinha(ByVal i As Integer)
    Dim lblSit As Object, txtEmail As Object, cmdPedir As Object, cmdNotas As Object
    Set lblSit = Me.Controls("lblSit" & i)
    Set txtEmail = Me.Controls("txtEmailAno" & i)
    Set cmdPedir = Me.Controls("cmdPedirAno" & i)
    Set cmdNotas = Me.Controls("cmdNotasAno" & i)
    Dim yr As clsYearRecord
    Set yr = gTl.years(i)
    Dim temManual As Boolean
    temManual = (Not yr.ManualNotas Is Nothing) And (yr.ManualNotas.Count > 0)
    Dim faltando As Boolean
    faltando = (yr.mat Is Nothing) Or (yr.FinalNotas Is Nothing And Not temManual)
    If faltando Then
        lblSit.Caption = "FALTANDO"
        txtEmail.Visible = True
        cmdPedir.Visible = True
    Else
        If temManual And yr.FinalNotas Is Nothing Then
            lblSit.Caption = "OK (notas manuais) - turma " & yr.turma
        Else
            lblSit.Caption = "OK - turma " & yr.turma
        End If
        txtEmail.Visible = False
        cmdPedir.Visible = False
    End If
    cmdNotas.Visible = (yr.FinalNotas Is Nothing)
End Sub

Private Sub AtualizarTabelaAnos()
    Dim i As Integer
    For i = 1 To 5
        Dim txtAno As Object, lblSit As Object, txtEmail As Object, cmdPedir As Object
        Set txtAno = Me.Controls("txtAno" & i)
        Set lblSit = Me.Controls("lblSit" & i)
        Set txtEmail = Me.Controls("txtEmailAno" & i)
        Set cmdPedir = Me.Controls("cmdPedirAno" & i)
        If i <= gTl.years.Count Then
            Dim yr As clsYearRecord
            Set yr = gTl.years(i)
            txtAno.Value = CStr(yr.ano)
            AtualizarStatusLinha i
        Else
            txtAno.Value = ""
            lblSit.Caption = "-"
            txtEmail.Visible = False
            cmdPedir.Visible = False
            Me.Controls("cmdNotasAno" & i).Visible = False
        End If
    Next i
End Sub

Private Sub cmdBuscarRA_Click()
    Dim ra As String
    ra = Trim$(txtBuscaRA.Value)
    If ra = "" Then
        MsgBox "Digite o RA do aluno.", vbExclamation
        Exit Sub
    End If
    Dim anoConclusao As Long
    anoConclusao = CLng(val(txtAnoConclusao.Value))
    If anoConclusao < 2000 Then
        MsgBox "Digite um ano de conclusao valido (ex: 2025).", vbExclamation
        Exit Sub
    End If
    ResetCaches
    Dim errMsg As String
    Dim tl As clsTimeline
    Set tl = BuildTimelineByRAForYear(ra, anoConclusao, errMsg)
    If tl Is Nothing Then
        Set gTl = Nothing
        lblAlunoEncontrado.Caption = ""
        cmdGerar.Enabled = False
        LimparTabelaAnos
        MsgBox errMsg, vbExclamation
        Exit Sub
    End If
    Set gTl = tl
    lblAlunoEncontrado.Caption = "Aluno: " & tl.nome & " (" & tl.SerieAtual & "o ano em " & anoConclusao & ", turma " & tl.turmaAtual & ")"
    txtNascimento.Value = tl.Cadastro.Nascimento
    Dim natAuto As clsNaturalidade
    Dim origemNat As String
    Set natAuto = BuscarNaturalidadeAuto(tl, anoConclusao, origemNat)
    txtMunicipioNasc.Value = natAuto.Municipio
    txtEstadoNasc.Value = natAuto.Estado
    txtPaisNasc.Value = natAuto.pais
    Dim ci As Integer
    For ci = 1 To 5
        Me.Controls("txtCargaAno" & ci).Value = ""
    Next ci
    If tl.SerieAtual <> 5 Then
        AppendLog txtLog, "AVISO: aluno esta registrado como " & tl.SerieAtual & "o ano em " & anoConclusao & ", nao 5o - confirme o ano de conclusao."
    End If
    AtualizarTabelaAnos
    cmdGerar.Enabled = True
    AppendLog txtLog, "Aluno localizado por RA: " & tl.nome & ". Nascimento pre-preenchido da matricula (confira/ajuste se precisar). Confira a tabela de anos (ajuste o ano de qualquer linha se precisar, ex: repetencia, e clique Atualizar) e preencha a carga horaria antes de gerar."
    AppendLog txtLog, origemNat
End Sub

Private Sub AtualizarLinhaAno(ByVal i As Integer)
    If gTl Is Nothing Then Exit Sub
    Dim txtAno As Object
    Set txtAno = Me.Controls("txtAno" & i)
    Dim ano As Long
    ano = CLng(val(txtAno.Value))
    If ano < 2000 Then
        MsgBox "Digite um ano valido.", vbExclamation
        Exit Sub
    End If
    Dim yr As clsYearRecord
    Set yr = New clsYearRecord
    yr.ano = ano
    yr.Position = i
    Dim mat As clsMatriculaRow, turma As String
    If FindStudentInYearByRA(ano, gTl.Cadastro.ra, mat, turma) Then
        yr.turma = turma
        Set yr.mat = mat
        Dim fechamentos As Collection
        Set fechamentos = LoadNotas(ano, turma)
        Dim nf As clsNotasFechamento
        For Each nf In fechamentos
            If nf.isFinal Then Set yr.FinalNotas = nf
        Next nf
    End If
    If i <= gTl.years.Count Then gTl.years.Remove i
    If i <= gTl.years.Count Then
        gTl.years.Add yr, Before:=i
    Else
        gTl.years.Add yr
    End If
    AtualizarStatusLinha i
    If yr.mat Is Nothing Then
        AppendLog txtLog, "Ano " & i & " atualizado para " & ano & ": aluno nao encontrado nesse ano."
    Else
        AppendLog txtLog, "Ano " & i & " atualizado para " & ano & ": turma " & yr.turma & IIf(yr.FinalNotas Is Nothing, " (sem Mapao final)", " (Mapao final OK)")
    End If
End Sub

Private Sub cmdAtualizarAno1_Click()
    AtualizarLinhaAno 1
End Sub

Private Sub cmdAtualizarAno2_Click()
    AtualizarLinhaAno 2
End Sub

Private Sub cmdAtualizarAno3_Click()
    AtualizarLinhaAno 3
End Sub

Private Sub cmdAtualizarAno4_Click()
    AtualizarLinhaAno 4
End Sub

Private Sub cmdAtualizarAno5_Click()
    AtualizarLinhaAno 5
End Sub

Private Sub EnviarPedidoAno(ByVal i As Integer)
    If gTl Is Nothing Then Exit Sub
    Dim txtEmail As Object
    Set txtEmail = Me.Controls("txtEmailAno" & i)
    If Trim$(txtEmail.Value) = "" Then
        MsgBox "Digite o e-mail da escola desse ano.", vbExclamation
        Exit Sub
    End If
    If i > gTl.years.Count Then Exit Sub
    Dim yr As clsYearRecord
    Set yr = gTl.years(i)
    Dim detalhe As String
    detalhe = "- " & i & "o ano (" & yr.ano & ")" & vbCrLf
    Dim errMsg As String
    If EnviarEmailOutlook(txtEmail.Value, MontarAssuntoEscola(gTl.nome), MontarCorpoEscola(gTl.nome, detalhe), errMsg) Then
        AppendLog txtLog, "E-mail pedindo o " & i & "o ano (" & yr.ano & ") aberto no Outlook - revise e clique Enviar."
    Else
        MsgBox errMsg, vbCritical
    End If
End Sub

Private Sub cmdPedirAno1_Click()
    EnviarPedidoAno 1
End Sub

Private Sub cmdPedirAno2_Click()
    EnviarPedidoAno 2
End Sub

Private Sub cmdPedirAno3_Click()
    EnviarPedidoAno 3
End Sub

Private Sub cmdPedirAno4_Click()
    EnviarPedidoAno 4
End Sub

Private Sub cmdPedirAno5_Click()
    EnviarPedidoAno 5
End Sub

Private Sub AbrirNotasManuais(ByVal i As Integer)
    If gTl Is Nothing Then Exit Sub
    If i > gTl.years.Count Then Exit Sub
    Dim yr As clsYearRecord
    Set yr = gTl.years(i)
    frmNotasManuais.Preparar yr, i & "o ano (" & yr.ano & ")"
    frmNotasManuais.Show vbModal
    If frmNotasManuais.Salvo Then
        AppendLog txtLog, "Notas manuais do " & i & "o ano (" & yr.ano & ") atualizadas."
    End If
    Unload frmNotasManuais
    AtualizarStatusLinha i
End Sub

Private Sub cmdNotasAno1_Click()
    AbrirNotasManuais 1
End Sub

Private Sub cmdNotasAno2_Click()
    AbrirNotasManuais 2
End Sub

Private Sub cmdNotasAno3_Click()
    AbrirNotasManuais 3
End Sub

Private Sub cmdNotasAno4_Click()
    AbrirNotasManuais 4
End Sub

Private Sub cmdNotasAno5_Click()
    AbrirNotasManuais 5
End Sub

Private Sub cmdGerar_Click()
    If gTl Is Nothing Then
        MsgBox "Busque um aluno por RA primeiro.", vbExclamation
        Exit Sub
    End If
    cmdGerar.Enabled = False
    Application.ScreenUpdating = False
    If Trim$(BaseDir()) = "" Then
        AppendLog txtLog, "ERRO: esta pasta de trabalho ainda nao foi salva. Va em Arquivo > Salvar Como, escolha ""Pasta de Trabalho Habilitada para Macro (*.xlsm)"" e salve dentro da pasta ""historicos"" antes de clicar em Gerar."
        GoTo fim
    End If

    Dim modeloPath As String
    modeloPath = BaseDir() & "\histórico modelo\MODELO.xlsx"
    If Dir(modeloPath) = "" Then
        AppendLog txtLog, "ERRO: nao encontrei o modelo em " & modeloPath
        GoTo fim
    End If

    Dim nat As New clsNaturalidade
    nat.Municipio = Trim$(txtMunicipioNasc.Value)
    nat.Estado = Trim$(txtEstadoNasc.Value)
    nat.pais = Trim$(txtPaisNasc.Value)
    If Trim$(txtNascimento.Value) <> "" Then gTl.Cadastro.Nascimento = Trim$(txtNascimento.Value)

    Dim ci As Integer
    For ci = 1 To gTl.years.Count
        Dim yrc As clsYearRecord
        Set yrc = gTl.years(ci)
        yrc.CargaHoraria = Trim$(Me.Controls("txtCargaAno" & ci).Value)
    Next ci

    Dim faltouAlgumAno As Boolean
    faltouAlgumAno = False
    Dim yr As clsYearRecord
    For Each yr In gTl.years
        Dim temManualG As Boolean
        temManualG = (Not yr.ManualNotas Is Nothing) And (yr.ManualNotas.Count > 0)
        If yr.mat Is Nothing Or (yr.FinalNotas Is Nothing And Not temManualG) Then faltouAlgumAno = True
    Next yr

    Dim outDir As String
    If faltouAlgumAno Then
        outDir = BaseDir() & "\alunos que faltam completar o Histórico"
    Else
        outDir = FindConcOutDir()
    End If
    EnsureFolder outDir
    Dim outPath As String
    outPath = outDir & "\" & SanitizeFilename(gTl.nome) & " - CONCLUINTE.xlsx"

    Dim warnings As New Collection
    Dim ok As Boolean, errDetail As String
    ok = FillWorkbookToFile(modeloPath, outPath, gTl, nat, "CONC", Format(Date, "dd/mm/yyyy"), warnings, errDetail)

    Dim w As Variant
    For Each w In gTl.warnings
        AppendLog txtLog, "AVISO: " & w
    Next w
    For Each w In warnings
        AppendLog txtLog, "AVISO: " & w
    Next w

    If Not ok Then
        AppendLog txtLog, "ERRO salvando o arquivo de saida" & IIf(errDetail <> "", ": " & errDetail, ".")
    Else
        If faltouAlgumAno Then
            AppendLog txtLog, "OK (INCOMPLETO - falta pelo menos 1 ano, salvo em ""alunos que faltam completar o Historico"") -> " & outPath
        Else
            AppendLog txtLog, "OK -> " & outPath
        End If
    End If

fim:
    Application.ScreenUpdating = True
    cmdGerar.Enabled = True
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub

