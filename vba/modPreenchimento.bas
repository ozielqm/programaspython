Attribute VB_Name = "modPreenchimento"
Option Explicit

' =====================================================================
' modPreenchimento ? monta a linha do tempo do aluno (BuildTimeline) e
' preenche o MODELO.xlsx (FillWorkbook). Espelha model.go do .exe,
' mesmas coordenadas de c?lula, mesmas regras. S? cobre EF_Concluinte e
' EF_Transferencia (Ensino Fundamental regular) ? n?o as variantes
' Noturno/EJA/PEI/ETI/Ind?gena.
' =====================================================================

Private Const tCadNomeVal = "B14"
Private Const tCadRaVal = "M14"
Private Const tCadMunVal = "E15"
Private Const tCadUfVal = "I15"
Private Const tCadPaisVal = "M15"
Private Const tCadDataVal = "E16"
Private Const tAnoYearRow = 20
Private Const tEstudoStartRow = 52
Private Const tDeclaracaoCell = "A69"
Private Const tDataEmissaoCell = "A71"
Private Const tSerieAtualCell = "G80"
Private Const tBimGradeStartRow = 84
Private Const tBimDiverStartRow = 95
Private Const tDeclBaixoCell = "A79"
Private Const tDataEmissao2Cell = "A115"

Private Const cCadNomeVal = "B17"
Private Const cCadRaVal = "M17"
Private Const cCadMunVal = "E18"
Private Const cCadUfVal = "I18"
Private Const cCadPaisVal = "M18"
Private Const cCadDataVal = "E19"
Private Const cAnoYearRow = 23
Private Const cEstudoStartRow = 55
Private Const cCertificadoCell = "A72"
Private Const cDataEmissaoCell = "A75"
' cCargaHorariaRow (53) nao e mais usado: a carga horaria agora e
' localizada pelo rotulo da linha, em modAutoPreenche.

Private Const sheetTransferencia = "EF_Transferencia"
Private Const sheetConcluinte = "EF_Concluinte"

' ---------- tabelas de disciplinas (chave -> linha na aba TRAN / CONC) ----------

Function BnccKeys() As Variant
    BnccKeys = Array("LINGUA PORTUGUESA", "LINGUA INGLESA", "ARTE", "EDUCACAO FISICA", _
                      "MATEMATICA", "ENSINO RELIGIOSO", "CIENCIAS", "HISTORIA", "GEOGRAFIA")
End Function
Function BnccRowsTran() As Variant
    BnccRowsTran = Array(22, 23, 24, 25, 26, 27, 28, 29, 30)
End Function
Function BnccRowsConc() As Variant
    BnccRowsConc = Array(25, 26, 27, 28, 29, 30, 31, 32, 33)
End Function
Function DiverKeys() As Variant
    DiverKeys = Array("LINGUA INGLESA", "PROJETO DE VIDA", "TECNOLOGIA E INOVACAO", _
                       "PROJETO DE CONVIVENCIA", "ELETIVAS")
End Function
Function DiverRowsTran() As Variant
    DiverRowsTran = Array(34, 35, 36, 37, 38)
End Function
Function DiverRowsConc() As Variant
    DiverRowsConc = Array(37, 38, 39, 40, 41)
End Function

' ---------- constru??o da linha do tempo ----------

Function BuildTimelineByRA(ByVal ra As String, ByRef errMsg As String) As clsTimeline
    Dim anoAtual As Long
    anoAtual = LatestYear()
    If anoAtual = 0 Then
        errMsg = "n?o encontrei nenhuma pasta de ano em ""pasta da Lista de alunos ..."""
        Exit Function
    End If
    Dim mat As clsMatriculaRow, turmaAtual As String
    If Not FindStudentInYearByRA(anoAtual, ra, mat, turmaAtual) Then
        errMsg = "RA """ & ra & """ n?o encontrado na matr?cula de " & anoAtual
        Exit Function
    End If
    Set BuildTimelineByRA = MontarTimeline(mat, turmaAtual, anoAtual, "RA " & ra, True, ra, "")
End Function

Function BuildTimelineByName(ByVal nome As String, ByRef errMsg As String) As clsTimeline
    Dim anoAtual As Long
    anoAtual = LatestYear()
    If anoAtual = 0 Then
        errMsg = "n?o encontrei nenhuma pasta de ano em ""pasta da Lista de alunos ..."""
        Exit Function
    End If
    Dim mat As clsMatriculaRow, turmaAtual As String
    If Not FindStudentInYearByName(anoAtual, nome, mat, turmaAtual) Then
        errMsg = "aluno """ & nome & """ n?o encontrado na matr?cula de " & anoAtual
        Exit Function
    End If
    Set BuildTimelineByName = MontarTimeline(mat, turmaAtual, anoAtual, nome, False, "", nome)
End Function

' Como BuildTimelineByRA, mas ancorado num ano especifico em vez do ano
' mais recente da pasta - usado pelo Formandos, onde o aluno concluinte ja
' nao esta mais matriculado nesta escola no ano corrente (foi pro 6o ano
' em outra escola): "ano atual" pra ele e o ano em que fez o 5o ano aqui,
' nao o ultimo ano de pasta encontrado.
Function BuildTimelineByRAForYear(ByVal ra As String, ByVal ano As Long, ByRef errMsg As String) As clsTimeline
    Dim mat As clsMatriculaRow, turmaAtual As String
    If Not FindStudentInYearByRA(ano, ra, mat, turmaAtual) Then
        errMsg = "RA """ & ra & """ nao encontrado na matricula de " & ano
        Exit Function
    End If
    Set BuildTimelineByRAForYear = MontarTimeline(mat, turmaAtual, ano, "RA " & ra, True, ra, "")
End Function

' N?cleo compartilhado: j? temos a matr?cula do ano mais recente
' (mat/turmaAtual); monta o resto do ciclo buscando ano a ano, por RA ou
' por nome conforme porRA.
Private Function MontarTimeline(ByVal mat As clsMatriculaRow, ByVal turmaAtual As String, ByVal anoAtual As Long, _
                                  ByVal label As String, ByVal porRA As Boolean, ByVal ra As String, ByVal nome As String) As clsTimeline
    Dim tl As New clsTimeline
    tl.nome = mat.nome
    tl.anoAtual = anoAtual
    tl.SerieAtual = mat.Serie
    tl.turmaAtual = turmaAtual
    Set tl.Cadastro = mat

    If mat.Serie < 1 Or mat.Serie > 5 Then
        tl.warnings.Add "s?rie atual (" & mat.Serie & ") fora do intervalo 1-5 esperado para o ciclo ? confira manualmente"
    End If

    Dim anoInicio As Long
    anoInicio = anoAtual - (mat.Serie - 1)
    Dim ano As Long
    Dim yr As clsYearRecord
    For ano = anoInicio To anoAtual
        Set yr = New clsYearRecord
        yr.ano = ano
        yr.Position = ano - anoInicio + 1
        If ano = anoAtual Then
            yr.turma = turmaAtual
            Set yr.mat = mat
        Else
            Dim m As clsMatriculaRow, turma As String
            Dim achou As Boolean
            If porRA Then
                achou = FindStudentInYearByRA(ano, ra, m, turma)
            Else
                achou = FindStudentInYearByName(ano, nome, m, turma)
            End If
            If achou Then
                yr.turma = turma
                Set yr.mat = m
            Else
                yr.warnings.Add "n?o encontrei " & label & " na matr?cula de " & ano & " em nenhuma turma"
            End If
        End If
        If yr.turma <> "" Then
            Dim fechamentos As Collection
            Set fechamentos = LoadNotas(ano, yr.turma)
            Dim nf As clsNotasFechamento
            For Each nf In fechamentos
                If nf.isFinal Then
                    Set yr.FinalNotas = nf
                ElseIf nf.bimestre >= 1 And nf.bimestre <= 4 Then
                    Set yr.Bimestres(CStr(nf.bimestre)) = nf
                End If
            Next nf
        End If
        tl.years.Add yr
        Dim w As Variant
        For Each w In yr.warnings
            tl.warnings.Add w
        Next w
    Next ano
    Set MontarTimeline = tl
End Function

Function GradeValueFor(ByVal nf As clsNotasFechamento, ByVal nome As String, ByVal key As String) As String
    GradeValueFor = ""
    If nf Is Nothing Then Exit Function
    Dim nn As String
    nn = NormName(nome)
    If Not nf.PorAluno.Exists(nn) Then Exit Function
    Dim grades As Collection
    Set grades = nf.PorAluno(nn)
    Dim g As clsSubjectGrade
    For Each g In grades
        If g.Subject = key Then
            GradeValueFor = g.Value
            Exit Function
        End If
    Next g
End Function

' ---------- preenchimento do modelo ----------

Sub FillCadastro(ByVal ws As Worksheet, ByVal tl As clsTimeline, ByVal nat As clsNaturalidade, ByVal isConc As Boolean)
    Dim nomeCell As String, raCell As String, munCell As String, ufCell As String, paisCell As String, dataCell As String
    If isConc Then
        nomeCell = cCadNomeVal: raCell = cCadRaVal: munCell = cCadMunVal
        ufCell = cCadUfVal: paisCell = cCadPaisVal: dataCell = cCadDataVal
    Else
        nomeCell = tCadNomeVal: raCell = tCadRaVal: munCell = tCadMunVal
        ufCell = tCadUfVal: paisCell = tCadPaisVal: dataCell = tCadDataVal
    End If
    Dim ra As String
    ra = tl.Cadastro.ra
    If tl.Cadastro.DigRA <> "" Then ra = ra & "-" & tl.Cadastro.DigRA
    If tl.Cadastro.UFRA <> "" Then ra = ra & " " & tl.Cadastro.UFRA

    ws.Range(nomeCell).Value = tl.nome
    ' RG deliberadamente n?o preenchido (nenhuma planilha de origem traz esse dado)
    ws.Range(raCell).Value = ra
    ws.Range(munCell).Value = nat.Municipio
    ws.Range(ufCell).Value = nat.Estado
    Dim pais As String
    pais = nat.pais
    If pais = "" And nat.Municipio <> "" Then pais = "BRASIL"
    ws.Range(paisCell).Value = pais
    ws.Range(dataCell).Value = tl.Cadastro.Nascimento
End Sub

Sub FillGradeTables(ByVal ws As Worksheet, ByVal tl As clsTimeline, ByVal isConc As Boolean, ByVal warnings As Collection)
    Dim yearRow As Long
    yearRow = IIf(isConc, cAnoYearRow, tAnoYearRow)
    Dim numAnosFinais As Integer
    If isConc Then
        numAnosFinais = tl.SerieAtual
    Else
        numAnosFinais = tl.SerieAtual - 1
    End If
    ' serie atual invalida (0 ou fora do range) -> nao ha colunas de ano pra
    ' preencher; so avisa (ja acontece em MontarTimeline) e segue sem travar.
    If numAnosFinais < 0 Then numAnosFinais = 0
    If numAnosFinais > 5 Then numAnosFinais = 5
    Dim anoCols As Variant
    anoCols = Array("G", "H", "I", "J", "K")

    Dim bnccKeysArr As Variant, bnccRows As Variant, diverKeysArr As Variant, diverRows As Variant
    bnccKeysArr = BnccKeys()
    diverKeysArr = DiverKeys()
    If isConc Then
        bnccRows = BnccRowsConc()
        diverRows = DiverRowsConc()
    Else
        bnccRows = BnccRowsTran()
        diverRows = DiverRowsTran()
    End If

    Dim i As Integer
    For i = 0 To numAnosFinais - 1
        If i > tl.years.Count - 1 Then Exit For
        Dim yr As clsYearRecord
        Set yr = tl.years(i + 1)
        Dim col As String
        col = anoCols(i)
        ws.Range(col & yearRow).Value = yr.ano

        Dim nomeAno As String
        nomeAno = tl.nome
        If Not yr.mat Is Nothing Then
            If yr.mat.nome <> "" Then nomeAno = yr.mat.nome
        End If

        Dim temManual As Boolean
        temManual = (yr.FinalNotas Is Nothing) And (Not yr.ManualNotas Is Nothing) And (yr.ManualNotas.Count > 0)

        Dim k As Integer
        For k = LBound(bnccKeysArr) To UBound(bnccKeysArr)
            Dim v As String
            v = ValorNotaAno(yr, nomeAno, bnccKeysArr(k), CStr(bnccRows(k)), temManual)
            If v <> "" Then ws.Range(col & bnccRows(k)).Value = v
        Next k
        For k = LBound(diverKeysArr) To UBound(diverKeysArr)
            Dim v2 As String
            v2 = ValorNotaAno(yr, nomeAno, diverKeysArr(k), CStr(diverRows(k)), temManual)
            If v2 <> "" Then ws.Range(col & diverRows(k)).Value = v2
        Next k

        If isConc Then
            ' carga horaria digitada -> linhas de TOTAL DE AULAS (BNCC e
            ' Parte Diversificada); a linha TOTAL DA CARGA HORARIA ANUAL
            ' (Horas) continua sendo formula, com o fator do ano. Ver
            ' modAutoPreenche.
            AplicarCargaHoraria ws, col, yr.CargaHoraria, warnings
            AplicarHorasAnuais ws, col, yr.ano, warnings
        End If

        If yr.FinalNotas Is Nothing And Not temManual Then
            warnings.Add "ano " & yr.ano & " (turma " & yr.turma & "): não encontrei o Mapão de fechamento final — notas desse ano ficaram em branco"
        ElseIf temManual Then
            warnings.Add "ano " & yr.ano & " (turma " & yr.turma & "): notas digitadas manualmente (Mapão não encontrado) — confira antes de imprimir"
        End If
    Next i
End Sub

' Nota final de uma disciplina num ano: usa o Map?o quando achou; sen?o,
' cai pra nota digitada ? m?o (ManualNotas), guardada pela chave de linha
' (bnccRows/diverRows) em vez do texto da disciplina, porque "LINGUA
' INGLESA" aparece tanto na BNCC quanto na parte diversificada com a
' mesma chave de texto.
Private Function ValorNotaAno(ByVal yr As clsYearRecord, ByVal nomeAno As String, ByVal key As String, _
                               ByVal manualKey As String, ByVal temManual As Boolean) As String
    If Not yr.FinalNotas Is Nothing Then
        ValorNotaAno = GradeValueFor(yr.FinalNotas, nomeAno, key)
    ElseIf temManual And yr.ManualNotas.Exists(manualKey) Then
        ValorNotaAno = yr.ManualNotas(manualKey)
    Else
        ValorNotaAno = ""
    End If
End Function

Sub FillEstudosRealizados(ByVal ws As Worksheet, ByVal tl As clsTimeline, ByVal isConc As Boolean)
    Dim start As Long
    start = IIf(isConc, cEstudoStartRow, tEstudoStartRow)
    Dim limit As Integer
    limit = tl.SerieAtual
    If limit < 0 Then limit = 0
    If limit > 5 Then limit = 5
    Dim i As Integer
    For i = 0 To limit - 1
        If i > tl.years.Count - 1 Then Exit For
        Dim yr As clsYearRecord
        Set yr = tl.years(i + 1)
        Dim r As Long
        r = start + i
        Dim estab As String
        estab = EscolaNome
        If Not isConc And i = limit - 1 Then estab = EscolaNome & "(TRANSF)"
        ws.Range("F" & r).Value = yr.ano
        ws.Range("G" & r).Value = estab
        ws.Range("L" & r).Value = "S?O CARLOS"
        ws.Range("O" & r).Value = "SP"
    Next i
End Sub

Sub FillDeclaracaoTransferencia(ByVal ws As Worksheet, ByVal tl As clsTimeline, ByVal dataEmissao As String, ByVal warnings As Collection)
    Dim ra As String
    ra = tl.Cadastro.ra
    If tl.Cadastro.DigRA <> "" Then ra = ra & "-" & tl.Cadastro.DigRA
    Dim serieTxt As String
    serieTxt = CStr(tl.SerieAtual) & "?"
    Dim texto As String
    texto = "O Diretor da " & EscolaNome & ", DECLARA, nos termos do Inciso VII, Artigo 24 da Lei Federal 9394/96, que " & _
            tl.nome & ", RA " & ra & ", est? cursando o " & serieTxt & " ano do Ensino Fundamental no ano de " & tl.anoAtual & "."
    Dim textoBaixo As String
    textoBaixo = "DECLARO, nos termos do Inciso VII, artigo 24, da Lei 9394/96, que o aluno " & tl.nome & _
                 ", RA n? " & ra & ", encontra-se matriculado na " & serieTxt & " S?rie/Ano, do Ensino Fundamental, " & _
                 "no presente ano letivo, e apresenta aproveitamento e frequ?ncia descritos abaixo, at? " & dataEmissao & "."

    ws.Range(tDeclaracaoCell).Value = texto
    ws.Range(tDataEmissaoCell).Value = "Data: " & dataEmissao
    ws.Range(tSerieAtualCell).Value = serieTxt & " S?rie / Ano"
    ws.Range(tDeclBaixoCell).Value = textoBaixo
    ws.Range(tDataEmissao2Cell).Value = "Data: " & dataEmissao
    warnings.Add "IMPORTANTE: revise manualmente o par?grafo da declara??o (c?lula A69/A79) antes de imprimir ? ? texto gerado automaticamente."
End Sub

Sub FillCertificadoConcluinte(ByVal ws As Worksheet, ByVal tl As clsTimeline, ByVal dataEmissao As String, ByVal warnings As Collection)
    Dim ra As String
    ra = tl.Cadastro.ra
    If tl.Cadastro.DigRA <> "" Then ra = ra & "-" & tl.Cadastro.DigRA
    Dim texto As String
    texto = "O Diretor da " & EscolaNome & ", CERTIFICA, nos termos do Inciso VII, Artigo 24 da Lei Federal 9394/96, que " & _
            tl.nome & ", RA " & ra & ", concluiu o Ensino Fundamental, no ano de " & tl.anoAtual & "."
    ws.Range(cCertificadoCell).Value = texto
    ws.Range(cDataEmissaoCell).Value = "Data: " & dataEmissao
    warnings.Add "IMPORTANTE: revise manualmente o par?grafo do certificado (c?lula A72) antes de imprimir ? ? texto gerado automaticamente."
End Sub

Sub FillBimestres(ByVal ws As Worksheet, ByVal tl As clsTimeline, ByVal warnings As Collection)
    If tl.years.Count = 0 Then Exit Sub
    Dim cur As clsYearRecord
    Set cur = tl.years(tl.years.Count)
    Dim nomeAno As String
    nomeAno = tl.nome
    If Not cur.mat Is Nothing Then
        If cur.mat.nome <> "" Then nomeAno = cur.mat.nome
    End If
    Dim bimCols As Variant
    bimCols = Array("G", "H", "I", "J")
    Dim bnccKeysArr As Variant, bnccRowsTranArr As Variant, diverKeysArr As Variant, diverRowsTranArr As Variant
    bnccKeysArr = BnccKeys()
    bnccRowsTranArr = BnccRowsTran()
    diverKeysArr = DiverKeys()
    diverRowsTranArr = DiverRowsTran()

    Dim anyFound As Boolean
    anyFound = False
    Dim bim As Integer
    For bim = 1 To 4
        If cur.Bimestres.Exists(CStr(bim)) Then
            anyFound = True
            Dim nf As clsNotasFechamento
            Set nf = cur.Bimestres(CStr(bim))
            Dim col As String
            col = bimCols(bim - 1)
            Dim k As Integer
            For k = LBound(bnccKeysArr) To UBound(bnccKeysArr)
                Dim v As String
                v = GradeValueFor(nf, nomeAno, bnccKeysArr(k))
                If v <> "" Then
                    Dim rowIdx As Long
                    rowIdx = tBimGradeStartRow + (bnccRowsTranArr(k) - 22)
                    ws.Range(col & rowIdx).Value = v
                End If
            Next k
            For k = LBound(diverKeysArr) To UBound(diverKeysArr)
                Dim v2 As String
                v2 = GradeValueFor(nf, nomeAno, diverKeysArr(k))
                If v2 <> "" Then
                    Dim rowIdx2 As Long
                    rowIdx2 = tBimDiverStartRow + (diverRowsTranArr(k) - 34)
                    ws.Range(col & rowIdx2).Value = v2
                End If
            Next k
        End If
    Next bim
    If Not anyFound Then
        warnings.Add "n?o encontrei nenhum Map?o de bimestre para o ano corrente (" & cur.ano & ", turma " & cur.turma & ") ? a se??o de notas do ano em curso ficou em branco"
    End If
End Sub

' Abre o MODELO.xlsx, salva j? como o arquivo de sa?da (nunca sobrescreve
' o modelo original), preenche a aba certa, salva e fecha.
Function FillWorkbookToFile(ByVal modeloPath As String, ByVal outPath As String, ByVal tl As clsTimeline, _
                             ByVal nat As clsNaturalidade, ByVal modo As String, ByVal dataEmissao As String, _
                             ByVal warnings As Collection, ByRef errDetail As String) As Boolean
    On Error GoTo falhou
    Dim wbOut As Workbook
    Set wbOut = Workbooks.Open(Filename:=modeloPath, ReadOnly:=False, UpdateLinks:=0, AddToMru:=False, Notify:=False)
    Application.DisplayAlerts = False
    wbOut.SaveAs Filename:=outPath, FileFormat:=51 ' xlOpenXMLWorkbook (.xlsx)
    Application.DisplayAlerts = True

    Dim ws As Worksheet
    If modo = "TRAN" Then
        Set ws = wbOut.Worksheets(sheetTransferencia)
        FillCadastro ws, tl, nat, False
        FillGradeTables ws, tl, False, warnings
        FillEstudosRealizados ws, tl, False
        FillDeclaracaoTransferencia ws, tl, dataEmissao, warnings
        FillBimestres ws, tl, warnings
    Else
        Set ws = wbOut.Worksheets(sheetConcluinte)
        FillCadastro ws, tl, nat, True
        FillGradeTables ws, tl, True, warnings
        FillEstudosRealizados ws, tl, True
        FillCertificadoConcluinte ws, tl, dataEmissao, warnings
    End If

    wbOut.Save
    wbOut.Close SaveChanges:=False
    FillWorkbookToFile = True
    Exit Function

falhou:
    errDetail = "Erro " & Err.Number & ": " & Err.Description
    FillWorkbookToFile = False
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not wbOut Is Nothing Then wbOut.Close SaveChanges:=False
    On Error GoTo 0
End Function

' Acha, dentro de BaseDir(), a subpasta cujo nome começa com prefixo —
' tolera um sufixo tipo "desde 2020" ou "em 2026" que as pastas de saída
' reais dessa escola têm, em vez do nome literal exato. Se nada bater,
' devolve BaseDir() & "\" & fallback (que o EnsureFolder cria na hora).
Function FindOutDirByPrefix(ByVal prefixo As String, ByVal fallback As String) As String
    Dim d As String
    On Error Resume Next
    d = Dir(BaseDir() & "\", vbDirectory)
    On Error GoTo 0
    Do While d <> ""
        If d <> "." And d <> ".." Then
            If (GetAttr(BaseDir() & "\" & d) And vbDirectory) = vbDirectory Then
                If Left$(d, Len(prefixo)) = prefixo Then
                    FindOutDirByPrefix = BaseDir() & "\" & d
                    Exit Function
                End If
            End If
        End If
        d = Dir()
    Loop
    FindOutDirByPrefix = BaseDir() & "\" & fallback
End Function

Function FindConcOutDir() As String
    FindConcOutDir = FindOutDirByPrefix("pasta onde serão guardados os históricos gerados dos", _
                                         "pasta onde serão guardados os históricos gerados dos concluintes")
End Function

Function FindTranOutDir() As String
    FindTranOutDir = FindOutDirByPrefix("pasta onde serão guardados os históricos de transferência gerados", _
                                         "pasta onde serão guardados os históricos de transferência gerados")
End Function

' ---------- orquestra??o de um aluno (equivalente a processStudent no .exe) ----------

' tb ? passado como Object (late binding) de prop?sito ? assim este
' m?dulo n?o depende de a refer?ncia "Microsoft Forms 2.0 Object
' Library" j? estar presente no projeto no momento em que este c?digo ?
' importado (ela s? ? adicionada automaticamente quando voc? insere o
' UserForm, no Passo 4 do guia de instala??o).
Sub AppendLog(ByVal tb As Object, ByVal msg As String)
    tb.Value = tb.Value & msg & vbCrLf
    tb.SelStart = Len(tb.Value)
    DoEvents
End Sub

Sub ProcessStudent(ByVal ra As String, ByVal nomeFallback As String, ByVal modo As String, _
                    ByVal natOverride As clsNaturalidade, ByVal nascimentoOverride As String, _
                    ByVal outDirTran As String, ByVal outDirConc As String, ByVal dataEmissao As String, _
                    ByVal modeloPath As String, ByVal tb As Object)
    Dim label As String
    label = ra
    If Trim$(label) = "" Then label = nomeFallback
    AppendLog tb, "-----------------------------------------"
    AppendLog tb, "Aluno informado: " & label

    Dim tl As clsTimeline
    Dim errMsg As String, errMsg2 As String
    Dim usouFallback As Boolean
    usouFallback = False

    If Trim$(ra) <> "" Then
        Set tl = BuildTimelineByRA(ra, errMsg)
    End If
    If tl Is Nothing And Trim$(nomeFallback) <> "" Then
        Dim tl2 As clsTimeline
        Set tl2 = BuildTimelineByName(nomeFallback, errMsg2)
        If Not tl2 Is Nothing Then
            Set tl = tl2
            If Trim$(ra) <> "" Then usouFallback = True
            errMsg = ""
        ElseIf errMsg = "" Then
            errMsg = errMsg2
        End If
    End If

    If tl Is Nothing Then
        AppendLog tb, "  ERRO: " & errMsg
        Exit Sub
    End If
    If usouFallback Then
        AppendLog tb, "  AVISO: n?o encontrei esse RA na matr?cula mais recente ? localizei o aluno pelo nome como alternativa. Confira o RA na planilha de matr?cula."
    End If
    AppendLog tb, "  Encontrado: " & tl.nome & " (RA " & tl.Cadastro.ra & ")"

    If modo = "CONC" And tl.SerieAtual <> 5 Then
        AppendLog tb, "  AVISO: aluno est? no " & tl.SerieAtual & "? ano, n?o no 5? ? confirme se HIST?RICO_CONC ? mesmo o modo certo."
    End If
    If modo = "TRAN" And InStr(1, UCase$(tl.Cadastro.Situacao), "TRAN") = 0 Then
        AppendLog tb, "  AVISO: situa??o da matr?cula ? """ & tl.Cadastro.Situacao & """ (n?o cont?m ""TRAN"") ? confirme se ? mesmo um caso de transfer?ncia."
    End If

    Dim nat As clsNaturalidade
    If Not natOverride Is Nothing Then
        Set nat = natOverride
    Else
        Set nat = New clsNaturalidade
        AppendLog tb, "  AVISO: naturalidade n?o informada ? campos de munic?pio/estado/pa?s de nascimento ficar?o em branco (preencha manualmente depois)."
    End If
    If nascimentoOverride <> "" Then tl.Cadastro.Nascimento = nascimentoOverride

    Dim outDir As String, sufixo As String
    If modo = "TRAN" Then
        outDir = outDirTran: sufixo = "TRANSFERENCIA"
    Else
        outDir = outDirConc: sufixo = "CONCLUINTE"
    End If
    EnsureFolder outDir
    Dim outPath As String
    outPath = outDir & "\" & SanitizeFilename(tl.nome) & " - " & sufixo & ".xlsx"

    Dim warnings As New Collection
    Dim ok As Boolean, errDetail As String
    ok = FillWorkbookToFile(modeloPath, outPath, tl, nat, modo, dataEmissao, warnings, errDetail)

    Dim w As Variant
    For Each w In tl.warnings
        AppendLog tb, "  AVISO: " & w
    Next w
    For Each w In warnings
        AppendLog tb, "  AVISO: " & w
    Next w

    If Not ok Then
        AppendLog tb, "  ERRO salvando o arquivo de sa?da" & IIf(errDetail <> "", ": " & errDetail, ".")
        Exit Sub
    End If
    AppendLog tb, "  OK -> " & outPath
End Sub




