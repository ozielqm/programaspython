Attribute VB_Name = "modAutoPreenche"
Option Explicit

' =====================================================================
' modAutoPreenche - duas coisas:
'
'   1) Naturalidade automatica: descobre municipio / UF / pais de
'      nascimento do aluno sozinho, sem o usuario digitar. Procura em
'      duas fontes, nesta ordem:
'        a) colunas de nascimento na propria planilha de matricula do
'           ano (ex.: "Municipio de Nascimento", "UF de Nascimento",
'           "Pais de Nascimento", ou uma coluna "Naturalidade");
'        b) um arquivo "naturalidade.xlsx" (ou .htm) em qualquer lugar
'           dentro da pasta de trabalho, com colunas Nome / Municipio /
'           Estado / Pais - o mesmo formato que LoadNaturalidade (em
'           modLeitura) sempre soube ler.
'
'   2) Carga horaria: escreve o valor digitado nas linhas
'      "TOTAL DE AULAS DA BASE NACIONAL COMUM CURRICULAR" e
'      "TOTAL DE AULAS DA PARTE DIVERSIFICADA" do modelo, na coluna do
'      ano correspondente. As linhas sao localizadas pelo texto do
'      rotulo, nao por numero fixo - se o modelo mudar de lugar, o
'      codigo continua achando.
'
'      Um valor unico (ex.: "800") vai para as duas linhas. Para valores
'      diferentes, digite os dois separados por barra:
'      "800/200" -> BNCC = 800, Parte Diversificada = 200.
'
'   3) Total da carga horaria anual (Horas): a linha
'      "TOTAL DA CARGA HORARIA ANUAL (Horas)" continua sendo uma formula
'      (nunca mais e sobrescrita por um numero digitado), mas o fator
'      muda sozinho conforme o ano da coluna:
'        ano < 2020  -> =(G52*50/60)   (aula de 50 minutos, como esta no modelo)
'        ano >= 2020 -> =(G52*45/60)   (aula de 45 minutos)
'      A linha referenciada (o "52") nao e fixa no codigo: sai da propria
'      formula que ja esta no modelo, ou da linha "TOTAL DE AULAS".
'
' Modulo autocontido: so usa funcoes que ja existem em modNucleo e
' modLeitura. Escrito sem acentos de proposito (o VBE importa .bas como
' ANSI; acento em arquivo UTF-8 vira lixo na hora de importar).
' =====================================================================

' ---------- cache ----------
Private gNatPath As String
Private gNatPathResolved As Boolean
Private gNatDict As Object

Private gRotuloChave As String
Private gRowBncc As Long
Private gRowDiver As Long
Private gRowHoras As Long
Private gRowTotalAulas As Long

Public Const RotuloTotalBncc As String = "TOTAL DE AULAS DA BASE NACIONAL COMUM CURRICULAR"
Public Const RotuloTotalDiver As String = "TOTAL DE AULAS DA PARTE DIVERSIFICADA"
Public Const RotuloTotalHoras As String = "TOTAL DA CARGA HORARIA ANUAL"
Public Const RotuloTotalAulas As String = "TOTAL DE AULAS"

' A partir deste ano a aula passou a ter 45 minutos (antes, 50).
Public Const AnoAula45 As Long = 2020

' Limpa o cache desta execucao. Opcional: chame junto com ResetCaches se
' voce trocar o arquivo de naturalidade com o Excel aberto.
Public Sub ResetNaturalidadeCache()
    gNatPath = ""
    gNatPathResolved = False
    Set gNatDict = Nothing
End Sub

' =====================================================================
' 1) NATURALIDADE AUTOMATICA
' =====================================================================

' Devolve sempre um clsNaturalidade (nunca Nothing). "origem" volta com
' a mensagem pronta pro log: de onde veio, ou o aviso de que nao achou.
Public Function BuscarNaturalidadeAuto(ByVal tl As clsTimeline, ByVal ano As Long, ByRef origem As String) As clsNaturalidade
    Dim n As clsNaturalidade
    Set n = New clsNaturalidade
    origem = ""

    Dim de As String
    If NaturalidadeDaMatricula(ano, tl.Cadastro.ra, n, de) Then
        origem = "Naturalidade preenchida automaticamente da matricula (" & de & ") - confira antes de gerar."
    ElseIf NaturalidadeDoArquivo(tl.nome, n, de) Then
        origem = "Naturalidade preenchida automaticamente de " & de & " - confira antes de gerar."
    End If

    If n.Municipio <> "" And n.pais = "" Then n.pais = "BRASIL"

    If n.Municipio = "" And n.Estado = "" And n.pais = "" Then
        origem = "AVISO: nao consegui descobrir municipio/UF/pais de nascimento automaticamente. " & _
                 "Procurei colunas de nascimento na matricula de " & ano & " e um arquivo comecando com " & _
                 """naturalidade"" dentro de " & BaseDir() & ". Preencha a mao antes de gerar."
    End If

    Set BuscarNaturalidadeAuto = n
End Function

' ---------- fonte (a): colunas de nascimento na planilha de matricula ----------

Private Function NaturalidadeDaMatricula(ByVal ano As Long, ByVal ra As String, ByVal n As clsNaturalidade, ByRef de As String) As Boolean
    Dim target As String
    target = NormRA(ra)
    If target = "" Then Exit Function

    Dim t As Integer
    For t = 0 To 1
        Dim files As Collection
        Set files = ListTurmaFiles(ListaDir(t, ano))
        Dim tf As clsTurmaFile
        For Each tf In files
            If IsMatriculaFormatSuportado(tf.ext) Then
                If LerNaturalidadeDaMatricula(tf.path, target, n) Then
                    de = "arquivo " & NomeDoArquivo(tf.path)
                    NaturalidadeDaMatricula = True
                    Exit Function
                End If
            End If
        Next tf
    Next t
End Function

Private Function LerNaturalidadeDaMatricula(ByVal path As String, ByVal targetRA As String, ByVal n As clsNaturalidade) As Boolean
    Dim wb As Workbook, ws As Worksheet
    On Error GoTo limpar
    Set wb = OpenReadOnly(path)
    Set ws = wb.Worksheets(1)

    Dim headerRow As Long
    If Not FindHeaderRow(ws, "Nome do Aluno", headerRow) Then GoTo limpar

    Dim colRA As Long
    colRA = ColOfHeader(ws, headerRow, "RA")
    If colRA = 0 Then GoTo limpar

    ' Municipio: "Municipio de Nascimento" > "Naturalidade" > qualquer "Municipio"
    Dim colMun As Long, colUF As Long, colPais As Long
    colMun = ColContendo(ws, headerRow, "MUNICIPIO", "NASC")
    If colMun = 0 Then colMun = ColContendo(ws, headerRow, "NATURALIDADE", "")
    If colMun = 0 Then colMun = ColContendo(ws, headerRow, "MUNICIPIO", "")

    ' UF: exige "NASC" junto, senao "UF RA" (o UF do RA) seria confundido
    ' com o UF de nascimento.
    colUF = ColContendo(ws, headerRow, "UF", "NASC")
    If colUF = 0 Then colUF = ColContendo(ws, headerRow, "ESTADO", "NASC")
    If colUF = 0 Then colUF = ColOfHeader(ws, headerRow, "Estado")

    colPais = ColContendo(ws, headerRow, "PAIS", "NASC")
    If colPais = 0 Then colPais = ColOfHeader(ws, headerRow, "Pais")
    If colPais = 0 Then colPais = ColContendo(ws, headerRow, "NACIONALIDADE", "")

    If colMun = 0 And colUF = 0 And colPais = 0 Then GoTo limpar

    Dim lastRow As Long
    lastRow = ws.UsedRange.row + ws.UsedRange.rows.Count - 1
    Dim r As Long
    For r = headerRow + 1 To lastRow
        If NormRA(CellText(ws.Cells(r, colRA))) = targetRA Then
            If colMun > 0 Then n.Municipio = CellText(ws.Cells(r, colMun))
            If colUF > 0 Then n.Estado = CellText(ws.Cells(r, colUF))
            If colPais > 0 Then n.pais = CellText(ws.Cells(r, colPais))
            LerNaturalidadeDaMatricula = (n.Municipio <> "" Or n.Estado <> "" Or n.pais <> "")
            Exit For
        End If
    Next r

limpar:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
End Function

' Primeira coluna da linha de cabecalho cujo texto (normalizado) contem
' must1 e tambem must2 (must2 = "" ignora a segunda condicao).
Private Function ColContendo(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal must1 As String, ByVal must2 As String) As Long
    Dim c As Long, firstCol As Long, lastCol As Long, h As String
    firstCol = ws.UsedRange.Column
    lastCol = firstCol + ws.UsedRange.Columns.Count - 1
    For c = firstCol To lastCol
        h = NormSubject(CellText(ws.Cells(headerRow, c)))
        If h <> "" And InStr(h, must1) > 0 Then
            If must2 = "" Or InStr(h, must2) > 0 Then
                ColContendo = c
                Exit Function
            End If
        End If
    Next c
End Function

' ---------- fonte (b): arquivo naturalidade.xlsx ----------

Private Function NaturalidadeDoArquivo(ByVal nome As String, ByVal n As clsNaturalidade, ByRef de As String) As Boolean
    Dim path As String
    path = AcharArquivoNaturalidade()
    If path = "" Then Exit Function

    If gNatDict Is Nothing Then Set gNatDict = LoadNaturalidade(path)
    Dim k As String
    k = NormName(nome)
    If Not gNatDict.Exists(k) Then Exit Function

    Dim src As clsNaturalidade
    Set src = gNatDict(k)
    n.Municipio = src.Municipio
    n.Estado = src.Estado
    n.pais = src.pais
    de = NomeDoArquivo(path)
    NaturalidadeDoArquivo = (n.Municipio <> "" Or n.Estado <> "" Or n.pais <> "")
End Function

' Procura, dentro de BaseDir() e ate 3 niveis abaixo, um arquivo cujo
' nome comece com "naturalidade" e tenha extensao suportada.
Private Function AcharArquivoNaturalidade() As String
    If gNatPathResolved Then
        AcharArquivoNaturalidade = gNatPath
        Exit Function
    End If
    gNatPath = ProcurarNaturalidadeEm(BaseDir(), 0)
    gNatPathResolved = True
    AcharArquivoNaturalidade = gNatPath
End Function

Private Function ProcurarNaturalidadeEm(ByVal dirPath As String, ByVal nivel As Integer) As String
    If nivel > 3 Then Exit Function

    Dim subdirs As New Collection
    Dim fn As String
    On Error Resume Next
    fn = Dir(dirPath & "\*.*", vbDirectory)
    On Error GoTo 0
    Do While fn <> ""
        If fn <> "." And fn <> ".." Then
            Dim full As String, attr As Long
            full = dirPath & "\" & fn
            attr = 0
            On Error Resume Next
            attr = GetAttr(full)
            On Error GoTo 0
            If (attr And vbDirectory) = vbDirectory Then
                subdirs.Add full
            Else
                Dim dotPos As Long, ext As String, stem As String
                dotPos = InStrRev(fn, ".")
                If dotPos > 0 Then
                    ext = LCase$(Mid$(fn, dotPos))
                    stem = Left$(fn, dotPos - 1)
                    If IsMatriculaFormatSuportado(ext) Then
                        If Left$(NormName(stem), 12) = "NATURALIDADE" Then
                            ProcurarNaturalidadeEm = full
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        fn = Dir()
    Loop

    Dim d As Variant
    For Each d In subdirs
        Dim achou As String
        achou = ProcurarNaturalidadeEm(CStr(d), nivel + 1)
        If achou <> "" Then
            ProcurarNaturalidadeEm = achou
            Exit Function
        End If
    Next d
End Function

Private Function NomeDoArquivo(ByVal path As String) As String
    Dim p As Long
    p = InStrRev(path, "\")
    If p > 0 Then
        NomeDoArquivo = Mid$(path, p + 1)
    Else
        NomeDoArquivo = path
    End If
End Function

' =====================================================================
' 2) CARGA HORARIA NAS LINHAS DE "TOTAL DE AULAS"
' =====================================================================

Public Sub AplicarCargaHoraria(ByVal ws As Worksheet, ByVal col As String, ByVal valor As String, ByVal warnings As Collection)
    If Trim$(valor) = "" Then Exit Sub

    Dim vBncc As String, vDiver As String
    SepararCarga valor, vBncc, vDiver
    ResolverLinhasTotal ws

    If gRowBncc > 0 Then
        EscreverCarga ws, col, gRowBncc, vBncc
    ElseIf Not warnings Is Nothing Then
        warnings.Add "nao achei a linha """ & RotuloTotalBncc & """ no modelo - a carga horaria da BNCC nao foi escrita"
    End If

    If gRowDiver > 0 Then
        EscreverCarga ws, col, gRowDiver, vDiver
    ElseIf Not warnings Is Nothing Then
        warnings.Add "nao achei a linha """ & RotuloTotalDiver & """ no modelo - a carga horaria da parte diversificada nao foi escrita"
    End If
End Sub

' "800" -> as duas linhas recebem 800.
' "800/200" (ou "800;200") -> BNCC = 800, Parte Diversificada = 200.
Private Sub SepararCarga(ByVal valor As String, ByRef vBncc As String, ByRef vDiver As String)
    Dim s As String, p As Long
    s = Trim$(valor)
    s = Replace(s, ";", "/")
    s = Replace(s, "|", "/")
    p = InStr(s, "/")
    If p > 0 Then
        vBncc = Trim$(Left$(s, p - 1))
        vDiver = Trim$(Mid$(s, p + 1))
    Else
        vBncc = s
        vDiver = s
    End If
End Sub

' ---------- total da carga horaria anual (formula 50/60 ou 45/60) ----------

' Reescreve a formula da linha "TOTAL DA CARGA HORARIA ANUAL (Horas)" na
' coluna do ano, trocando so o fator conforme o ano. Roda sempre, mesmo
' que o usuario nao tenha digitado carga horaria nenhuma.
Public Sub AplicarHorasAnuais(ByVal ws As Worksheet, ByVal col As String, ByVal ano As Long, ByVal warnings As Collection)
    ResolverLinhasTotal ws

    If gRowHoras = 0 Then
        If Not warnings Is Nothing Then
            warnings.Add "nao achei a linha """ & RotuloTotalHoras & """ no modelo - a formula das horas nao foi ajustada"
        End If
        Exit Sub
    End If

    Dim refRow As Long
    refRow = LinhaDaReferencia(ws.Range(col & gRowHoras).Formula)
    If refRow = 0 Then refRow = gRowTotalAulas
    If refRow = 0 Then
        If Not warnings Is Nothing Then
            warnings.Add "nao consegui descobrir qual linha a formula de """ & RotuloTotalHoras & """ soma - a formula da coluna " & col & " ficou como estava"
        End If
        Exit Sub
    End If

    Dim fator As Long
    If ano >= AnoAula45 Then fator = 45 Else fator = 50
    ws.Range(col & gRowHoras).Formula = "=(" & col & refRow & "*" & fator & "/60)"
End Sub

' Numero da linha da primeira referencia de celula da formula:
' "=(G52*50/60)" -> 52. Devolve 0 se a celula nao tiver formula.
Private Function LinhaDaReferencia(ByVal f As String) As Long
    Dim i As Long, ch As String, prev As String, digits As String
    If Left$(Trim$(f), 1) <> "=" Then Exit Function
    digits = ""
    For i = 1 To Len(f)
        ch = Mid$(f, i, 1)
        If ch >= "0" And ch <= "9" Then
            If digits = "" Then
                prev = ""
                If i > 1 Then prev = UCase$(Mid$(f, i - 1, 1))
                ' so conta se o digito vier logo depois da letra da coluna
                ' (ou do "$" de uma referencia absoluta, tipo $G$52)
                If (prev >= "A" And prev <= "Z") Or prev = "$" Then digits = ch
            Else
                digits = digits & ch
            End If
        ElseIf digits <> "" Then
            Exit For
        End If
    Next i
    If digits <> "" Then LinhaDaReferencia = CLng(digits)
End Function

' Acha as linhas pelo texto do rotulo (em qualquer coluna) e guarda o
' resultado, pra nao varrer a planilha inteira uma vez por ano.
Private Sub ResolverLinhasTotal(ByVal ws As Worksheet)
    Dim chave As String
    chave = ws.Parent.Name & "|" & ws.Name
    If gRotuloChave = chave Then Exit Sub
    gRotuloChave = chave
    gRowBncc = LinhaPorRotulo(ws, RotuloTotalBncc)
    gRowDiver = LinhaPorRotulo(ws, RotuloTotalDiver)
    gRowHoras = LinhaPorRotulo(ws, RotuloTotalHoras)
    ' "TOTAL DE AULAS" sozinho: precisa ser igual, senao casaria com
    ' "TOTAL DE AULAS DA BASE NACIONAL COMUM CURRICULAR".
    gRowTotalAulas = LinhaPorRotuloExato(ws, RotuloTotalAulas)
End Sub

Private Function LinhaPorRotulo(ByVal ws As Worksheet, ByVal rotulo As String) As Long
    Dim alvo As String
    alvo = NormSubject(rotulo)
    Dim firstRow As Long, lastRow As Long, firstCol As Long, lastCol As Long
    firstRow = ws.UsedRange.row
    lastRow = firstRow + ws.UsedRange.rows.Count - 1
    firstCol = ws.UsedRange.Column
    lastCol = firstCol + ws.UsedRange.Columns.Count - 1
    Dim r As Long, c As Long
    For r = firstRow To lastRow
        For c = firstCol To lastCol
            If InStr(NormSubject(CellText(ws.Cells(r, c))), alvo) > 0 Then
                LinhaPorRotulo = r
                Exit Function
            End If
        Next c
    Next r
End Function

Private Function LinhaPorRotuloExato(ByVal ws As Worksheet, ByVal rotulo As String) As Long
    Dim alvo As String
    alvo = NormSubject(rotulo)
    Dim firstRow As Long, lastRow As Long, firstCol As Long, lastCol As Long
    firstRow = ws.UsedRange.row
    lastRow = firstRow + ws.UsedRange.rows.Count - 1
    firstCol = ws.UsedRange.Column
    lastCol = firstCol + ws.UsedRange.Columns.Count - 1
    Dim r As Long, c As Long
    For r = firstRow To lastRow
        For c = firstCol To lastCol
            If NormSubject(CellText(ws.Cells(r, c))) = alvo Then
                LinhaPorRotuloExato = r
                Exit Function
            End If
        Next c
    Next r
End Function

Private Sub EscreverCarga(ByVal ws As Worksheet, ByVal col As String, ByVal r As Long, ByVal valor As String)
    If Trim$(valor) = "" Then Exit Sub
    Dim num As Double
    If ValorNumerico(valor, num) Then
        ws.Range(col & r).Value = num
    Else
        ws.Range(col & r).Value = valor
    End If
End Sub

' IsNumeric/CDbl seguem a configuracao regional do Windows (em pt-BR,
' CDbl("800.5") daria 8005). Val() nao segue: sempre le ponto como
' separador decimal. Por isso a checagem e feita na mao aqui.
Private Function ValorNumerico(ByVal s As String, ByRef outNum As Double) As Boolean
    Dim t As String, i As Long, ch As String, seps As Long
    t = Replace(Trim$(s), " ", "")
    t = Replace(t, ",", ".")
    If t = "" Then Exit Function
    seps = 0
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch = "." Then
            seps = seps + 1
            If seps > 1 Then Exit Function
        ElseIf ch < "0" Or ch > "9" Then
            Exit Function
        End If
    Next i
    outNum = Val(t)
    ValorNumerico = True
End Function
