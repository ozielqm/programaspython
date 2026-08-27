Attribute VB_Name = "modLeitura"
Option Explicit

' =====================================================================
' modLeitura — carrega Matrícula, Mapão e naturalidade.xlsx; e as
' funções de "arquivo" que varrem a estrutura de pastas (anos/turmas)
' pra achar o aluno em cada ano. Espelha data.go + as funções Archive
' de model.go do .exe.
' =====================================================================

' Cache da pasta "raiz" resolvida (Arvore) de cada árvore (concluintes/
' transferidos) — fica no topo do módulo de propósito, não logo depois de
' um End Sub/End Function.
Private gArvoreDir(0 To 1) As String
Private gArvoreResolved(0 To 1) As Boolean

' ---------- Matrícula ----------

Function LoadMatricula(ByVal path As String, ByVal turma As String) As Collection
    Dim result As New Collection
    Dim wb As Workbook, ws As Worksheet
    Set wb = OpenReadOnly(path)
    Set ws = wb.Worksheets(1)

    Dim headerRow As Long
    If Not FindHeaderRow(ws, "Nome do Aluno", headerRow) Then
        wb.Close False
        Err.Raise vbObjectError + 1001, , "coluna ""Nome do Aluno"" não encontrada em " & path & " (layout inesperado)"
    End If

    Dim colNome As Long, colRA As Long, colDig As Long, colUF As Long
    Dim colNasc As Long, colSit As Long, colSerie As Long, colTipo As Long
    colNome = ColOfHeader(ws, headerRow, "Nome do Aluno")
    colRA = ColOfHeader(ws, headerRow, "RA")
    colDig = ColOfHeader(ws, headerRow, "Dig. RA")
    colUF = ColOfHeader(ws, headerRow, "UF RA")
    colNasc = ColOfHeader(ws, headerRow, "Data de Nascimento")
    colSit = ColOfHeader(ws, headerRow, "Situação")
    colSerie = ColOfHeader(ws, headerRow, "Série")
    colTipo = ColOfHeader(ws, headerRow, "Tipo de Ensino")

    Dim lastRow As Long
    lastRow = ws.UsedRange.row + ws.UsedRange.rows.Count - 1
    Dim r As Long
    Dim row As clsMatriculaRow
    For r = headerRow + 1 To lastRow
        Dim nome As String
        nome = CellText(ws.Cells(r, colNome))
        If nome <> "" Then
            Set row = New clsMatriculaRow
            row.nome = nome
            If colRA > 0 Then row.ra = CellText(ws.Cells(r, colRA))
            If colDig > 0 Then row.DigRA = CellText(ws.Cells(r, colDig))
            If colUF > 0 Then row.UFRA = CellText(ws.Cells(r, colUF))
            If colNasc > 0 Then row.Nascimento = FormatDateCell(ws.Cells(r, colNasc))
            If colSit > 0 Then row.Situacao = CellText(ws.Cells(r, colSit))
            If colSerie > 0 Then row.Serie = CIntSafe(ws.Cells(r, colSerie).Value)
            If colTipo > 0 Then row.TipoEnsino = CellText(ws.Cells(r, colTipo))
            row.turma = turma
            result.Add row
        End If
    Next r
    wb.Close False
    Set LoadMatricula = result
End Function

' ---------- Mapão (notas) ----------

Function IsWordy(ByVal s As String) As Boolean
    Dim t As String, i As Long, letters As Long, ch As String
    t = Trim$(s)
    If Len(t) < 3 Then
        IsWordy = False
        Exit Function
    End If
    letters = 0
    For i = 1 To Len(t)
        ch = UCase$(Mid$(t, i, 1))
        If ch >= "A" And ch <= "Z" Then letters = letters + 1
    Next i
    IsWordy = (letters >= 3)
End Function

Function IsExcludedSubject(ByVal s As String) As Boolean
    Select Case s
        Case "NOME", "SIT", "SITUACAO", "ALUNO", "TOTAL", "N", "M"
            IsExcludedSubject = True
        Case Else
            IsExcludedSubject = False
    End Select
End Function

Function ScoreRow(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal firstCol As Long, ByVal lastCol As Long) As Long
    Dim c As Long, n As Long
    n = 0
    For c = firstCol To lastCol
        If IsWordy(CellText(ws.Cells(rowNum, c))) Then n = n + 1
    Next c
    ScoreRow = n
End Function

Function FindFourDigits(ByVal s As String) As String
    Dim i As Long, run As String, ch As String
    run = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            run = run & ch
            If Len(run) = 4 Then
                FindFourDigits = run
                Exit Function
            End If
        Else
            run = ""
        End If
    Next i
    FindFourDigits = ""
End Function

Sub ExtractAnoETipo(ByVal ws As Worksheet, ByRef anoLetivo As Long, ByRef tipo As String)
    Dim firstRow As Long, lastRow As Long, firstCol As Long, lastCol As Long
    firstRow = ws.UsedRange.row
    lastRow = firstRow + ws.UsedRange.rows.Count - 1
    firstCol = ws.UsedRange.Column
    lastCol = firstCol + ws.UsedRange.Columns.Count - 1
    anoLetivo = 0
    tipo = ""
    Dim r As Long, c As Long
    For r = firstRow To lastRow
        For c = firstCol To lastCol
            Dim v As String
            v = CellText(ws.Cells(r, c))
            If InStr(v, "Ano Letivo") > 0 Then
                Dim digits As String
                digits = FindFourDigits(v)
                If digits = "" Then
                    Dim c2 As Long
                    For c2 = firstCol To lastCol
                        digits = FindFourDigits(CellText(ws.Cells(r, c2)))
                        If digits <> "" Then Exit For
                    Next c2
                End If
                If digits <> "" Then anoLetivo = CLng(digits)
            End If
            If InStr(v, "Tipo de Fechamento") > 0 Or InStr(v, "Tipo Fechamento") > 0 Then
                Dim t As String
                t = v
                Dim c3 As Long
                For c3 = firstCol To lastCol
                    Dim v3 As String
                    v3 = CellText(ws.Cells(r, c3))
                    If v3 <> v And v3 <> "" And InStr(v3, "Tipo") = 0 Then
                        t = t & " " & v3
                    End If
                Next c3
                tipo = t
            End If
        Next c
    Next r
End Sub

Sub ClassifyTipo(ByVal tipo As String, ByRef isFinal As Boolean, ByRef bimestre As Integer)
    Dim up As String
    up = StripAccents(UCase$(tipo))
    isFinal = False
    bimestre = 0
    If InStr(up, "FINAL") > 0 Or InStr(up, "QUINTO CONCEITO") > 0 Then
        isFinal = True
        Exit Sub
    End If
    If InStr(up, "PRIMEIRO BIMESTRE") > 0 Or InStr(up, "1 BIMESTRE") > 0 Or InStr(up, "1BIMESTRE") > 0 Then
        bimestre = 1
    ElseIf InStr(up, "SEGUNDO BIMESTRE") > 0 Or InStr(up, "2 BIMESTRE") > 0 Then
        bimestre = 2
    ElseIf InStr(up, "TERCEIRO BIMESTRE") > 0 Or InStr(up, "3 BIMESTRE") > 0 Then
        bimestre = 3
    ElseIf InStr(up, "QUARTO BIMESTRE") > 0 Or InStr(up, "4 BIMESTRE") > 0 Then
        bimestre = 4
    End If
End Sub

' Lê tanto o layout "Conselho Final" (Nome/Sit/Nº/M juntos, nome de
' disciplina 1-3 linhas acima) quanto o layout de bimestre único
' (ALUNO/SITUAÇÃO + nomes de disciplina na mesma linha, Nº/M/F/AC na
' linha seguinte). Ver LOGICA_DO_PROGRAMA.md seção 8 para a explicação
' completa desta heurística.
Function LoadMapao(ByVal path As String) As clsNotasFechamento
    Dim nf As New clsNotasFechamento
    Dim wb As Workbook, ws As Worksheet
    Set wb = OpenReadOnly(path)
    Set ws = wb.Worksheets(1)

    Dim anoLetivo As Long, tipo As String
    ExtractAnoETipo ws, anoLetivo, tipo
    nf.anoLetivo = anoLetivo
    nf.tipo = tipo
    Dim isFinal As Boolean, bimestre As Integer
    ClassifyTipo tipo, isFinal, bimestre
    nf.isFinal = isFinal
    nf.bimestre = bimestre

    Dim headerRow As Long
    If Not FindHeaderRow(ws, "Nome", headerRow) Then
        If Not FindHeaderRow(ws, "Aluno", headerRow) Then
            wb.Close False
            Err.Raise vbObjectError + 1002, , "coluna ""Nome/Aluno"" não encontrada em " & path & " (layout inesperado)"
        End If
    End If
    Dim colNome As Long, colSit As Long
    colNome = ColOfHeader(ws, headerRow, "Nome")
    If colNome = 0 Then colNome = ColOfHeader(ws, headerRow, "Aluno")
    colSit = ColOfHeader(ws, headerRow, "Sit")
    If colSit = 0 Then colSit = ColOfHeader(ws, headerRow, "Situação")

    Dim firstCol As Long, lastCol As Long
    firstCol = ws.UsedRange.Column
    lastCol = firstCol + ws.UsedRange.Columns.Count - 1

    Dim hasM As Boolean, c As Long
    hasM = False
    For c = firstCol To lastCol
        If UCase$(CellText(ws.Cells(headerRow, c))) = "M" Then
            hasM = True
            Exit For
        End If
    Next c
    Dim subHeaderRow As Long
    subHeaderRow = headerRow
    If Not hasM Then subHeaderRow = headerRow + 1

    Dim subjectRow As Long, bestScore As Long, back As Integer, rr As Long, sc As Long
    subjectRow = 0
    bestScore = 0
    For back = 1 To 3
        rr = headerRow - back
        If rr >= ws.UsedRange.row Then
            sc = ScoreRow(ws, rr, firstCol, lastCol)
            If sc > bestScore Then
                bestScore = sc
                subjectRow = rr
            End If
        End If
    Next back

    Dim subjectByCol As Object
    Set subjectByCol = CreateObject("Scripting.Dictionary")
    If subjectRow > 0 Then
        For c = firstCol To lastCol
            Dim val As String
            val = NormSubject(CellText(ws.Cells(subjectRow, c)))
            If val <> "" And Not IsExcludedSubject(val) And IsWordy(val) Then
                subjectByCol(CStr(c)) = val
            End If
        Next c
    End If
    For c = firstCol To lastCol
        Dim val2 As String
        val2 = NormSubject(CellText(ws.Cells(headerRow, c)))
        If val2 <> "" And Not IsExcludedSubject(val2) And IsWordy(val2) Then
            If Not subjectByCol.Exists(CStr(c)) Then subjectByCol(CStr(c)) = val2
        End If
    Next c

    Dim mColFor() As Long
    ReDim mColFor(firstCol To lastCol)
    For c = firstCol + 1 To lastCol
        If UCase$(CellText(ws.Cells(subHeaderRow, c))) = "M" Then
            mColFor(c - 1) = c
        End If
    Next c

    Dim dataStartRow As Long
    dataStartRow = headerRow
    If subHeaderRow > dataStartRow Then dataStartRow = subHeaderRow
    Dim lastRow As Long
    lastRow = ws.UsedRange.row + ws.UsedRange.rows.Count - 1

    Dim r As Long
    Dim grades As Collection
    Dim g As clsSubjectGrade
    For r = dataStartRow + 1 To lastRow
        Dim nome As String
        nome = CellText(ws.Cells(r, colNome))
        If nome <> "" Then
            Dim nn As String
            nn = NormName(nome)
            If colSit > 0 Then nf.SituacaoPorAluno(nn) = CellText(ws.Cells(r, colSit))
            Set grades = New Collection
            Dim key As Variant
            For Each key In subjectByCol.Keys
                Dim subjCol As Long, mCol As Long
                subjCol = CLng(key)
                mCol = mColFor(subjCol)
                If mCol <> 0 Then
                    Set g = New clsSubjectGrade
                    g.Subject = subjectByCol(key)
                    g.Value = CellText(ws.Cells(r, mCol))
                    grades.Add g
                End If
            Next key
            Set nf.PorAluno(nn) = grades
        End If
    Next r

    wb.Close False
    Set LoadMapao = nf
End Function

' ---------- Naturalidade ----------

Function LoadNaturalidade(ByVal path As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    If Dir(path) = "" Then
        Set LoadNaturalidade = dict
        Exit Function
    End If
    Dim wb As Workbook, ws As Worksheet
    Set wb = OpenReadOnly(path)
    Set ws = wb.Worksheets(1)
    Dim headerRow As Long
    If Not FindHeaderRow(ws, "Nome", headerRow) Then
        wb.Close False
        Set LoadNaturalidade = dict
        Exit Function
    End If
    Dim colNome As Long, colMun As Long, colUF As Long, colPais As Long
    colNome = ColOfHeader(ws, headerRow, "Nome")
    colMun = ColOfHeader(ws, headerRow, "Municipio")
    colUF = ColOfHeader(ws, headerRow, "Estado")
    colPais = ColOfHeader(ws, headerRow, "Pais")
    Dim lastRow As Long
    lastRow = ws.UsedRange.row + ws.UsedRange.rows.Count - 1
    Dim r As Long
    Dim n As clsNaturalidade
    For r = headerRow + 1 To lastRow
        Dim nome As String
        nome = CellText(ws.Cells(r, colNome))
        If nome <> "" Then
            Set n = New clsNaturalidade
            If colMun > 0 Then n.Municipio = CellText(ws.Cells(r, colMun))
            If colUF > 0 Then n.Estado = CellText(ws.Cells(r, colUF))
            If colPais > 0 Then n.pais = CellText(ws.Cells(r, colPais))
            Set dict(NormName(nome)) = n
        End If
    Next r
    wb.Close False
    Set LoadNaturalidade = dict
End Function

' ---------- Archive: varredura da estrutura de pastas ----------

Sub ResetArvoreCache()
    gArvoreResolved(0) = False
    gArvoreResolved(1) = False
End Sub

Function Arvore(ByVal tipoIdx As Integer) As String
    If tipoIdx = 0 Then
        Arvore = "pasta para alunos concluintes"
    Else
        Arvore = "pasta para alunos transferidos"
    End If
End Function

' Retorna o último grupo de exatamente 4 dígitos encontrado em s (ex.: "...
' em 2026" -> 2026), ou -1 se não achar nenhum.
Function ExtractYearFromName(ByVal s As String) As Long
    Dim i As Long, ch As String, digits As String, best As Long
    best = -1
    digits = ""
    For i = 1 To Len(s) + 1
        If i <= Len(s) Then ch = Mid$(s, i, 1) Else ch = " "
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        Else
            If Len(digits) = 4 Then best = CLng(digits)
            digits = ""
        End If
    Next i
    ExtractYearFromName = best
End Function

' Acha, dentro de BaseDir(), a subpasta cujo nome COMEÇA com o nome
' esperado (ignorando acento/maiúscula/minúscula) — tolera um sufixo tipo
' " em 2026" que algumas máquinas usam em vez do nome exato "pasta para
' alunos transferidos". Se houver mais de uma pasta assim, prefere a que
' tiver o ano mais recente no nome. Resolve uma vez só por execução
' (ResetArvoreCache limpa o cache no início de cada "Gerar Histórico(s)").
Function ArvoreDir(ByVal tipoIdx As Integer) As String
    If gArvoreResolved(tipoIdx) Then
        ArvoreDir = gArvoreDir(tipoIdx)
        Exit Function
    End If
    Dim wantPrefix As String
    wantPrefix = NormName(Arvore(tipoIdx))
    Dim bestName As String, bestYear As Long
    bestName = ""
    bestYear = -1
    Dim folder As String
    On Error Resume Next
    folder = Dir(BaseDir() & "\", vbDirectory)
    On Error GoTo 0
    Do While folder <> ""
        If folder <> "." And folder <> ".." Then
            If (GetAttr(BaseDir() & "\" & folder) And vbDirectory) = vbDirectory Then
                If Left$(NormName(folder), Len(wantPrefix)) = wantPrefix Then
                    Dim y As Long
                    y = ExtractYearFromName(folder)
                    If bestName = "" Or y > bestYear Then
                        bestName = folder
                        bestYear = y
                    End If
                End If
            End If
        End If
        folder = Dir()
    Loop
    If bestName = "" Then
        ' não achou nenhuma pasta com esse prefixo — usa o nome literal
        ' mesmo assim, pra mensagem de erro mais adiante mostrar esse caminho
        bestName = Arvore(tipoIdx)
    End If
    gArvoreDir(tipoIdx) = BaseDir() & "\" & bestName
    gArvoreResolved(tipoIdx) = True
    ArvoreDir = gArvoreDir(tipoIdx)
End Function

Function ListaDir(ByVal tipoIdx As Integer, ByVal ano As Long) As String
    Dim sub_ As String
    If tipoIdx = 0 Then
        sub_ = "pasta da Lista de alunos concluintes"
    Else
        sub_ = "pasta da Lista de alunos transferidos"
    End If
    If ano = 0 Then
        ListaDir = ArvoreDir(tipoIdx) & "\" & sub_
    Else
        ListaDir = ArvoreDir(tipoIdx) & "\" & sub_ & "\" & CStr(ano)
    End If
End Function

Function NotasDir(ByVal tipoIdx As Integer, ByVal ano As Long) As String
    NotasDir = ArvoreDir(tipoIdx) & "\pasta de notas dos alunos\" & CStr(ano)
End Function

Function ListYearDirs(ByVal parentDir As String) As Collection
    Dim result As New Collection
    Dim folder As String
    On Error Resume Next
    folder = Dir(parentDir & "\", vbDirectory)
    On Error GoTo 0
    Do While folder <> ""
        If folder <> "." And folder <> ".." Then
            If (GetAttr(parentDir & "\" & folder) And vbDirectory) = vbDirectory Then
                If IsNumeric(folder) Then result.Add CLng(folder)
            End If
        End If
        folder = Dir()
    Loop
    Set ListYearDirs = result
End Function

Function ListTurmaFiles(ByVal dirPath As String) As Collection
    Dim result As New Collection
    Dim fn As String
    Dim tf As clsTurmaFile
    On Error Resume Next
    fn = Dir(dirPath & "\*.*")
    On Error GoTo 0
    Do While fn <> ""
        Dim dotPos As Long, ext As String, stem As String
        dotPos = InStrRev(fn, ".")
        If dotPos > 0 Then
            ext = LCase$(Mid$(fn, dotPos))
            stem = Left$(fn, dotPos - 1)
        Else
            ext = ""
            stem = fn
        End If
        Dim grade As Integer, letter As String
        If ParseTurma(stem, grade, letter) Then
            Set tf = New clsTurmaFile
            tf.grade = grade
            tf.letter = letter
            tf.path = dirPath & "\" & fn
            tf.ext = ext
            result.Add tf
        End If
        fn = Dir()
    Loop
    Set ListTurmaFiles = result
End Function

Function LatestYear() As Long
    Dim best As Long, t As Integer
    best = 0
    For t = 0 To 1
        Dim years As Collection
        Set years = ListYearDirs(ListaDir(t, 0))
        Dim y As Variant
        For Each y In years
            If CLng(y) > best Then best = CLng(y)
        Next y
    Next t
    LatestYear = best
End Function

' Procura o RA em todas as turmas (ambas as árvores) do ano indicado.
Function FindStudentInYearByRA(ByVal ano As Long, ByVal ra As String, ByRef matOut As clsMatriculaRow, ByRef turmaOut As String) As Boolean
    Dim target As String
    target = NormRA(ra)
    If target = "" Then
        FindStudentInYearByRA = False
        Exit Function
    End If
    Dim t As Integer
    For t = 0 To 1
        Dim files As Collection
        Set files = ListTurmaFiles(ListaDir(t, ano))
        Dim tf As clsTurmaFile
        For Each tf In files
            If IsMatriculaFormatSuportado(tf.ext) Then
                Dim turmaLabel As String
                turmaLabel = CStr(tf.grade) & tf.letter
                Dim rows As Collection
                Set rows = LoadMatriculaCached(tf.path, turmaLabel)
                Dim row As clsMatriculaRow
                For Each row In rows
                    If NormRA(row.ra) = target Then
                        Set matOut = row
                        turmaOut = turmaLabel
                        FindStudentInYearByRA = True
                        Exit Function
                    End If
                Next row
            End If
        Next tf
    Next t
    FindStudentInYearByRA = False
End Function

' Alternativa por nome (usada como fallback quando o RA não bate).
Function FindStudentInYearByName(ByVal ano As Long, ByVal nome As String, ByRef matOut As clsMatriculaRow, ByRef turmaOut As String) As Boolean
    Dim target As String
    target = NormName(nome)
    If target = "" Then
        FindStudentInYearByName = False
        Exit Function
    End If
    Dim t As Integer
    For t = 0 To 1
        Dim files As Collection
        Set files = ListTurmaFiles(ListaDir(t, ano))
        Dim tf As clsTurmaFile
        For Each tf In files
            If IsMatriculaFormatSuportado(tf.ext) Then
                Dim turmaLabel As String
                turmaLabel = CStr(tf.grade) & tf.letter
                Dim rows As Collection
                Set rows = LoadMatriculaCached(tf.path, turmaLabel)
                Dim row As clsMatriculaRow
                For Each row In rows
                    If NormName(row.nome) = target Then
                        Set matOut = row
                        turmaOut = turmaLabel
                        FindStudentInYearByName = True
                        Exit Function
                    End If
                Next row
            End If
        Next tf
    Next t
    FindStudentInYearByName = False
End Function

Function LoadMatriculaOfTurma(ByVal ano As Long, ByVal grade As Integer, ByVal letter As String) As Collection
    Dim t As Integer
    For t = 0 To 1
        Dim files As Collection
        Set files = ListTurmaFiles(ListaDir(t, ano))
        Dim tf As clsTurmaFile
        For Each tf In files
            If tf.grade = grade And tf.letter = letter And IsMatriculaFormatSuportado(tf.ext) Then
                Set LoadMatriculaOfTurma = LoadMatriculaCached(tf.path, CStr(grade) & letter)
                Exit Function
            End If
        Next tf
    Next t
    Set LoadMatriculaOfTurma = Nothing
End Function

' Todos os Mapões (final + bimestres) de um ano/turma, em ambas as árvores.
Function LoadNotas(ByVal ano As Long, ByVal turma As String) As Collection
    Dim grade As Integer, letter As String
    ParseTurma turma, grade, letter
    Dim result As New Collection
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim t As Integer
    For t = 0 To 1
        Dim files As Collection
        Set files = ListTurmaFiles(NotasDir(t, ano))
        Dim tf As clsTurmaFile
        For Each tf In files
            If tf.grade = grade And tf.letter = letter Then
                If tf.ext = ".xlsx" Then
                    Dim fname As String
                    fname = Mid$(tf.path, InStrRev(tf.path, "\") + 1)
                    If Not seen.Exists(fname) Then
                        seen.Add fname, True
                        Dim nf As clsNotasFechamento
                        Set nf = LoadMapaoCached(tf.path)
                        result.Add nf
                    End If
                End If
            End If
        Next tf
    Next t
    Set LoadNotas = result
End Function

' Todos os alunos matriculados num ano especifico, juntando as duas
' arvores (concluintes e transferidos) e todas as turmas — usado pelo
' painel de comparacao de alunos entre anos.
Function ListarAlunosDoAno(ByVal ano As Long) As Collection
    Dim result As New Collection
    Dim t As Integer
    For t = 0 To 1
        Dim files As Collection
        Set files = ListTurmaFiles(ListaDir(t, ano))
        Dim tf As clsTurmaFile
        For Each tf In files
            If IsMatriculaFormatSuportado(tf.ext) Then
                Dim turmaLabel As String
                turmaLabel = CStr(tf.grade) & tf.letter
                Dim rows As Collection
                Set rows = LoadMatriculaCached(tf.path, turmaLabel)
                Dim row As clsMatriculaRow
                For Each row In rows
                    result.Add row
                Next row
            End If
        Next tf
    Next t
    Set ListarAlunosDoAno = result
End Function

' Lista de turmas (rótulos, ex: "5A") encontradas no ano mais recente —
' usada para popular o combo de turmas no formulário.
Function ListarTurmasAnoAtual() As Collection
    Dim result As New Collection
    Dim ano As Long
    ano = LatestYear()
    If ano = 0 Then
        Set ListarTurmasAnoAtual = result
        Exit Function
    End If
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim t As Integer
    For t = 0 To 1
        Dim files As Collection
        Set files = ListTurmaFiles(ListaDir(t, ano))
        Dim tf As clsTurmaFile
        For Each tf In files
            If IsMatriculaFormatSuportado(tf.ext) Then
                Dim label As String
                label = CStr(tf.grade) & tf.letter
                If Not seen.Exists(label) Then
                    seen.Add label, True
                    result.Add label
                End If
            End If
        Next tf
    Next t
    Set ListarTurmasAnoAtual = result
End Function

