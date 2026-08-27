Attribute VB_Name = "modNucleo"
Option Explicit

' =====================================================================
' modNucleo — normalização de texto, helpers compartilhados de leitura
' de planilha, e o cache de arquivos já abertos nesta execução.
' Espelha data.go do programa PREENCHENDO_HISTORICO (.exe), adaptado
' para usar o modelo de objetos do próprio Excel em vez de ler o XML
' na mão.
' =====================================================================

Public Const EscolaNome As String = "EE MARILENE TEREZINHA LONGHIM"

' ---------- cache (evita reabrir o mesmo arquivo várias vezes na mesma execução) ----------
Public gMatriculaCache As Object   ' Dictionary: path -> Collection de clsMatriculaRow
Public gMapaoCache As Object       ' Dictionary: path -> clsNotasFechamento

' Pasta onde este arquivo .xlsm está salvo — é a raiz de tudo, igual ao
' .exe (que usa a pasta onde ele está salvo).
'
' Em arquivos do OneDrive com o AutoSave ligado, o Excel às vezes devolve
' em ThisWorkbook.Path um endereço de nuvem (tipo
' "https://d.docs.live.net/<id>/Desktop/HISTÓRICOS") em vez do caminho
' local de disco. Dir() não entende endereço de nuvem e trava com erro 52.
' Então, quando isso acontece, convertemos pro caminho local de verdade
' usando a variável de ambiente que o próprio OneDrive mantém.
' Único computador/pasta onde este arquivo roda — usado como garantia
' caso a conversão do endereço de nuvem abaixo não bata com uma pasta
' local de verdade (ex.: formato de URL diferente do esperado).
' FallbackLocalBaseDir foi substituida por PastaHistoricosLocal(), em
' modAutoPreenche - ver comentario la.

Public Function BaseDir() As String
    Dim p As String
    p = ThisWorkbook.path
    If LCase$(Left$(p, 4)) = "http" Then
        p = ResolveOneDriveCloudPath(p)
        If LCase$(Left$(p, 4)) = "http" Or Not FolderReallyExists(p) Then
            ' antes: FallbackLocalBaseDir (perfil de usuario fixo, que so
            ' existia numa das maquinas). Agora a pasta e procurada.
            p = PastaHistoricosLocal()
        End If
    End If
    BaseDir = p
End Function

Private Function FolderReallyExists(ByVal path As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderReallyExists = fso.FolderExists(path)
    On Error GoTo 0
End Function

Private Function ResolveOneDriveCloudPath(ByVal cloudPath As String) As String
    Dim oneDriveLocal As String
    oneDriveLocal = Environ$("OneDrive")
    If oneDriveLocal = "" Then oneDriveLocal = Environ$("OneDriveConsumer")
    If oneDriveLocal = "" Then oneDriveLocal = Environ$("OneDriveCommercial")

    ' O caminho de nuvem tem o formato "https://<host>/<id>/<caminho relativo>"
    ' — ou seja, tudo depois da 4ª barra é o caminho relativo dentro da
    ' pasta do OneDrive (ex.: "Desktop/HISTÓRICOS").
    Dim rel As String, i As Long, slashCount As Long
    slashCount = 0
    rel = ""
    For i = 1 To Len(cloudPath)
        If Mid$(cloudPath, i, 1) = "/" Then
            slashCount = slashCount + 1
            If slashCount = 4 Then
                rel = Mid$(cloudPath, i + 1)
                Exit For
            End If
        End If
    Next i
    rel = Replace(rel, "/", "\")

    If oneDriveLocal <> "" And rel <> "" Then
        ResolveOneDriveCloudPath = oneDriveLocal & "\" & rel
    Else
        ResolveOneDriveCloudPath = cloudPath
    End If
End Function

' ---------- normalização de texto ----------

Function StripAccents(ByVal s As String) As String
    Dim de As String, para As String
    de = "áàâãäéèêëíìîïóòôõöúùûüçñÁÀÂÃÄÉÈÊËÍÌÎÏÓÒÔÕÖÚÙÛÜÇÑ"
    para = "aaaaaeeeeiiiiooooouuuucnAAAAAEEEEIIIIOOOOOUUUUCN"
    Dim i As Long, p As Long, ch As String, out As String
    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        p = InStr(de, ch)
        If p > 0 Then
            out = out & Mid$(para, p, 1)
        Else
            out = out & ch
        End If
    Next i
    StripAccents = out
End Function

Function CollapseSpaces(ByVal s As String) As String
    Dim t As String
    t = s
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    CollapseSpaces = Trim$(t)
End Function

Function NormName(ByVal s As String) As String
    NormName = CollapseSpaces(UCase$(StripAccents(Trim$(s))))
End Function

' Igual ao normSubject do .exe: maiúsculas, sem acento, e qualquer
' caractere que não seja letra ou espaço vira espaço (descarta código
' numérico de disciplina colado no nome, tipo "LINGUA PORTUGUESA\n1813").
Function NormSubject(ByVal s As String) As String
    Dim t As String, i As Long, ch As String, out As String
    t = UCase$(StripAccents(Trim$(s)))
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbCr, " ")
    out = ""
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If (ch >= "A" And ch <= "Z") Or ch = " " Then
            out = out & ch
        Else
            out = out & " "
        End If
    Next i
    NormSubject = CollapseSpaces(out)
End Function

Function DigitsOnly(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next i
    DigitsOnly = out
End Function

Function NormRA(ByVal ra As String) As String
    NormRA = DigitsOnly(ra)
End Function

' Alguns arquivos de "Lista de alunos" (principalmente concluintes, anos
' mais antigos) não são um .xlsx de verdade — são uma tabela HTML exportada
' pelo Sistema (Secretaria Escolar Digital), salva com extensão .htm. O
' Excel abre esses arquivos normalmente com Workbooks.Open (ele detecta o
' conteúdo automaticamente), então só precisamos parar de rejeitar essa
' extensão nos pontos que filtram "só .xlsx". .xls antigo continua não
' suportado.
Function IsMatriculaFormatSuportado(ByVal ext As String) As Boolean
    Dim e As String
    e = LCase$(ext)
    IsMatriculaFormatSuportado = (e = ".xlsx" Or e = ".htm" Or e = ".html")
End Function

' Heurística: o texto digitado parece um RA (maioria de dígitos, sem
' nenhuma letra) ou um nome?
Function LooksLikeRA(ByVal s As String) As Boolean
    Dim t As String, d As String, i As Long, ch As String
    t = Trim$(s)
    d = DigitsOnly(t)
    If Len(d) < 5 Then
        LooksLikeRA = False
        Exit Function
    End If
    For i = 1 To Len(t)
        ch = UCase$(Mid$(t, i, 1))
        If ch >= "A" And ch <= "Z" Then
            LooksLikeRA = False
            Exit Function
        End If
    Next i
    LooksLikeRA = True
End Function

' ---------- reconhecimento de nome de arquivo/turma ("5°A", "5A", "1A") ----------

Function ParseTurma(ByVal stem As String, ByRef grade As Integer, ByRef letter As String) As Boolean
    Dim s As String, i As Long, ch As String
    s = Trim$(stem)
    letter = ""
    grade = 0
    If Len(s) = 0 Then
        ParseTurma = False
        Exit Function
    End If
    i = 1
    ch = Mid$(s, i, 1)
    If ch < "1" Or ch > "9" Then
        ParseTurma = False
        Exit Function
    End If
    grade = CInt(ch)
    i = i + 1
    Do While i <= Len(s) And Mid$(s, i, 1) = " "
        i = i + 1
    Loop
    If i <= Len(s) Then
        ch = Mid$(s, i, 1)
        ' ° = Chr(176), º = Chr(186), ª = Chr(170)
        If ch = Chr$(176) Or ch = Chr$(186) Or ch = Chr$(170) Or ch = "o" Or ch = "O" Then
            i = i + 1
        End If
    End If
    Do While i <= Len(s) And Mid$(s, i, 1) = " "
        i = i + 1
    Loop
    If i <= Len(s) Then
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Then
            letter = UCase$(ch)
        End If
    End If
    ParseTurma = True
End Function

Function SanitizeFilename(ByVal s As String) As String
    Dim out As String
    out = s
    out = Replace(out, "/", "-")
    out = Replace(out, "\", "-")
    out = Replace(out, ":", "-")
    out = Replace(out, "*", "")
    out = Replace(out, "?", "")
    out = Replace(out, """", "")
    out = Replace(out, "<", "")
    out = Replace(out, ">", "")
    out = Replace(out, "|", "")
    SanitizeFilename = out
End Function

' ---------- helpers de leitura de planilha ----------

Function OpenReadOnly(ByVal path As String) As Workbook
    Set OpenReadOnly = Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False, Notify:=False)
End Function

' Texto exatamente como aparece na célula (respeita formatação/locale) —
' evita os problemas de separador decimal (7 vs 7,5) que CStr(.Value)
' teria dependendo da configuração regional do Windows.
Function CellText(ByVal rng As Range) As String
    On Error Resume Next
    CellText = Trim$(rng.Text)
    On Error GoTo 0
End Function

Function CIntSafe(ByVal v As Variant) As Integer
    If IsNumeric(v) Then
        CIntSafe = CInt(v)
    Else
        CIntSafe = 0
    End If
End Function

Function FormatDateCell(ByVal c As Range) As String
    On Error Resume Next
    If IsDate(c.Value) Then
        FormatDateCell = Format(c.Value, "dd/mm/yyyy")
    Else
        FormatDateCell = CellText(c)
    End If
    On Error GoTo 0
End Function

' Acha a linha cuja faixa usada contém, em alguma célula, um texto que
' contém (normalizado) mustContain. Retorna True/False; a linha vai em
' headerRowOut.
Function FindHeaderRow(ByVal ws As Worksheet, ByVal mustContain As String, ByRef headerRowOut As Long) As Boolean
    Dim target As String, r As Long, c As Long
    Dim firstRow As Long, lastRow As Long, firstCol As Long, lastCol As Long
    target = NormSubject(mustContain)
    firstRow = ws.UsedRange.row
    lastRow = firstRow + ws.UsedRange.rows.Count - 1
    firstCol = ws.UsedRange.Column
    lastCol = firstCol + ws.UsedRange.Columns.Count - 1
    For r = firstRow To lastRow
        For c = firstCol To lastCol
            If InStr(NormSubject(CellText(ws.Cells(r, c))), target) > 0 Then
                headerRowOut = r
                FindHeaderRow = True
                Exit Function
            End If
        Next c
    Next r
    FindHeaderRow = False
End Function

' Acha, na linha headerRow, a coluna cujo texto (normalizado) é
' EXATAMENTE igual a want. Retorna 0 se não achar.
Function ColOfHeader(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal want As String) As Long
    Dim target As String, c As Long, firstCol As Long, lastCol As Long
    target = NormSubject(want)
    firstCol = ws.UsedRange.Column
    lastCol = firstCol + ws.UsedRange.Columns.Count - 1
    For c = firstCol To lastCol
        If NormSubject(CellText(ws.Cells(headerRow, c))) = target Then
            ColOfHeader = c
            Exit Function
        End If
    Next c
    ColOfHeader = 0
End Function

Sub EnsureFolder(ByVal path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = Replace(path, "/", "\")
    If Not fso.FolderExists(path) Then
        ' cria pasta a pasta (e pais, se preciso) recursivamente
        Dim parts() As String, cur As String, i As Long
        parts = Split(path, "\")
        cur = parts(0)
        For i = 1 To UBound(parts)
            cur = cur & "\" & parts(i)
            If Not fso.FolderExists(cur) Then
                On Error Resume Next
                fso.CreateFolder cur
                If Err.Number <> 0 Then
                    Dim erroOriginal As Long, descOriginal As String
                    erroOriginal = Err.Number
                    descOriginal = Err.Description
                    On Error GoTo 0
                    Err.Raise erroOriginal, "EnsureFolder", _
                        descOriginal & " — falhou criando: """ & cur & """ (pasta completa pedida: """ & path & """)"
                End If
                On Error GoTo 0
            End If
        Next i
    End If
End Sub

' ---------- conversão automática de .htm para .xlsx ----------

' Varre toda a árvore a partir de rootDir procurando .htm/.html; pra cada
' um que encontrar, se ainda não existir um .xlsx de mesmo nome ao lado,
' abre e salva uma cópia .xlsx (não mexe no .htm original). Roda no
' Workbook_Open dos dois programas — rápido quando não há nada novo pra
' converter, porque só converte o que ainda não tem par .xlsx.
Sub ConverterHtmParaXlsx(ByVal rootDir As String)
    On Error Resume Next
    ConverterHtmParaXlsxRecursivo rootDir
    On Error GoTo 0
End Sub

Private Sub ConverterHtmParaXlsxRecursivo(ByVal dirPath As String)
    Dim fn As String
    On Error Resume Next
    fn = Dir(dirPath & "\*.*", vbDirectory)
    On Error GoTo 0
    Dim subdirs As New Collection
    Do While fn <> ""
        If fn <> "." And fn <> ".." Then
            Dim full As String
            full = dirPath & "\" & fn
            Dim attr As Long
            attr = 0
            On Error Resume Next
            attr = GetAttr(full)
            On Error GoTo 0
            If (attr And vbDirectory) = vbDirectory Then
                subdirs.Add full
            Else
                Dim dotPos As Long, ext As String
                dotPos = InStrRev(fn, ".")
                If dotPos > 0 Then
                    ext = LCase$(Mid$(fn, dotPos))
                    If ext = ".htm" Or ext = ".html" Then
                        Dim xlsxPath As String
                        xlsxPath = Left$(full, Len(full) - Len(fn) + dotPos - 1) & ".xlsx"
                        If Dir(xlsxPath) = "" Then
                            ConverterUmHtm full, xlsxPath
                        End If
                    End If
                End If
            End If
        End If
        fn = Dir()
    Loop
    Dim d As Variant
    For Each d In subdirs
        ConverterHtmParaXlsxRecursivo CStr(d)
    Next d
End Sub

Private Sub ConverterUmHtm(ByVal htmPath As String, ByVal xlsxPath As String)
    On Error Resume Next
    Dim wbConv As Workbook
    Set wbConv = Workbooks.Open(Filename:=htmPath, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False, Notify:=False)
    If Not wbConv Is Nothing Then
        Dim prevAlerts As Boolean
        prevAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        wbConv.SaveAs Filename:=xlsxPath, FileFormat:=51 ' xlOpenXMLWorkbook (.xlsx)
        Application.DisplayAlerts = prevAlerts
        wbConv.Close False
        Set wbConv = Nothing
    End If
    On Error GoTo 0
End Sub

' ---------- funções de cache (as variáveis Public ficam no topo do módulo) ----------

Sub ResetCaches()
    Set gMatriculaCache = CreateObject("Scripting.Dictionary")
    Set gMapaoCache = CreateObject("Scripting.Dictionary")
    ResetArvoreCache
End Sub

Function LoadMatriculaCached(ByVal path As String, ByVal turma As String) As Collection
    If gMatriculaCache Is Nothing Then ResetCaches
    If gMatriculaCache.Exists(path) Then
        Set LoadMatriculaCached = gMatriculaCache(path)
    Else
        Dim c As Collection
        Set c = LoadMatricula(path, turma)
        gMatriculaCache.Add path, c
        Set LoadMatriculaCached = c
    End If
End Function

Function LoadMapaoCached(ByVal path As String) As clsNotasFechamento
    If gMapaoCache Is Nothing Then ResetCaches
    If gMapaoCache.Exists(path) Then
        Set LoadMapaoCached = gMapaoCache(path)
    Else
        Dim nf As clsNotasFechamento
        Set nf = LoadMapao(path)
        gMapaoCache.Add path, nf
        Set LoadMapaoCached = nf
    End If
End Function

