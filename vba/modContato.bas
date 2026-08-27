Attribute VB_Name = "modContato"
Option Explicit

' =====================================================================
' modContato - solicitacao de historico de transferencia pendente.
' So monta e permite enviar e-mail/WhatsApp quando realmente falta
' matricula ou Mapao de fechamento final de algum ano ANTERIOR ao ano
' atual do aluno nas pastas de origem (mesma checagem que o resto do
' programa ja faz ao montar a timeline em BuildTimeline). Se os dados ja
' estiverem completos, o envio e bloqueado.
' =====================================================================

Function NotasFaltando(ByVal ra As String, ByVal nomeFallback As String, ByRef nomeAluno As String, ByRef detalhe As String) As Boolean
    Dim errMsg As String
    Dim tl As clsTimeline
    If Trim$(ra) <> "" Then Set tl = BuildTimelineByRA(ra, errMsg)
    If tl Is Nothing And Trim$(nomeFallback) <> "" Then
        Dim errMsg2 As String
        Set tl = BuildTimelineByName(nomeFallback, errMsg2)
        If tl Is Nothing And errMsg = "" Then errMsg = errMsg2
    End If
    If tl Is Nothing Then
        detalhe = "nao consegui localizar esse aluno para verificar (" & errMsg & ")"
        NotasFaltando = False
        Exit Function
    End If
    nomeAluno = tl.nome

    Dim faltando As String
    faltando = ""
    Dim yr As clsYearRecord
    For Each yr In tl.years
        If yr.ano < tl.anoAtual Then
            If yr.mat Is Nothing Then
                faltando = faltando & "- matricula de " & yr.ano & " nao encontrada" & vbCrLf
            ElseIf yr.FinalNotas Is Nothing Then
                faltando = faltando & "- Mapao de fechamento final de " & yr.ano & " (turma " & yr.turma & ") nao encontrado" & vbCrLf
            End If
        End If
    Next yr
    detalhe = faltando
    NotasFaltando = (faltando <> "")
End Function

' ---------- montagem das mensagens ----------

Function MontarAssuntoEscola(ByVal nomeAluno As String) As String
    MontarAssuntoEscola = "Solicitacao de Historico Escolar de Transferencia - " & nomeAluno
End Function

Function MontarCorpoEscola(ByVal nomeAluno As String, ByVal detalheFaltando As String) As String
    MontarCorpoEscola = "Prezados(as)," & vbCrLf & vbCrLf & _
        "A " & EscolaNome & " vem, por meio deste, solicitar o envio do Historico Escolar de Transferencia do(a) aluno(a) " & _
        nomeAluno & ", atualmente matriculado(a) nesta unidade." & vbCrLf & vbCrLf & _
        "Verificamos que ainda nao recebemos os seguintes dados:" & vbCrLf & detalheFaltando & vbCrLf & _
        "Solicitamos gentilmente o envio da documentacao para este e-mail o quanto antes, para regularizacao da matricula." & vbCrLf & vbCrLf & _
        "Agradecemos a atencao." & vbCrLf & vbCrLf & _
        "Atenciosamente," & vbCrLf & "Secretaria - " & EscolaNome
End Function

Function MontarAssuntoPais(ByVal nomeAluno As String) As String
    MontarAssuntoPais = "Historico Escolar de Transferencia pendente - " & nomeAluno
End Function

Function MontarMensagemPais(ByVal nomeAluno As String, ByVal detalheFaltando As String) As String
    MontarMensagemPais = "Prezados pais/responsaveis por " & nomeAluno & "," & vbCrLf & vbCrLf & _
        "Ainda nao recebemos, na " & EscolaNome & ", o Historico Escolar de Transferencia do(a) aluno(a), documento que precisa ser solicitado a escola anterior. Faltam os seguintes dados:" & vbCrLf & detalheFaltando & vbCrLf & _
        "Pedimos a gentileza de entrar em contato com a escola anterior e solicitar o envio desse documento para o e-mail oficial da " & EscolaNome & " o quanto antes, para regularizar a matricula." & vbCrLf & vbCrLf & _
        "Agradecemos a atencao." & vbCrLf & vbCrLf & _
        "Secretaria - " & EscolaNome
End Function

' ---------- envio de e-mail via Outlook ----------

' Abre um e-mail ja pronto (destinatario/assunto/corpo) no Outlook para
' revisao antes de enviar. Se quiser enviar direto sem revisao manual,
' troque a linha "mail.Display" por "mail.Send" logo abaixo.
Function EnviarEmailOutlook(ByVal destinatario As String, ByVal assunto As String, ByVal corpo As String, ByRef errMsg As String) As Boolean
    On Error GoTo falhou
    Dim ol As Object, mail As Object
    Set ol = CreateObject("Outlook.Application")
    Set mail = ol.CreateItem(0) ' olMailItem
    mail.To = destinatario
    mail.Subject = assunto
    mail.Body = corpo
    mail.Display
    EnviarEmailOutlook = True
    Exit Function
falhou:
    errMsg = "Erro " & Err.Number & ": " & Err.Description & " (o Outlook precisa estar instalado/configurado nesta maquina)"
    EnviarEmailOutlook = False
End Function

' ---------- WhatsApp ----------

' Mantem so digitos e garante o prefixo do Brasil (55) se nao vier com
' codigo de pais. Nao valida DDD nem quantidade de digitos alem disso -
' confira o numero antes de enviar.
Function NormalizaTelefoneWhatsApp(ByVal tel As String) As String
    Dim d As String
    d = DigitsOnly(tel)
    If Len(d) <= 11 Then d = "55" & d
    NormalizaTelefoneWhatsApp = d
End Function

' Abre o WhatsApp (Web ou Desktop, o que estiver associado a wa.me) ja
' com a conversa e a mensagem prontas - o envio final e manual, clicando
' no botao de enviar dentro do proprio WhatsApp.
Sub AbrirWhatsApp(ByVal telefone As String, ByVal mensagem As String)
    Dim tel As String
    tel = NormalizaTelefoneWhatsApp(telefone)
    Dim url As String
    url = "https://wa.me/" & tel & "?text=" & EncodeURIComponent(mensagem)
    ThisWorkbook.FollowHyperlink url
End Sub

' Percent-encoding via UTF-8 (usa ADODB.Stream para converter acentos
' corretamente antes de codificar).
Function EncodeURIComponent(ByVal s As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open
    stream.WriteText s
    stream.Position = 0
    stream.Type = 1 ' adTypeBinary
    Dim bytes() As Byte
    bytes = stream.Read
    stream.Close

    Dim out As String, i As Long, b As Long, startAt As Long
    out = ""
    startAt = 0
    ' ADODB.Stream com Charset UTF-8 grava um BOM (EF BB BF) no comeco -
    ' pula esses 3 bytes pra nao aparecer como lixo no inicio da mensagem.
    If UBound(bytes) >= 2 Then
        If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then startAt = 3
    End If
    For i = startAt To UBound(bytes)
        b = bytes(i)
        Select Case b
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126 ' 0-9 A-Z a-z - . _ ~
                out = out & Chr$(b)
            Case Else
                out = out & "%" & Right$("0" & Hex$(b), 2)
        End Select
    Next i
    EncodeURIComponent = out
End Function




