Attribute VB_Name = "clsNotasFechamento"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Resultado de ler um arquivo de MapÃ£o (seja o "Conselho Final"/"Quinto
' Conceito" de um ano, seja o MapÃ£o de um bimestre especÃ­fico).
Public anoLetivo As Long
Public tipo As String          ' texto bruto de "Tipo de Fechamento"
Public isFinal As Boolean
Public bimestre As Integer     ' 1-4, ou 0 se for o fechamento final/desconhecido

' PorAluno: Dictionary normName(String) -> Collection de clsSubjectGrade
Public PorAluno As Object
' SituacaoPorAluno: Dictionary normName(String) -> texto da situaÃ§Ã£o naquele MapÃ£o
Public SituacaoPorAluno As Object

Private Sub Class_Initialize()
    Set PorAluno = CreateObject("Scripting.Dictionary")
    Set SituacaoPorAluno = CreateObject("Scripting.Dictionary")
End Sub
