Attribute VB_Name = "clsTimeline"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' A "linha do tempo" completa de um aluno: quem ele Ã© hoje + um clsYearRecord
' para cada ano do ciclo (1Âº ano atÃ© a sÃ©rie atual).
Public nome As String
Public anoAtual As Long
Public SerieAtual As Integer
Public turmaAtual As String
Public Cadastro As clsMatriculaRow     ' matrÃ­cula do ano atual
Public years As Collection             ' Collection de clsYearRecord, em ordem crescente de ano
Public warnings As Collection

Private Sub Class_Initialize()
    Set years = New Collection
    Set warnings = New Collection
End Sub
