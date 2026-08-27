Attribute VB_Name = "clsMatriculaRow"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Uma linha da planilha de MatrÃ­cula (Lista de alunos por classe).
Public nome As String
Public ra As String
Public DigRA As String
Public UFRA As String
Public Nascimento As String   ' jÃ¡ formatado dd/mm/aaaa
Public Situacao As String
Public Serie As Integer
Public turma As String        ' ex: "5A"
Public TipoEnsino As String
