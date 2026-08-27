Attribute VB_Name = "clsSubjectGrade"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Uma nota de uma disciplina para um aluno, lida de um MapÃ£o.
Public Subject As String   ' jÃ¡ normalizado (ver NormSubject)
Public Value As String     ' nÃºmero, "-" ou cÃ³digo de menÃ§Ã£o (ET/ES/EP), como estÃ¡ na origem
