Attribute VB_Name = "clsTurmaFile"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Um arquivo de turma encontrado numa pasta de ano (matrÃ­cula ou MapÃ£o).
Public grade As Integer    ' 1-9
Public letter As String    ' "A", "B", ... ou "" se a turma nÃ£o Ã© dividida por letra
Public path As String      ' caminho completo do arquivo
Public ext As String       ' extensÃ£o em minÃºsculas, com ponto (".xlsx", ".xls")
