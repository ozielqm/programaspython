Attribute VB_Name = "clsYearRecord"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Um dos anos do ciclo do aluno (1º ano até a série atual).
Public ano As Long
Public Position As Integer     ' 1..5
Public turma As String
Public mat As clsMatriculaRow  ' matrícula do aluno NESSE ano especificamente (pode ser Nothing)
Public FinalNotas As clsNotasFechamento    ' Conselho Final desse ano, se encontrado (pode ser Nothing)
Public Bimestres As Object     ' Dictionary: bimestre(1-4, como texto) -> clsNotasFechamento
Public warnings As Collection
Public ManualNotas As Object   ' Dictionary: chave da disciplina -> nota digitada a mão (só usado quando FinalNotas Is Nothing)
Public CargaHoraria As String  ' carga horária anual digitada a mão pro Concluinte (em branco = mantém o padrão do modelo)

Private Sub Class_Initialize()
    Set Bimestres = CreateObject("Scripting.Dictionary")
    Set warnings = New Collection
    Set ManualNotas = CreateObject("Scripting.Dictionary")
    CargaHoraria = ""
End Sub

