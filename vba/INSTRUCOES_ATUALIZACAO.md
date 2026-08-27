# Atualização do GERADOR_HISTORICO_FORMANDOS_5_ANOS.xlsm

Três mudanças, todas concentradas num módulo novo (`modAutoPreenche.bas`) + duas
edições pequenas em código que já existia.

## O que muda

### 1. Município, UF e País voltam a ser automáticos

Antes, `cmdBuscarRA_Click` limpava os três campos e você digitava tudo à mão.
Agora eles são preenchidos na busca por RA. O programa procura, nesta ordem:

1. **Colunas de nascimento na própria planilha de matrícula do ano** — aceita
   `Município de Nascimento`, `UF de Nascimento`, `País de Nascimento`, e
   também uma coluna única `Naturalidade`. A comparação ignora acento e
   maiúscula/minúscula.
   O `UF` só é aceito se o cabeçalho também tiver "nasc" — senão a coluna
   `UF RA` (que é o estado do RA, não o de nascimento) seria pega por engano.
2. **Um arquivo `naturalidade.xlsx`** (ou `.htm`) em qualquer lugar dentro da
   pasta de trabalho, até 3 níveis abaixo, com colunas `Nome`, `Municipio`,
   `Estado`, `Pais`. Esse formato já era lido pela função `LoadNaturalidade`,
   em `modLeitura` — ela existia mas ninguém chamava.

Se o município vier preenchido e o país não, o país vira `BRASIL` (era o que o
`FillCadastro` já fazia na hora de gravar).

Se nenhuma das duas fontes tiver o aluno, os campos ficam em branco e o log
diz exatamente onde foi procurado — aí é digitar à mão, como antes.

### 2. Carga horária vai para as linhas de TOTAL DE AULAS

O valor digitado em "carga horária" de cada ano passa a ser escrito em:

- `TOTAL DE AULAS DA BASE NACIONAL COMUM CURRICULAR`
- `TOTAL DE AULAS DA PARTE DIVERSIFICADA`

na coluna do ano correspondente (G a K). Antes ia para a linha 53, fixa no
código — que é justamente a linha da fórmula de horas (ver item 3), ou seja,
o valor digitado apagava a fórmula.

As linhas são localizadas **pelo texto do rótulo**, não por número: se o
modelo mudar de layout, o código continua achando.

Um valor único vai para as duas linhas. Para valores diferentes, digite os
dois separados por barra:

| você digita | BNCC | Parte Diversificada |
|---|---|---|
| `800`     | 800 | 800 |
| `800/200` | 800 | 200 |

(`;` e `|` funcionam como separador também.)

### 3. TOTAL DA CARGA HORÁRIA ANUAL (Horas): 50/60 até 2019, 45/60 de 2020 em diante

Essa linha continua sendo **fórmula** — nunca mais é sobrescrita por um número
digitado. O que muda sozinho é o fator, conforme o ano daquela coluna:

| ano da coluna | fórmula escrita |
|---|---|
| até 2019      | `=(G52*50/60)` |
| 2020 em diante| `=(G52*45/60)` |

Um aluno que estudou de 2018 a 2022 fica com colunas mistas — cada coluna usa
o fator do seu próprio ano. Nada a configurar.

A linha referenciada (o `52`) **não é fixa no código**: sai da própria fórmula
que já está no modelo naquela célula; se a célula estiver vazia, cai para a
linha rotulada `TOTAL DE AULAS` (comparação exata, para não confundir com
`TOTAL DE AULAS DA BASE NACIONAL COMUM CURRICULAR`).

## Como instalar

1. Abra o `.xlsm`, `Alt+F11` (editor do VBA).
2. `Arquivo > Importar Arquivo...` → escolha **`modAutoPreenche.bas`**.
3. Em **`modPreenchimento`**, dentro de `FillGradeTables`, troque este bloco:

   ```vba
        If isConc And yr.CargaHoraria <> "" Then
            Dim chNum As String
            chNum = Replace(yr.CargaHoraria, ",", ".")
            If IsNumeric(chNum) Then
                ws.Range(col & cCargaHorariaRow).Value = CDbl(chNum)
            Else
                ws.Range(col & cCargaHorariaRow).Value = yr.CargaHoraria
            End If
        End If
   ```

   por este:

   ```vba
        If isConc Then
            AplicarCargaHoraria ws, col, yr.CargaHoraria, warnings
            AplicarHorasAnuais ws, col, yr.ano, warnings
        End If
   ```

   (a constante `cCargaHorariaRow = 53`, lá no topo do módulo, fica sem uso —
   pode apagar ou deixar.)

4. Em **`frmHistorico`** (código do formulário), dentro de `cmdBuscarRA_Click`,
   troque estas três linhas:

   ```vba
    txtMunicipioNasc.Value = ""
    txtEstadoNasc.Value = ""
    txtPaisNasc.Value = ""
   ```

   por estas:

   ```vba
    Dim natAuto As clsNaturalidade
    Dim origemNat As String
    Set natAuto = BuscarNaturalidadeAuto(tl, anoConclusao, origemNat)
    txtMunicipioNasc.Value = natAuto.Municipio
    txtEstadoNasc.Value = natAuto.Estado
    txtPaisNasc.Value = natAuto.pais
   ```

   e, mais abaixo, depois de `AppendLog txtLog, "Aluno localizado por RA: ..."`,
   acrescente:

   ```vba
    AppendLog txtLog, origemNat
   ```

5. `Depurar > Compilar VBAProject` e salve.

Os arquivos deste diretório já estão com os passos 3 e 4 aplicados — dá para
comparar lado a lado se ficar dúvida.

## Como testar

1. Busque um aluno por RA e confira o log: ele diz de onde a naturalidade veio
   (`da matricula (arquivo 5A.xlsx)`, `de naturalidade.xlsx`) ou por que não
   achou.
2. Digite `800/200` na carga horária de um ano anterior a 2020 e de um ano de
   2020 em diante, gere, e abra o arquivo de saída:
   - as duas linhas de `TOTAL DE AULAS` com 800 e 200 na coluna certa;
   - a linha de horas com `=(G52*50/60)` na coluna do ano antigo e
     `=(H52*45/60)` na coluna do ano novo.

## Observação sobre acentos no código

`modPreenchimento.bas` já tinha, dentro do próprio `.xlsm`, 93 acentos
perdidos (`"n?o encontrei"`, `"Map?o"`, `"s?rie"`) — provavelmente de uma
importação antiga com a codificação errada. Não mexi nisso aqui, mas as
mensagens aparecem assim para o usuário no log; vale corrigir num passe
separado.

Por causa disso, `modAutoPreenche.bas` foi escrito **sem nenhum acento**: o
editor do VBA importa `.bas` como ANSI, e acento em arquivo UTF-8 vira lixo na
importação.
