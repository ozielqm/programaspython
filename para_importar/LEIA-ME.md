# Módulos prontos para importar

Estes três `.bas` já estão com todas as correções aplicadas e gravados em
**ANSI (cp1252)** — que é como o editor do VBA lê arquivos importados. Se você
importar um `.bas` salvo em UTF-8, todo acento vira lixo (`"não"` → `"nÃ£o"`);
por isso não use os arquivos da pasta `vba/` para importar: aqueles são a
versão de leitura/histórico, em UTF-8.

## Passo a passo (5 minutos)

Abra o `.xlsm` → `Alt+F11`.

### 1. `modAutoPreenche` — módulo novo

`Arquivo > Importar Arquivo...` → `modAutoPreenche.bas`.

Nada a remover: esse módulo não existia antes.

### 2. `modPreenchimento` e `modNucleo` — substituem os antigos

Para cada um dos dois:

1. No painel da esquerda, clique com o botão direito no módulo de mesmo nome →
   **Remover**. Quando perguntar "Deseja exportar antes de remover?", clique
   **Não** (a versão antiga já está guardada no repositório).
2. `Arquivo > Importar Arquivo...` → o `.bas` correspondente desta pasta.

> Se importar **sem** remover antes, o VBA cria `modNucleo1` e o projeto não
> compila. Remover primeiro é o que evita isso.

### 3. `frmHistorico` — única edição à mão

O formulário não pode ser importado (o layout mora num `.frx` que eu não tenho),
então essa parte continua manual. Dê duplo clique em `frmHistorico` → `Exibir
código`, ache `Private Sub cmdBuscarRA_Click()`, e troque estas três linhas:

```vba
    txtMunicipioNasc.Value = ""
    txtEstadoNasc.Value = ""
    txtPaisNasc.Value = ""
```

por estas seis:

```vba
    Dim natAuto As clsNaturalidade
    Dim origemNat As String
    Set natAuto = BuscarNaturalidadeAuto(tl, anoConclusao, origemNat)
    txtMunicipioNasc.Value = natAuto.Municipio
    txtEstadoNasc.Value = natAuto.Estado
    txtPaisNasc.Value = natAuto.pais
```

Depois, uns dez linhas abaixo, logo após a linha que começa com
`AppendLog txtLog, "Aluno localizado por RA: "`, acrescente:

```vba
    AppendLog txtLog, origemNat
```

(essa última só serve para o log dizer de onde veio a naturalidade — se pular,
o resto funciona igual.)

### 4. Compilar e salvar

`Depurar > Compilar VBAProject`. Não deve acusar nada. Salve o `.xlsm`.

## Como saber se deu certo

Busque um aluno por RA que esteja no `naturalidade.xlsx`. Município, Estado e
País devem aparecer preenchidos sozinhos, e o log deve dizer:

```
Naturalidade preenchida automaticamente de naturalidade.xlsx - confira antes de gerar.
```

Depois digite `1000/80` na carga horária de um ano anterior a 2020 e de um de
2020 em diante, gere, e confira no arquivo de saída:

- linha 36 (`TOTAL DE AULAS DA BASE NACIONAL COMUM CURRICULAR`) = 1000
- linha 51 (`TOTAL DE AULAS DA PARTE DIVERSIFICADA`) = 80
- linha 53 (`TOTAL DA CARGA HORÁRIA ANUAL`) = `=(G52*50/60)` na coluna do ano
  antigo e `=(H52*45/60)` na do ano de 2020 em diante

## O que mudou em cada módulo

| módulo | mudança |
|---|---|
| `modAutoPreenche` | **novo** — naturalidade automática, carga horária por rótulo, fator 50/45, e a busca da pasta `HISTÓRICOS` |
| `modPreenchimento` | `FillGradeTables` chama as duas rotinas novas em vez de escrever na linha 53 fixa |
| `modNucleo` | `BaseDir()` usa `PastaHistoricosLocal()` no lugar da constante com perfil de usuário fixo |
| `frmHistorico` | preenche município/UF/país na busca por RA (edição à mão) |

Detalhes de cada mudança: `vba/INSTRUCOES_ATUALIZACAO.md`.
