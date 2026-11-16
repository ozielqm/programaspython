# Gerador de jogos da Mega-Sena

Programa em Python (sem dependências externas) que baixa o histórico de
sorteios, mostra estatísticas e gera jogos equilibrados.

```
python megasena.py               # menu interativo
python megasena.py atualizar     # baixa/atualiza o histórico (resultados.json)
python megasena.py estatisticas  # mais sorteadas, pares x ímpares
python megasena.py gerar -n 5    # 5 jogos de 6 dezenas
python megasena.py gerar -n 2 -d 8   # 2 jogos de 8 dezenas
python megasena.py backtest      # estratégia x jogos aleatórios
python megasena.py importar resultados.csv   # se o download falhar
```

O CSV de importação deve ter, em cada linha: concurso, data e as 6 dezenas
(é o formato da planilha de resultados do site da Caixa salva como CSV).

## A estratégia

1. **Peso por frequência** – cada dezena recebe um peso: 60% pela frequência
   em todo o histórico e 40% pela frequência nos últimos 50 concursos.
2. **Pares x ímpares** – só passam jogos com 3/3, 4/2 ou 2/4.
3. **Soma na faixa central** – a soma das dezenas fica na faixa mais comum.
4. **Sem padrões** – no máximo 2 dezenas consecutivas e sem concentração em
   uma mesma dezena (01–10, 11–20...).
5. **Pelo menos 2 dezenas acima de 31** – evita jogos "de aniversário".
6. **Jogos diferentes entre si** – no máximo 2 dezenas repetidas entre jogos.

## O que isso muda de verdade

**A chance de acertar a sena não muda.** Cada combinação de 6 dezenas tem
exatamente 1 chance em 50.063.860, e o sorteio não tem memória: dezenas
"quentes" não têm mais chance de sair de novo. A divisão 3 pares / 3 ímpares
sai em ~33% dos sorteios apenas porque existem mais combinações desse tipo,
não porque cada uma delas seja mais provável. O comando `backtest` mostra
isso na prática: a estratégia acerta o mesmo que jogos aleatórios.

O que realmente tem efeito:

- **Mais combinações** – a única forma de aumentar a chance é apostar mais
  jogos distintos ou mais dezenas (bolão). 10 jogos = 10 em 50.063.860.
- **Fugir de jogos populares** (filtros 4 e 5) – não aumenta a chance de
  ganhar, mas diminui a de **dividir o prêmio**, porque muita gente aposta
  datas, sequências e desenhos no volante.

Jogue apenas o que pode perder.
