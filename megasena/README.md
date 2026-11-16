# Gerador de jogos da Mega-Sena e da Lotofácil

Programa em Python (sem dependências externas) que baixa o histórico de
sorteios, mostra estatísticas e gera jogos equilibrados. Acrescente
`-j lotofacil` antes do comando para usar a Lotofácil.

```
python megasena.py               # menu interativo (opção 5 troca de loteria)
python megasena.py atualizar     # baixa/atualiza o histórico
python megasena.py estatisticas  # mais sorteadas, pares x ímpares
python megasena.py gerar -n 5    # 5 jogos de 6 dezenas
python megasena.py gerar -n 2 -d 8   # 2 jogos de 8 dezenas
python megasena.py backtest      # estratégia x jogos aleatórios
python megasena.py importar resultados.csv   # se o download falhar

python megasena.py -j lotofacil atualizar
python megasena.py -j lotofacil gerar -n 5        # 5 jogos de 15 dezenas
python megasena.py -j lotofacil gerar -n 2 -d 17  # 2 jogos de 17 dezenas
```

O histórico fica em `resultados.json` (Mega-Sena) e
`resultados_lotofacil.json` (Lotofácil). O CSV de importação deve ter, em
cada linha: concurso, data e as dezenas sorteadas (é o formato da planilha
de resultados do site da Caixa salva como CSV).

## A estratégia

1. **Peso por frequência** – cada dezena recebe um peso: 60% pela frequência
   em todo o histórico e 40% pela frequência nos últimos 50 concursos.
2. **Pares x ímpares** – só passam jogos perto da proporção do volante:
   Mega-Sena 3/3, 4/2 ou 2/4; Lotofácil de 6 a 8 pares (7 a 9 ímpares).
3. **Soma na faixa central** – a soma das dezenas fica na faixa mais comum.
4. **Sem padrões** – Mega-Sena: no máximo 2 dezenas consecutivas e sem
   concentração em uma mesma dezena (01–10, 11–20...). Lotofácil: no máximo
   6 consecutivas e nenhuma linha do volante vazia.
5. **Pelo menos 2 dezenas acima de 31** (só Mega-Sena) – evita jogos "de
   aniversário".
6. **Jogos diferentes entre si** – Mega-Sena: no máximo 2 dezenas repetidas
   entre jogos; Lotofácil: no máximo 10.

## O que isso muda de verdade

**A chance de acertar não muda.** Cada jogo simples tem a mesma chance
(Mega-Sena: 1 em 50.063.860; Lotofácil: 1 em 3.268.760), e o sorteio não tem
memória: dezenas "quentes" não têm mais chance de sair de novo. As divisões
equilibradas entre pares e ímpares saem mais apenas porque existem mais
combinações desse tipo, não porque cada uma delas seja mais provável. O
comando `backtest` mostra isso na prática: a estratégia acerta o mesmo que
jogos aleatórios.

O que realmente tem efeito:

- **Mais combinações** – a única forma de aumentar a chance é apostar mais
  jogos distintos ou mais dezenas (bolão).
- **Fugir de jogos populares** (filtros 4 e 5) – não aumenta a chance de
  ganhar, mas diminui a de **dividir o prêmio**, porque muita gente aposta
  datas, sequências e desenhos no volante.

Jogue apenas o que pode perder.

## Versão para celular e computador

`megasena.html` é o mesmo gerador em um único arquivo, sem instalar nada:
basta abrir no navegador (Chrome, Safari, Edge...) do computador ou do
celular e escolher Mega-Sena ou Lotofácil nas abas do topo. O histórico
baixado fica salvo no próprio aparelho.
