#!/usr/bin/env python3
"""Gerador de jogos da Mega-Sena baseado em estatisticas dos sorteios.

Uso rapido:
    python megasena.py              menu interativo
    python megasena.py atualizar    baixa/atualiza o historico de sorteios
    python megasena.py estatisticas mostra frequencias e pares x impares
    python megasena.py gerar -n 5   gera 5 jogos
    python megasena.py backtest     testa a estrategia contra o historico

AVISO: nenhum metodo altera a probabilidade de acertar a sena. Todo jogo de
6 dezenas tem exatamente 1 chance em 50.063.860. Veja o README.
"""

import argparse
import csv
import json
import math
import random
import re
import sys
import urllib.request
from collections import Counter
from pathlib import Path

ARQUIVO_HISTORICO = Path(__file__).with_name("resultados.json")
URL_TODOS = "https://loteriascaixa-api.herokuapp.com/api/megasena"
URL_CAIXA = "https://servicebus2.caixa.gov.br/portaldeloterias/api/megasena/"
NUMEROS = range(1, 61)
JANELA_RECENTE = 50
TOTAL_COMBINACOES = math.comb(60, 6)


# --------------------------------------------------------------------------
# Historico de sorteios
# --------------------------------------------------------------------------

def carregar_historico():
    """Devolve a lista de sorteios [{concurso, data, dezenas}] em ordem."""
    if not ARQUIVO_HISTORICO.exists():
        return []
    with open(ARQUIVO_HISTORICO, encoding="utf-8") as f:
        return json.load(f)


def salvar_historico(sorteios):
    unicos = {s["concurso"]: s for s in sorteios}
    ordenados = [unicos[c] for c in sorted(unicos)]
    with open(ARQUIVO_HISTORICO, "w", encoding="utf-8") as f:
        json.dump(ordenados, f)
    return ordenados


def _baixar_json(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0", "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.load(resp)


def _normalizar(item):
    """Aceita tanto o formato da API da Caixa quanto o da API comunitaria."""
    concurso = item.get("concurso") or item.get("numero")
    dezenas = item.get("dezenas") or item.get("listaDezenas") or []
    dezenas = sorted(int(d) for d in dezenas)
    if not concurso or len(dezenas) != 6 or not all(1 <= d <= 60 for d in dezenas):
        return None
    return {
        "concurso": int(concurso),
        "data": item.get("data") or item.get("dataApuracao") or "",
        "dezenas": dezenas,
    }


def atualizar_historico():
    """Baixa os sorteios que faltam. Devolve o historico atualizado."""
    sorteios = carregar_historico()
    antes = len(sorteios)
    try:
        dados = _baixar_json(URL_TODOS)
        novos = [n for n in map(_normalizar, dados) if n]
        if not novos:
            raise ValueError("resposta sem sorteios")
        sorteios = salvar_historico(sorteios + novos)
    except Exception as erro:  # cai para a API oficial, concurso a concurso
        print(f"Fonte principal indisponivel ({erro}); usando a API da Caixa...")
        try:
            ultimo = _normalizar(_baixar_json(URL_CAIXA))
            tenho = {s["concurso"] for s in sorteios}
            faltam = [c for c in range(1, ultimo["concurso"]) if c not in tenho]
            novos = [ultimo]
            for i, concurso in enumerate(faltam, 1):
                item = _normalizar(_baixar_json(f"{URL_CAIXA}{concurso}"))
                if item:
                    novos.append(item)
                if i % 50 == 0:
                    print(f"  {i}/{len(faltam)} concursos baixados...")
                    salvar_historico(sorteios + novos)
            sorteios = salvar_historico(sorteios + novos)
        except Exception as erro2:
            print(f"Nao foi possivel baixar os resultados: {erro2}")
            print("Alternativa: baixe a planilha de resultados no site da Caixa,")
            print("salve como CSV e use:  python megasena.py importar arquivo.csv")
            return sorteios
    print(f"Historico: {len(sorteios)} sorteios ({len(sorteios) - antes} novos).")
    return sorteios


def importar_csv(caminho):
    """Importa um CSV com concurso, data e as 6 dezenas em cada linha."""
    texto = Path(caminho).read_text(encoding="utf-8-sig", errors="replace")
    delimitador = ";" if texto.count(";") > texto.count(",") else ","
    novos = []
    for linha in csv.reader(texto.splitlines(), delimiter=delimitador):
        campos = [c.strip() for c in linha]
        if len(campos) < 7 or not campos[0].isdigit():
            continue
        data = next((c for c in campos[1:3] if re.fullmatch(r"\d{1,2}/\d{1,2}/\d{4}", c)), "")
        inicio = 2 if data else 1
        dezenas = campos[inicio:inicio + 6]
        if all(d.isdigit() and 1 <= int(d) <= 60 for d in dezenas) and len(set(dezenas)) == 6:
            novos.append({"concurso": int(campos[0]), "data": data,
                          "dezenas": sorted(int(d) for d in dezenas)})
    if not novos:
        print("Nenhum sorteio reconhecido no arquivo.")
        return carregar_historico()
    sorteios = salvar_historico(carregar_historico() + novos)
    print(f"{len(novos)} sorteios lidos. Historico: {len(sorteios)} sorteios.")
    return sorteios


# --------------------------------------------------------------------------
# Estatisticas
# --------------------------------------------------------------------------

def frequencias(sorteios):
    cont = Counter({n: 0 for n in NUMEROS})
    for s in sorteios:
        cont.update(s["dezenas"])
    return cont


def pares_impares(sorteios):
    """Conta quantos sorteios tiveram cada divisao (pares, impares)."""
    cont = Counter()
    for s in sorteios:
        pares = sum(1 for d in s["dezenas"] if d % 2 == 0)
        cont[(pares, 6 - pares)] += 1
    return cont


def prob_teorica_pares(pares):
    """Probabilidade de um sorteio aleatorio ter exatamente `pares` pares."""
    return math.comb(30, pares) * math.comb(30, 6 - pares) / TOTAL_COMBINACOES


def fmt(dezenas):
    return " ".join(f"{d:02d}" for d in dezenas)


def mostrar_estatisticas(sorteios):
    if not sorteios:
        print("Sem historico. Rode primeiro:  python megasena.py atualizar")
        return
    total = len(sorteios)
    freq = frequencias(sorteios)
    recente = frequencias(sorteios[-JANELA_RECENTE:])
    print(f"\n{total} sorteios (concursos {sorteios[0]['concurso']} a {sorteios[-1]['concurso']})")
    print(f"Ultimo sorteio: {fmt(sorteios[-1]['dezenas'])}  {sorteios[-1]['data']}")

    esperado = total * 6 / 60
    print(f"\nDezenas que MAIS sairam (esperado por dezena: {esperado:.0f} vezes)")
    for n, q in freq.most_common(10):
        print(f"  {n:02d}: {q} vezes")
    print("Dezenas que MENOS sairam")
    for n, q in sorted(freq.items(), key=lambda x: x[1])[:10]:
        print(f"  {n:02d}: {q} vezes")
    print(f"Mais sorteadas nos ultimos {JANELA_RECENTE} concursos: "
          + fmt(sorted(n for n, _ in recente.most_common(10))))

    print("\nPares x impares        sorteios   observado   teorico")
    divisao = pares_impares(sorteios)
    for pares in range(7):
        q = divisao[(pares, 6 - pares)]
        print(f"  {pares} pares / {6 - pares} impares   {q:6d}     {q / total:6.1%}    "
              f"{prob_teorica_pares(pares):6.1%}")
    todas = [d for s in sorteios for d in s["dezenas"]]
    pares_total = sum(1 for d in todas if d % 2 == 0)
    print(f"  Dezenas sorteadas: {pares_total} pares ({pares_total / len(todas):.1%}) e "
          f"{len(todas) - pares_total} impares ({1 - pares_total / len(todas):.1%})")


# --------------------------------------------------------------------------
# Estrategia de geracao
# --------------------------------------------------------------------------

def pesos(sorteios):
    """Peso de cada dezena: 60% frequencia historica + 40% frequencia recente.

    Sem historico, todas as dezenas recebem o mesmo peso.
    """
    if not sorteios:
        return {n: 1.0 for n in NUMEROS}
    geral = frequencias(sorteios)
    recente = frequencias(sorteios[-JANELA_RECENTE:])
    max_g = max(geral.values()) or 1
    max_r = max(recente.values()) or 1
    # o termo fixo 0.25 garante que nenhuma dezena fique impossivel de sair
    return {n: 0.25 + 0.6 * geral[n] / max_g + 0.4 * recente[n] / max_r for n in NUMEROS}


def sortear_ponderado(peso, quantidade, rng):
    """Amostra sem reposicao proporcional aos pesos (Efraimidis-Spirakis)."""
    chaves = sorted(NUMEROS, key=lambda n: rng.random() ** (1.0 / peso[n]), reverse=True)
    return sorted(chaves[:quantidade])


def maior_sequencia(dezenas):
    maior = atual = 1
    for a, b in zip(dezenas, dezenas[1:]):
        atual = atual + 1 if b == a + 1 else 1
        maior = max(maior, atual)
    return maior


def jogo_valido(dezenas, ja_sorteados=frozenset()):
    """Filtros de equilibrio: o jogo deve 'parecer' um sorteio tipico."""
    qtd = len(dezenas)
    pares = sum(1 for d in dezenas if d % 2 == 0)
    if abs(pares - (qtd - pares)) > 2:               # pares x impares equilibrados
        return False
    desvio = math.sqrt(qtd * 299.92 * (60 - qtd) / 59)
    if abs(sum(dezenas) - qtd * 30.5) > 1.5 * desvio:  # soma na faixa mais comum
        return False
    if qtd <= 8 and maior_sequencia(dezenas) > 2:    # no maximo 2 consecutivas
        return False
    if max(Counter((d - 1) // 10 for d in dezenas).values()) > math.ceil(qtd / 2):
        return False                                 # nao concentrar numa dezena
    if sum(1 for d in dezenas if d > 31) < 2:        # evita jogo "de aniversario"
        return False
    if tuple(dezenas) in ja_sorteados:               # sena repetida nunca ocorreu
        return False
    return True


def gerar_jogos(sorteios, quantidade=1, dezenas=6, rng=None):
    """Gera jogos ponderados pela frequencia, equilibrados e pouco sobrepostos."""
    rng = rng or random.Random()
    peso = pesos(sorteios)
    ja_sorteados = frozenset(tuple(s["dezenas"]) for s in sorteios)
    max_comum = max(2, dezenas // 3)  # jogos diferentes entre si cobrem mais
    jogos = []
    while len(jogos) < quantidade:
        candidato = None
        for tentativa in range(5000):
            candidato = sortear_ponderado(peso, dezenas, rng)
            if not jogo_valido(candidato, ja_sorteados):
                continue
            # depois de muitas tentativas, relaxa a exigencia de diversidade
            if tentativa > 2500 or all(len(set(candidato) & set(j)) <= max_comum for j in jogos):
                break
        jogos.append(candidato)
    return jogos


def mostrar_jogos(sorteios, quantidade, dezenas):
    if not sorteios:
        print("Sem historico: os jogos serao equilibrados, mas sem usar frequencias.")
        print("Rode  python megasena.py atualizar  para baixar os sorteios.")
    jogos = gerar_jogos(sorteios, quantidade, dezenas)
    print(f"\n{quantidade} jogo(s) de {dezenas} dezenas:")
    for i, jogo in enumerate(jogos, 1):
        pares = sum(1 for d in jogo if d % 2 == 0)
        print(f"  Jogo {i:2d}: {fmt(jogo)}   ({pares} pares / {dezenas - pares} impares, soma {sum(jogo)})")
    apostas = quantidade * math.comb(dezenas, 6)
    print(f"\nChance de sena com esses jogos: 1 em {TOTAL_COMBINACOES / apostas:,.0f}".replace(",", "."))
    print("(a mesma de quaisquer outros jogos distintos na mesma quantidade)")


# --------------------------------------------------------------------------
# Backtest: a estrategia e melhor que o acaso?
# --------------------------------------------------------------------------

def backtest(sorteios, concursos=300, jogos_por_concurso=10, semente=1):
    """Para cada concurso, gera jogos usando SO os sorteios anteriores e
    compara os acertos com jogos totalmente aleatorios."""
    if len(sorteios) < 100:
        print("Historico insuficiente para o teste (minimo 100 sorteios).")
        return None
    concursos = min(concursos, len(sorteios) - 60)
    rng = random.Random(semente)
    acertos = {"estrategia": Counter(), "aleatorio": Counter()}
    for i in range(len(sorteios) - concursos, len(sorteios)):
        resultado = set(sorteios[i]["dezenas"])
        for jogo in gerar_jogos(sorteios[:i], jogos_por_concurso, 6, rng):
            acertos["estrategia"][len(resultado & set(jogo))] += 1
        for _ in range(jogos_por_concurso):
            acertos["aleatorio"][len(resultado & set(rng.sample(NUMEROS, 6)))] += 1

    total = concursos * jogos_por_concurso
    print(f"\nBacktest: {concursos} concursos x {jogos_por_concurso} jogos = {total} jogos por metodo")
    print("                 media de acertos   ternos+   quadras   quinas   senas")
    for nome, cont in acertos.items():
        media = sum(k * v for k, v in cont.items()) / total
        ternos = sum(v for k, v in cont.items() if k >= 3)
        print(f"  {nome:<12}        {media:.3f}        {ternos:5d}    {cont[4]:5d}    {cont[5]:5d}   {cont[6]:5d}")
    print("  esperado pelo acaso   0.600")
    print("\nSe as duas linhas ficam parecidas, a estrategia nao supera o acaso -")
    print("e isso que a matematica preve para um sorteio honesto.")
    return acertos


# --------------------------------------------------------------------------
# Interface
# --------------------------------------------------------------------------

def ler_int(pergunta, padrao, minimo, maximo):
    resposta = input(f"{pergunta} [{padrao}]: ").strip()
    try:
        return max(minimo, min(maximo, int(resposta))) if resposta else padrao
    except ValueError:
        return padrao


def menu():
    sorteios = carregar_historico()
    while True:
        print(f"\n=== MEGA-SENA ===  ({len(sorteios)} sorteios no historico)")
        print("1) Atualizar historico de sorteios")
        print("2) Ver estatisticas (mais sorteadas, pares x impares)")
        print("3) Gerar jogos")
        print("4) Testar a estrategia contra o historico (backtest)")
        print("0) Sair")
        opcao = input("Opcao: ").strip()
        if opcao == "1":
            sorteios = atualizar_historico()
        elif opcao == "2":
            mostrar_estatisticas(sorteios)
        elif opcao == "3":
            mostrar_jogos(sorteios, ler_int("Quantos jogos", 3, 1, 100),
                          ler_int("Dezenas por jogo (6 a 15)", 6, 6, 15))
        elif opcao == "4":
            backtest(sorteios)
        elif opcao == "0":
            break


def main():
    parser = argparse.ArgumentParser(description="Gerador de jogos da Mega-Sena")
    sub = parser.add_subparsers(dest="comando")
    sub.add_parser("atualizar", help="baixa/atualiza o historico de sorteios")
    sub.add_parser("estatisticas", help="mostra frequencias e pares x impares")
    p = sub.add_parser("gerar", help="gera jogos")
    p.add_argument("-n", "--jogos", type=int, default=3)
    p.add_argument("-d", "--dezenas", type=int, default=6, choices=range(6, 16), metavar="6-15")
    p = sub.add_parser("backtest", help="compara a estrategia com jogos aleatorios")
    p.add_argument("-c", "--concursos", type=int, default=300)
    p.add_argument("-n", "--jogos", type=int, default=10)
    p = sub.add_parser("importar", help="importa sorteios de um arquivo CSV")
    p.add_argument("arquivo")
    args = parser.parse_args()

    if args.comando is None:
        menu()
    elif args.comando == "atualizar":
        atualizar_historico()
    elif args.comando == "estatisticas":
        mostrar_estatisticas(carregar_historico())
    elif args.comando == "gerar":
        mostrar_jogos(carregar_historico(), max(1, args.jogos), args.dezenas)
    elif args.comando == "backtest":
        backtest(carregar_historico(), args.concursos, args.jogos)
    elif args.comando == "importar":
        importar_csv(args.arquivo)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(0)
