#!/usr/bin/env python3
"""Gerador de jogos da Mega-Sena e da Lotofacil baseado em estatisticas.

Uso rapido (acrescente  -j lotofacil  para a Lotofacil):
    python megasena.py                        menu interativo
    python megasena.py atualizar              baixa/atualiza o historico
    python megasena.py estatisticas           frequencias e pares x impares
    python megasena.py gerar -n 5             gera 5 jogos
    python megasena.py backtest               testa a estrategia no historico
    python megasena.py -j lotofacil gerar -n 5

AVISO: nenhum metodo altera a probabilidade de acertar. Todo jogo simples
tem a mesma chance (Mega-Sena: 1 em 50.063.860; Lotofacil: 1 em 3.268.760).
Veja o README.
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

PASTA = Path(__file__).parent
URL_TODOS = "https://loteriascaixa-api.herokuapp.com/api/"
URL_CAIXA = "https://servicebus2.caixa.gov.br/portaldeloterias/api/"
JANELA_RECENTE = 50

# total: dezenas no volante; sorteadas: dezenas por sorteio;
# seq: (maior sequencia de consecutivas aceita, ate quantas dezenas por jogo);
# faixa: tamanho da faixa/linha do volante; min_altas: minimo de dezenas > 31;
# acertos: quantidades de acertos mostradas no backtest.
JOGOS = {
    "megasena": dict(nome="Mega-Sena", total=60, sorteadas=6, aposta_max=15,
                     arquivo="resultados.json", faixa=10, faixa_regra="max_metade",
                     seq=(2, 8), min_altas=2, acertos=(3, 4, 5, 6)),
    "lotofacil": dict(nome="Lotofacil", total=25, sorteadas=15, aposta_max=20,
                      arquivo="resultados_lotofacil.json", faixa=5, faixa_regra="sem_vazia",
                      seq=(6, 16), min_altas=0, acertos=(11, 12, 13, 14, 15)),
}
for _chave, _cfg in JOGOS.items():
    _cfg["chave"] = _chave
    _cfg["combinacoes"] = math.comb(_cfg["total"], _cfg["sorteadas"])


def numeros(cfg):
    return range(1, cfg["total"] + 1)


def ponto(n):
    """Formata inteiro com ponto de milhar."""
    return f"{n:,.0f}".replace(",", ".")


# --------------------------------------------------------------------------
# Historico de sorteios
# --------------------------------------------------------------------------

def carregar_historico(cfg):
    """Devolve a lista de sorteios [{concurso, data, dezenas}] em ordem."""
    arquivo = PASTA / cfg["arquivo"]
    if not arquivo.exists():
        return []
    with open(arquivo, encoding="utf-8") as f:
        return json.load(f)


def salvar_historico(cfg, sorteios):
    unicos = {s["concurso"]: s for s in sorteios}
    ordenados = [unicos[c] for c in sorted(unicos)]
    with open(PASTA / cfg["arquivo"], "w", encoding="utf-8") as f:
        json.dump(ordenados, f)
    return ordenados


def _baixar_json(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0", "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.load(resp)


def _normalizar(cfg, item):
    """Aceita tanto o formato da API da Caixa quanto o da API comunitaria."""
    concurso = item.get("concurso") or item.get("numero")
    dezenas = item.get("dezenas") or item.get("listaDezenas") or []
    dezenas = sorted(int(d) for d in dezenas)
    if (not concurso or len(set(dezenas)) != cfg["sorteadas"]
            or not all(1 <= d <= cfg["total"] for d in dezenas)):
        return None
    return {
        "concurso": int(concurso),
        "data": item.get("data") or item.get("dataApuracao") or "",
        "dezenas": dezenas,
    }


def atualizar_historico(cfg):
    """Baixa os sorteios que faltam. Devolve o historico atualizado."""
    sorteios = carregar_historico(cfg)
    antes = len(sorteios)
    try:
        dados = _baixar_json(URL_TODOS + cfg["chave"])
        novos = [n for n in (_normalizar(cfg, d) for d in dados) if n]
        if not novos:
            raise ValueError("resposta sem sorteios")
        sorteios = salvar_historico(cfg, sorteios + novos)
    except Exception as erro:  # cai para a API oficial, concurso a concurso
        print(f"Fonte principal indisponivel ({erro}); usando a API da Caixa...")
        try:
            base = f"{URL_CAIXA}{cfg['chave']}/"
            ultimo = _normalizar(cfg, _baixar_json(base))
            tenho = {s["concurso"] for s in sorteios}
            faltam = [c for c in range(1, ultimo["concurso"]) if c not in tenho]
            novos = [ultimo]
            for i, concurso in enumerate(faltam, 1):
                item = _normalizar(cfg, _baixar_json(f"{base}{concurso}"))
                if item:
                    novos.append(item)
                if i % 50 == 0:
                    print(f"  {i}/{len(faltam)} concursos baixados...")
                    salvar_historico(cfg, sorteios + novos)
            sorteios = salvar_historico(cfg, sorteios + novos)
        except Exception as erro2:
            print(f"Nao foi possivel baixar os resultados: {erro2}")
            print("Alternativa: baixe a planilha de resultados no site da Caixa,")
            print(f"salve como CSV e use:  python megasena.py -j {cfg['chave']} importar arquivo.csv")
            return sorteios
    print(f"Historico {cfg['nome']}: {len(sorteios)} sorteios ({len(sorteios) - antes} novos).")
    return sorteios


def importar_csv(cfg, caminho):
    """Importa um CSV com concurso, data e as dezenas sorteadas em cada linha."""
    texto = Path(caminho).read_text(encoding="utf-8-sig", errors="replace")
    delimitador = ";" if texto.count(";") > texto.count(",") else ","
    qtd = cfg["sorteadas"]
    novos = []
    for linha in csv.reader(texto.splitlines(), delimiter=delimitador):
        campos = [c.strip() for c in linha]
        if len(campos) < qtd + 1 or not campos[0].isdigit():
            continue
        data = campos[1] if re.fullmatch(r"\d{1,2}/\d{1,2}/\d{4}", campos[1]) else ""
        inicio = 2 if data else 1
        dezenas = campos[inicio:inicio + qtd]
        if (len(dezenas) == qtd and all(d.isdigit() and 1 <= int(d) <= cfg["total"] for d in dezenas)
                and len(set(map(int, dezenas))) == qtd):
            novos.append({"concurso": int(campos[0]), "data": data,
                          "dezenas": sorted(int(d) for d in dezenas)})
    if not novos:
        print(f"Nenhum sorteio da {cfg['nome']} reconhecido no arquivo.")
        return carregar_historico(cfg)
    sorteios = salvar_historico(cfg, carregar_historico(cfg) + novos)
    print(f"{len(novos)} sorteios lidos. Historico {cfg['nome']}: {len(sorteios)} sorteios.")
    return sorteios


# --------------------------------------------------------------------------
# Estatisticas
# --------------------------------------------------------------------------

def frequencias(cfg, sorteios):
    cont = Counter({n: 0 for n in numeros(cfg)})
    for s in sorteios:
        cont.update(s["dezenas"])
    return cont


def prob_teorica_pares(cfg, pares):
    """Probabilidade de um sorteio aleatorio ter exatamente `pares` pares."""
    n_pares = cfg["total"] // 2
    return (math.comb(n_pares, pares) * math.comb(cfg["total"] - n_pares, cfg["sorteadas"] - pares)
            / cfg["combinacoes"])


def fmt(dezenas):
    return " ".join(f"{d:02d}" for d in dezenas)


def mostrar_estatisticas(cfg, sorteios):
    if not sorteios:
        print(f"Sem historico. Rode primeiro:  python megasena.py -j {cfg['chave']} atualizar")
        return
    total, qtd = len(sorteios), cfg["sorteadas"]
    freq = frequencias(cfg, sorteios)
    recente = frequencias(cfg, sorteios[-JANELA_RECENTE:])
    top = 10 if cfg["total"] > 30 else 5
    print(f"\n{cfg['nome']}: {total} sorteios (concursos {sorteios[0]['concurso']} a {sorteios[-1]['concurso']})")
    print(f"Ultimo sorteio: {fmt(sorteios[-1]['dezenas'])}  {sorteios[-1]['data']}")

    esperado = total * qtd / cfg["total"]
    print(f"\nDezenas que MAIS sairam (esperado por dezena: {esperado:.0f} vezes)")
    for n, q in freq.most_common(top):
        print(f"  {n:02d}: {q} vezes")
    print("Dezenas que MENOS sairam")
    for n, q in sorted(freq.items(), key=lambda x: x[1])[:top]:
        print(f"  {n:02d}: {q} vezes")
    print(f"Mais sorteadas nos ultimos {JANELA_RECENTE} concursos: "
          + fmt(sorted(n for n, _ in recente.most_common(top))))

    print("\nPares x impares          sorteios   observado   teorico")
    divisao = Counter(sum(1 for d in s["dezenas"] if d % 2 == 0) for s in sorteios)
    n_pares = cfg["total"] // 2
    for pares in range(max(0, qtd - (cfg["total"] - n_pares)), min(qtd, n_pares) + 1):
        q = divisao[pares]
        print(f"  {pares:2d} pares / {qtd - pares:2d} impares   {q:6d}     {q / total:6.1%}    "
              f"{prob_teorica_pares(cfg, pares):6.1%}")
    pares_total = sum(p * q for p, q in divisao.items())
    print(f"  Dezenas sorteadas: {pares_total / (qtd * total):.1%} pares e "
          f"{1 - pares_total / (qtd * total):.1%} impares "
          f"(no volante: {n_pares} pares e {cfg['total'] - n_pares} impares)")


# --------------------------------------------------------------------------
# Estrategia de geracao
# --------------------------------------------------------------------------

def pesos(cfg, sorteios):
    """Peso de cada dezena: 60% frequencia historica + 40% frequencia recente.

    Sem historico, todas as dezenas recebem o mesmo peso.
    """
    if not sorteios:
        return {n: 1.0 for n in numeros(cfg)}
    geral = frequencias(cfg, sorteios)
    recente = frequencias(cfg, sorteios[-JANELA_RECENTE:])
    max_g = max(geral.values()) or 1
    max_r = max(recente.values()) or 1
    # o termo fixo 0.25 garante que nenhuma dezena fique impossivel de sair
    return {n: 0.25 + 0.6 * geral[n] / max_g + 0.4 * recente[n] / max_r for n in numeros(cfg)}


def sortear_ponderado(peso, quantidade, rng):
    """Amostra sem reposicao proporcional aos pesos (Efraimidis-Spirakis)."""
    chaves = sorted(peso, key=lambda n: rng.random() ** (1.0 / peso[n]), reverse=True)
    return sorted(chaves[:quantidade])


def maior_sequencia(dezenas):
    maior = atual = 1
    for a, b in zip(dezenas, dezenas[1:]):
        atual = atual + 1 if b == a + 1 else 1
        maior = max(maior, atual)
    return maior


def jogo_valido(cfg, dezenas, ja_sorteados=frozenset()):
    """Filtros de equilibrio: o jogo deve 'parecer' um sorteio tipico."""
    qtd, total = len(dezenas), cfg["total"]
    correcao = (total - qtd) / (total - 1)
    # pares x impares perto da proporcao do volante
    p = (total // 2) / total
    pares = sum(1 for d in dezenas if d % 2 == 0)
    if abs(pares - qtd * p) > 1.1 * math.sqrt(qtd * p * (1 - p) * correcao):
        return False
    # soma na faixa mais comum
    desvio = math.sqrt(qtd * (total * total - 1) / 12 * correcao)
    if abs(sum(dezenas) - qtd * (total + 1) / 2) > 1.5 * desvio:
        return False
    # sem longas sequencias de dezenas consecutivas
    limite, ate = cfg["seq"]
    if qtd <= ate and maior_sequencia(dezenas) > limite:
        return False
    # distribuicao pelas faixas/linhas do volante
    faixas = Counter((d - 1) // cfg["faixa"] for d in dezenas)
    if cfg["faixa_regra"] == "max_metade" and max(faixas.values()) > math.ceil(qtd / 2):
        return False
    if cfg["faixa_regra"] == "sem_vazia" and len(faixas) < total // cfg["faixa"]:
        return False
    # evita jogo "de aniversario" (so dezenas ate 31)
    if sum(1 for d in dezenas if d > 31) < cfg["min_altas"]:
        return False
    # nao repete um resultado que ja foi sorteado
    if tuple(dezenas) in ja_sorteados:
        return False
    return True


def gerar_jogos(cfg, sorteios, quantidade=1, dezenas=None, rng=None):
    """Gera jogos ponderados pela frequencia, equilibrados e pouco sobrepostos."""
    rng = rng or random.Random()
    dezenas = dezenas or cfg["sorteadas"]
    peso = pesos(cfg, sorteios)
    ja_sorteados = frozenset(tuple(s["dezenas"]) for s in sorteios)
    # jogos diferentes entre si: limite um pouco acima da sobreposicao media
    max_comum = math.ceil(dezenas * dezenas / cfg["total"]) + 1
    jogos = []
    while len(jogos) < quantidade:
        candidato = None
        for tentativa in range(5000):
            candidato = sortear_ponderado(peso, dezenas, rng)
            if not jogo_valido(cfg, candidato, ja_sorteados):
                continue
            # depois de muitas tentativas, relaxa a exigencia de diversidade
            if tentativa > 2500 or all(len(set(candidato) & set(j)) <= max_comum for j in jogos):
                break
        jogos.append(candidato)
    return jogos


def mostrar_jogos(cfg, sorteios, quantidade, dezenas):
    if not sorteios:
        print("Sem historico: os jogos serao equilibrados, mas sem usar frequencias.")
        print(f"Rode  python megasena.py -j {cfg['chave']} atualizar  para baixar os sorteios.")
    jogos = gerar_jogos(cfg, sorteios, quantidade, dezenas)
    print(f"\n{cfg['nome']}: {quantidade} jogo(s) de {dezenas} dezenas")
    for i, jogo in enumerate(jogos, 1):
        pares = sum(1 for d in jogo if d % 2 == 0)
        print(f"  Jogo {i:2d}: {fmt(jogo)}   ({pares} pares / {dezenas - pares} impares, soma {sum(jogo)})")
    apostas = quantidade * math.comb(dezenas, cfg["sorteadas"])
    print(f"\nChance de acertar as {cfg['sorteadas']} dezenas com esses jogos: "
          f"1 em {ponto(cfg['combinacoes'] / apostas)}")
    print("(a mesma de quaisquer outros jogos distintos na mesma quantidade)")


# --------------------------------------------------------------------------
# Backtest: a estrategia e melhor que o acaso?
# --------------------------------------------------------------------------

def backtest(cfg, sorteios, concursos=300, jogos_por_concurso=10, semente=1):
    """Para cada concurso, gera jogos usando SO os sorteios anteriores e
    compara os acertos com jogos totalmente aleatorios."""
    if len(sorteios) < 100:
        print("Historico insuficiente para o teste (minimo 100 sorteios).")
        return None
    concursos = min(concursos, len(sorteios) - 60)
    qtd = cfg["sorteadas"]
    rng = random.Random(semente)
    acertos = {"estrategia": Counter(), "aleatorio": Counter()}
    for i in range(len(sorteios) - concursos, len(sorteios)):
        resultado = set(sorteios[i]["dezenas"])
        for jogo in gerar_jogos(cfg, sorteios[:i], jogos_por_concurso, qtd, rng):
            acertos["estrategia"][len(resultado & set(jogo))] += 1
        for _ in range(jogos_por_concurso):
            acertos["aleatorio"][len(resultado & set(rng.sample(numeros(cfg), qtd)))] += 1

    total = concursos * jogos_por_concurso
    print(f"\nBacktest {cfg['nome']}: {concursos} concursos x {jogos_por_concurso} jogos = "
          f"{total} jogos por metodo")
    print("                 media de acertos   jogos com exatamente N acertos")
    print("                                   " + "".join(f"{k:>7d}" for k in cfg["acertos"]))
    for nome, cont in acertos.items():
        media = sum(k * v for k, v in cont.items()) / total
        print(f"  {nome:<12}      {media:6.3f}       " + "".join(f"{cont[k]:>7d}" for k in cfg["acertos"]))
    print(f"  esperado pelo acaso {qtd * qtd / cfg['total']:.3f}")
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


def menu(cfg):
    sorteios = carregar_historico(cfg)
    while True:
        print(f"\n=== {cfg['nome'].upper()} ===  ({len(sorteios)} sorteios no historico)")
        print("1) Atualizar historico de sorteios")
        print("2) Ver estatisticas (mais sorteadas, pares x impares)")
        print("3) Gerar jogos")
        print("4) Testar a estrategia contra o historico (backtest)")
        outro = JOGOS["lotofacil" if cfg["chave"] == "megasena" else "megasena"]
        print(f"5) Trocar para {outro['nome']}")
        print("0) Sair")
        opcao = input("Opcao: ").strip()
        if opcao == "1":
            sorteios = atualizar_historico(cfg)
        elif opcao == "2":
            mostrar_estatisticas(cfg, sorteios)
        elif opcao == "3":
            minimo, maximo = cfg["sorteadas"], cfg["aposta_max"]
            mostrar_jogos(cfg, sorteios, ler_int("Quantos jogos", 3, 1, 100),
                          ler_int(f"Dezenas por jogo ({minimo} a {maximo})", minimo, minimo, maximo))
        elif opcao == "4":
            backtest(cfg, sorteios)
        elif opcao == "5":
            cfg = outro
            sorteios = carregar_historico(cfg)
        elif opcao == "0":
            break


def main():
    parser = argparse.ArgumentParser(description="Gerador de jogos da Mega-Sena e da Lotofacil")
    parser.add_argument("-j", "--jogo", choices=list(JOGOS), default="megasena",
                        help="qual loteria usar (padrao: megasena)")
    sub = parser.add_subparsers(dest="comando")
    sub.add_parser("atualizar", help="baixa/atualiza o historico de sorteios")
    sub.add_parser("estatisticas", help="mostra frequencias e pares x impares")
    p = sub.add_parser("gerar", help="gera jogos")
    p.add_argument("-n", "--jogos", type=int, default=3)
    p.add_argument("-d", "--dezenas", type=int, help="dezenas por jogo (Mega 6-15, Lotofacil 15-20)")
    p = sub.add_parser("backtest", help="compara a estrategia com jogos aleatorios")
    p.add_argument("-c", "--concursos", type=int, default=300)
    p.add_argument("-n", "--jogos", type=int, default=10)
    p = sub.add_parser("importar", help="importa sorteios de um arquivo CSV")
    p.add_argument("arquivo")
    args = parser.parse_args()
    cfg = JOGOS[args.jogo]

    if args.comando is None:
        menu(cfg)
    elif args.comando == "atualizar":
        atualizar_historico(cfg)
    elif args.comando == "estatisticas":
        mostrar_estatisticas(cfg, carregar_historico(cfg))
    elif args.comando == "gerar":
        dezenas = args.dezenas or cfg["sorteadas"]
        if not cfg["sorteadas"] <= dezenas <= cfg["aposta_max"]:
            parser.error(f"{cfg['nome']}: dezenas por jogo de {cfg['sorteadas']} a {cfg['aposta_max']}")
        mostrar_jogos(cfg, carregar_historico(cfg), max(1, args.jogos), dezenas)
    elif args.comando == "backtest":
        backtest(cfg, carregar_historico(cfg), args.concursos, args.jogos)
    elif args.comando == "importar":
        importar_csv(cfg, args.arquivo)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(0)
