import re
import unicodedata
from collections import defaultdict

TOLERANCIA = 0.01


def _normalizar_nome(texto: str) -> str:
    s = unicodedata.normalize("NFKD", str(texto))
    s = s.encode("ascii", "ignore").decode("ascii").upper()
    return re.sub(r"[^A-Z0-9 ]", " ", s).strip()


def _score_nome(favorecido: str, memo: str) -> int:
    palavras_memo = set(_normalizar_nome(memo).split())
    palavras_fav = [p for p in _normalizar_nome(favorecido).split() if len(p) >= 3]
    return sum(1 for p in palavras_fav if p in palavras_memo)


def casar_lancamentos(df_xls, detalhados: list) -> tuple:
    banco_por_chave = defaultdict(list)
    for t in detalhados:
        banco_por_chave[(t["data"], round(t["valor"], 2))].append(t)

    xls_por_chave = defaultdict(list)
    for idx, row in df_xls.iterrows():
        xls_por_chave[(row["data"], round(float(row["valor_num"]), 2))].append(idx)

    casados, nao_casados = [], []
    for chave, indices in xls_por_chave.items():
        vagas = len(banco_por_chave.get(chave, []))
        if vagas == 0:
            nao_casados.extend(indices)
            continue
        if vagas >= len(indices):
            casados.extend(indices)
            continue
        memos = [t["memo"] for t in banco_por_chave[chave]]

        def melhor_score(idx):
            fav = df_xls.loc[idx, "favorecido"]
            return max(_score_nome(fav, m) for m in memos)

        ordenados = sorted(indices, key=melhor_score, reverse=True)
        casados.extend(ordenados[:vagas])
        nao_casados.extend(ordenados[vagas:])

    return casados, nao_casados


def validar_fechamento(df_xls, nao_casados: list, agrupados: list,
                       dias_extrato: set) -> dict:
    soma_xls_dia = defaultdict(float)
    for idx in nao_casados:
        row = df_xls.loc[idx]
        soma_xls_dia[row["data"]] += float(row["valor_num"])

    soma_agr_dia = defaultdict(float)
    for t in agrupados:
        soma_agr_dia[t["data"]] += t["valor"]

    resultado = {}
    for dia in sorted(set(soma_xls_dia) | set(soma_agr_dia)):
        s_xls = round(soma_xls_dia.get(dia, 0.0), 2)
        s_agr = round(soma_agr_dia.get(dia, 0.0), 2)
        dif = round(s_xls - s_agr, 2)
        if dia not in dias_extrato:
            status = "fora_do_extrato"
        elif abs(dif) <= TOLERANCIA:
            status = "ok"
        else:
            status = "divergente"
        resultado[dia] = {
            "soma_nao_casados": s_xls,
            "soma_agrupados": s_agr,
            "diferenca": dif,
            "status": status,
        }
    return resultado
