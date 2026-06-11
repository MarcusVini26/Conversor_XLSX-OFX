import re
from datetime import datetime

MEMO_AGRUPADO_PREFIXOS = ("SISPAG",)
MEMO_TRIBUTO_PREFIXOS = ("PAGAMENTOS TRIB",)


def _dtposted_para_ddmmaaaa(valor: str) -> str:
    m = re.match(r"\s*(\d{8})", str(valor))
    if not m:
        return ""
    try:
        return datetime.strptime(m.group(1), "%Y%m%d").strftime("%d/%m/%Y")
    except ValueError:
        return ""


def parse_ofx_conteudo(conteudo: str) -> dict:
    agrupados, detalhados, tributos = [], [], []
    n_creditos = 0
    dias = set()

    blocos = re.findall(r"<STMTTRN>(.*?)</STMTTRN>", conteudo, re.S | re.I)
    for bloco in blocos:
        campos = dict(re.findall(r"<([A-Za-z0-9]+)>([^\r\n<]*)", bloco))
        campos = {k.upper(): v.strip() for k, v in campos.items()}

        try:
            valor = float(campos.get("TRNAMT", ""))
        except ValueError:
            continue

        data = _dtposted_para_ddmmaaaa(campos.get("DTPOSTED", ""))
        if data:
            dias.add(data)

        trntype = campos.get("TRNTYPE", "").upper()
        if trntype == "CREDIT" or valor > 0:
            n_creditos += 1
            continue

        txn = {
            "data": data,
            "valor": abs(valor),
            "memo": campos.get("MEMO", ""),
            "fitid": campos.get("FITID", ""),
        }
        memo_upper = txn["memo"].upper()
        if memo_upper.startswith(MEMO_AGRUPADO_PREFIXOS):
            agrupados.append(txn)
        elif memo_upper.startswith(MEMO_TRIBUTO_PREFIXOS):
            tributos.append(txn)
        else:
            detalhados.append(txn)

    return {
        "agrupados": agrupados,
        "detalhados": detalhados,
        "tributos": tributos,
        "n_creditos": n_creditos,
        "dias": dias,
    }


def ler_ofx_banco(caminho: str) -> dict:
    with open(caminho, "r", encoding="cp1252", errors="replace") as f:
        return parse_ofx_conteudo(f.read())
