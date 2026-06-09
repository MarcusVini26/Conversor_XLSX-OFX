"""
Conversor: Extrato SISPAG Itaú (XLS/XLSX) → OFX
Uso: python xls_to_ofx.py <arquivo.xls|xlsx> [saida.ofx] [--incluir-nao-efetuado]

- Lê o relatório de pagamentos exportado pelo Itaú (SISPAG / Consulta Pagamentos)
- Gera um arquivo .ofx compatível com importação no Omie
- Apenas lançamentos com status "Efetuado" são incluídos (padrão)
  Use --incluir-nao-efetuado para incluir todos

Dependências: pip install pandas xlrd openpyxl
"""

import pandas as pd
import sys
import os
import hashlib
from datetime import datetime

# ─── Configurações ───────────────────────────────────────────────────────────

HEADER_ROW = 18
TRNTYPE = "DEBIT"

COLUNAS_ESPERADAS = {
    "favorecido / beneficiário": "favorecido",
    "cpf/cnpj": "cpf_cnpj",
    "tipo de pagamento": "tipo",
    "referência da empresa": "referencia",
    "data do pagamento": "data",
    "valor (r$)": "valor",
    "status": "status",
}

# ─── Utilitários ─────────────────────────────────────────────────────────────

def detectar_engine(caminho: str) -> str:
    """Escolhe o engine pandas correto com base na extensão do arquivo."""
    ext = os.path.splitext(caminho)[1].lower()
    if ext == ".xlsx":
        return "openpyxl"
    return "xlrd"


def to_float(valor):
    """Converte valor para float, retorna None se inválido."""
    try:
        return float(str(valor).replace(",", "."))
    except (ValueError, TypeError):
        return None


def parse_data(data_str: str):
    try:
        return datetime.strptime(str(data_str).strip(), "%d/%m/%Y")
    except ValueError:
        return None


def formatar_data_ofx(data_str: str) -> str:
    """Converte 'DD/MM/AAAA' → 'AAAAMMDD100000[-03:EST]' (formato OFX Itaú)."""
    dt = parse_data(data_str)
    base = dt.strftime("%Y%m%d") if dt else datetime.today().strftime("%Y%m%d")
    return base + "100000[-03:EST]"


def formatar_acctid(agencia: str, conta: str) -> str:
    """Formata ACCTID no padrão Itaú: agência (4 dígitos) + conta sem traço."""
    return agencia.strip().zfill(4) + conta.strip().replace("-", "").replace(" ", "")


# ─── Leitura do arquivo ───────────────────────────────────────────────────────

def ler_xls(caminho: str, apenas_efetuados: bool = True) -> tuple:
    """
    Lê e filtra o XLS/XLSX do Itaú.

    Retorna (df, stats) onde stats tem:
      total_lido, nao_efetuados, invalidos, incluidos
    """
    engine = detectar_engine(caminho)
    df = pd.read_excel(caminho, engine=engine, header=HEADER_ROW)

    df.columns = [str(c).strip().lower() for c in df.columns]
    df = df.dropna(how="all")

    # Valida colunas obrigatórias antes de tentar renomear
    colunas_faltando = [c for c in COLUNAS_ESPERADAS if c not in df.columns]
    if colunas_faltando:
        raise ValueError(
            "Colunas não encontradas no arquivo XLS:\n"
            + "\n".join(f"  - '{c}'" for c in colunas_faltando)
            + f"\n\nColunas presentes: {list(df.columns)}\n"
            "O layout do relatório pode ter mudado. Ajuste COLUNAS_ESPERADAS no script."
        )

    df = df.rename(columns=COLUNAS_ESPERADAS)

    # Remove linhas sem data válida (DD/MM/AAAA) — captura linha de Total e afins
    mask_data_invalida = ~df["data"].astype(str).str.strip().str.match(r"^\d{2}/\d{2}/\d{4}$")
    n_sem_data = int(mask_data_invalida.sum())
    df = df[~mask_data_invalida].copy()

    total_lido = len(df)

    # Comparação com fillna("") + match POSITIVO — robusto a NaN e variações de encoding
    # Inclui APENAS linhas onde status é exatamente "efetuado" (case-insensitive, sem espaços)
    mask_efetuado = df["status"].fillna("").str.strip().str.lower() == "efetuado"
    n_nao_efetuados = int((~mask_efetuado).sum())
    if apenas_efetuados:
        df = df[mask_efetuado].copy()

    # Remove linhas com valor nulo ou não numérico, emitindo aviso por linha
    df["_valor_float"] = df["valor"].apply(to_float)
    mask_invalido = df["_valor_float"].isna()
    n_invalidos = int(mask_invalido.sum())
    if n_invalidos:
        for _, row in df[mask_invalido].iterrows():
            print(
                f"  [AVISO] Linha ignorada — valor inválido: "
                f"favorecido='{row.get('favorecido', '')}' | valor='{row.get('valor', '')}'"
            )
    df = df[~mask_invalido].drop(columns=["_valor_float"])

    stats = {
        "total_lido": total_lido,
        "nao_efetuados": n_nao_efetuados,
        "invalidos": n_invalidos,
        "sem_data": n_sem_data,
        "incluidos": len(df),
    }

    return df.reset_index(drop=True), stats


# ─── Geração do OFX ──────────────────────────────────────────────────────────

def gerar_fitid(row: pd.Series, idx: int) -> str:
    """Gera FITID único por transação via hash — garante deduplicação no Omie."""
    chave = f"{row['data']}|{row['favorecido']}|{row['valor']}|{idx}"
    return hashlib.md5(chave.encode()).hexdigest()[:16].upper()


def montar_memo(row: pd.Series) -> str:
    """Compõe descrição legível para o campo MEMO do OFX."""
    partes = [str(row.get("favorecido", "")).strip()]
    ref = str(row.get("referencia", "")).strip()
    tipo = str(row.get("tipo", "")).strip()
    if ref and ref != "-" and ref.lower() != "nan":
        partes.append(f"Ref: {ref}")
    if tipo and tipo.lower() != "nan":
        partes.append(f"({tipo})")
    return " | ".join(partes)


def gerar_ofx(df: pd.DataFrame, agencia: str = "", conta: str = "") -> str:
    agora = datetime.now().strftime("%Y%m%d") + "100000[-03:EST]"
    acctid = formatar_acctid(agencia, conta)

    # DTSTART/DTEND via min/max — correto mesmo se o XLS não estiver ordenado
    datas_parsed = df["data"].dropna().apply(parse_data).dropna()
    if not datas_parsed.empty:
        primeira = datas_parsed.min().strftime("%Y%m%d") + "100000[-03:EST]"
        ultima = datas_parsed.max().strftime("%Y%m%d") + "100000[-03:EST]"
    else:
        primeira = ultima = agora

    # Tags sem indentação — padrão do OFX real do Itaú
    linhas = [
        "OFXHEADER:100",
        "DATA:OFXSGML",
        "VERSION:102",
        "SECURITY:NONE",
        "ENCODING:USASCII",
        "CHARSET:1252",
        "COMPRESSION:NONE",
        "OLDFILEUID:NONE",
        "NEWFILEUID:NONE",
        "",
        "<OFX>",
        "<SIGNONMSGSRSV1>",
        "<SONRS>",
        "<STATUS>",
        "<CODE>0",
        "<SEVERITY>INFO",
        "</STATUS>",
        f"<DTSERVER>{agora}",
        "<LANGUAGE>POR",
        "</SONRS>",
        "</SIGNONMSGSRSV1>",
        "<BANKMSGSRSV1>",
        "<STMTTRNRS>",
        "<TRNUID>1001",
        "<STATUS>",
        "<CODE>0",
        "<SEVERITY>INFO",
        "</STATUS>",
        "<STMTRS>",
        "<CURDEF>BRL",
        "<BANKACCTFROM>",
        "<BANKID>0341",
        f"<ACCTID>{acctid}",
        "<ACCTTYPE>CHECKING",
        "</BANKACCTFROM>",
        "<BANKTRANLIST>",
        f"<DTSTART>{primeira}",
        f"<DTEND>{ultima}",
    ]

    for idx, row in df.iterrows():
        valor_float = to_float(row.get("valor", 0)) or 0.0
        valor_ofx = -abs(valor_float)
        fitid = gerar_fitid(row, idx)

        linhas += [
            "<STMTTRN>",
            f"<TRNTYPE>{TRNTYPE}",
            f"<DTPOSTED>{formatar_data_ofx(row.get('data', ''))}",
            f"<TRNAMT>{valor_ofx:.2f}",
            f"<FITID>{fitid}",
            f"<CHECKNUM>{fitid}",
            f"<MEMO>{montar_memo(row)}",
            "</STMTTRN>",
        ]

    linhas += [
        "</BANKTRANLIST>",
        "</STMTRS>",
        "</STMTTRNRS>",
        "</BANKMSGSRSV1>",
        "</OFX>",
    ]

    return "\n".join(linhas)


# ─── Extração de metadados do cabeçalho do XLS ───────────────────────────────

def extrair_metadados(caminho: str) -> dict:
    """Lê agência/conta da seção de cabeçalho do XLS (linha 5, base 0)."""
    engine = detectar_engine(caminho)
    df_raw = pd.read_excel(caminho, engine=engine, header=None, nrows=10)
    agencia, conta = "", ""
    for _, row in df_raw.iterrows():
        col0 = str(row.iloc[0]).strip().lower()
        if "agência" in col0 or "agencia" in col0:
            valor = str(row.iloc[1]).strip()
            if "/" in valor:
                partes = valor.split("/")
                agencia = partes[0].strip()
                conta = partes[1].strip() if len(partes) > 1 else ""
            break
    return {"agencia": agencia, "conta": conta}


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("Uso: python xls_to_ofx.py <arquivo.xls|xlsx> [saida.ofx] [--incluir-nao-efetuado]")
        sys.exit(1)

    apenas_efetuados = "--incluir-nao-efetuado" not in sys.argv
    args = [a for a in sys.argv[1:] if a != "--incluir-nao-efetuado"]

    caminho_entrada = args[0]

    if len(args) >= 2:
        caminho_saida = args[1]
    else:
        base = os.path.splitext(caminho_entrada)[0]
        caminho_saida = base + ".ofx"

    print(f"Lendo: {caminho_entrada}")

    try:
        meta = extrair_metadados(caminho_entrada)
        df, stats = ler_xls(caminho_entrada, apenas_efetuados=apenas_efetuados)
    except ValueError as e:
        print(f"\n[ERRO] {e}")
        sys.exit(1)

    print(f"  Lançamentos lidos: {stats['total_lido']}")
    if stats["sem_data"]:
        print(f"  Linhas sem data (totais/cabeçalhos): {stats['sem_data']} ignoradas")
    if stats["nao_efetuados"]:
        acao = "ignorados" if apenas_efetuados else "incluídos"
        print(f"  Não efetuados: {stats['nao_efetuados']} ({acao})")
    if stats["invalidos"]:
        print(f"  Inválidos (valor nulo/não numérico): {stats['invalidos']} ignorados")
    print(f"  Incluídos no OFX: {stats['incluidos']}")
    print(f"  Agência: {meta['agencia']} | Conta: {meta['conta']}")

    ofx = gerar_ofx(df, agencia=meta["agencia"], conta=meta["conta"])

    with open(caminho_saida, "w", encoding="latin-1", errors="replace") as f:
        f.write(ofx)

    total = df["valor"].apply(lambda v: to_float(v) or 0.0).sum()
    print(f"OFX gerado: {caminho_saida}")
    print(f"  Total em débitos: R$ {total:,.2f}")


if __name__ == "__main__":
    main()
