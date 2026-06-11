import pytest
from ofx_reader import parse_ofx_conteudo

OFX_EXEMPLO = """OFXHEADER:100
DATA:OFXSGML
VERSION:102

<OFX>
<BANKTRANLIST>
<DTSTART>20260610100000[-03:EST]
<DTEND>20260610100000[-03:EST]
<STMTTRN>
<TRNTYPE>DEBIT
<DTPOSTED>20260610100000[-03:EST]
<TRNAMT>-113213.64
<FITID>20260610074
<MEMO>SISPAG FORNECEDORES
</STMTTRN>
<STMTTRN>
<TRNTYPE>DEBIT
<DTPOSTED>20260610100000[-03:EST]
<TRNAMT>-18839.11
<FITID>20260610001
<MEMO>TED ENVIADA CARLOS ALBERTO LTDA
</STMTTRN>
<STMTTRN>
<TRNTYPE>DEBIT
<DTPOSTED>20260610100000[-03:EST]
<TRNAMT>-64156.96
<FITID>20260610009
<MEMO>PAGAMENTOS TRIB COD BARRAS DARF 00.394.460/0058-87
</STMTTRN>
<STMTTRN>
<TRNTYPE>CREDIT
<DTPOSTED>20260610100000[-03:EST]
<TRNAMT>12953.52
<FITID>20260610006
<MEMO>PIX RECEBIDO FUNDAME10/06 FUNDAMENTOS SISTEMAS LTDA
</STMTTRN>
</BANKTRANLIST>
</OFX>
"""


def test_classifica_agrupado_sispag():
    r = parse_ofx_conteudo(OFX_EXEMPLO)
    assert len(r["agrupados"]) == 1
    assert r["agrupados"][0]["valor"] == pytest.approx(113213.64)
    assert r["agrupados"][0]["data"] == "10/06/2026"


def test_classifica_detalhado():
    r = parse_ofx_conteudo(OFX_EXEMPLO)
    assert len(r["detalhados"]) == 1
    assert r["detalhados"][0]["memo"] == "TED ENVIADA CARLOS ALBERTO LTDA"
    assert r["detalhados"][0]["valor"] == pytest.approx(18839.11)


def test_tributo_fora_do_matching():
    r = parse_ofx_conteudo(OFX_EXEMPLO)
    assert len(r["tributos"]) == 1
    memos_detalhados = [t["memo"] for t in r["detalhados"]]
    assert not any("TRIB" in m for m in memos_detalhados)


def test_credito_ignorado():
    r = parse_ofx_conteudo(OFX_EXEMPLO)
    assert r["n_creditos"] == 1
    todos = r["agrupados"] + r["detalhados"] + r["tributos"]
    assert all(t["valor"] > 0 for t in todos)


def test_dias_do_extrato():
    r = parse_ofx_conteudo(OFX_EXEMPLO)
    assert r["dias"] == {"10/06/2026"}
