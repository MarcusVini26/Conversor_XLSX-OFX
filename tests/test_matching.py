import pandas as pd
import pytest
from matching import casar_lancamentos, validar_fechamento


def _df(rows):
    return pd.DataFrame(
        [{"data": d, "valor_num": v, "favorecido": f} for d, v, f in rows]
    )


def _det(data, valor, memo):
    return {"data": data, "valor": valor, "memo": memo, "fitid": "X"}


def test_casa_por_valor_e_data():
    df = _df([("10/06/2026", 18839.11, "CARLOS ALBERTO LTDA")])
    det = [_det("10/06/2026", 18839.11, "TED ENVIADA CARLOS ALBERTO LTDA")]
    casados, nao_casados = casar_lancamentos(df, det)
    assert len(casados) == 1 and len(nao_casados) == 0


def test_nao_casa_data_diferente():
    df = _df([("11/06/2026", 18839.11, "CARLOS ALBERTO LTDA")])
    det = [_det("10/06/2026", 18839.11, "TED ENVIADA CARLOS ALBERTO LTDA")]
    casados, nao_casados = casar_lancamentos(df, det)
    assert len(casados) == 0 and len(nao_casados) == 1


def test_contagem_dois_iguais_no_xls_um_no_banco():
    df = _df([
        ("10/06/2026", 500.00, "FORNECEDOR A"),
        ("10/06/2026", 500.00, "FORNECEDOR B"),
    ])
    det = [_det("10/06/2026", 500.00, "TED ENVIADA FORNECEDOR A LTDA")]
    casados, nao_casados = casar_lancamentos(df, det)
    assert len(casados) == 1 and len(nao_casados) == 1


def test_desempate_por_nome():
    df = _df([
        ("10/06/2026", 500.00, "PADARIA DOIS IRMAOS"),
        ("10/06/2026", 500.00, "WESLEY TRANSPORTES"),
    ])
    det = [_det("10/06/2026", 500.00, "TED ENVIADA WESLEY TRANSPORTES ME")]
    casados, nao_casados = casar_lancamentos(df, det)
    assert df.loc[casados[0], "favorecido"] == "WESLEY TRANSPORTES"
    assert df.loc[nao_casados[0], "favorecido"] == "PADARIA DOIS IRMAOS"


def test_centavos_diferentes_nao_casam():
    df = _df([("10/06/2026", 500.01, "FORNECEDOR A")])
    det = [_det("10/06/2026", 500.00, "TED ENVIADA FORNECEDOR A")]
    casados, nao_casados = casar_lancamentos(df, det)
    assert len(casados) == 0 and len(nao_casados) == 1


def test_fechamento_ok():
    df = _df([
        ("10/06/2026", 100000.00, "A"),
        ("10/06/2026", 13213.64, "B"),
    ])
    agrupados = [_det("10/06/2026", 113213.64, "SISPAG FORNECEDORES")]
    resultado = validar_fechamento(df, nao_casados=[0, 1], agrupados=agrupados,
                                   dias_extrato={"10/06/2026"})
    assert resultado["10/06/2026"]["status"] == "ok"
    assert resultado["10/06/2026"]["diferenca"] == pytest.approx(0.0, abs=0.01)


def test_fechamento_divergente():
    df = _df([("10/06/2026", 100000.00, "A")])
    agrupados = [_det("10/06/2026", 113213.64, "SISPAG FORNECEDORES")]
    resultado = validar_fechamento(df, nao_casados=[0], agrupados=agrupados,
                                   dias_extrato={"10/06/2026"})
    assert resultado["10/06/2026"]["status"] == "divergente"
    assert resultado["10/06/2026"]["diferenca"] == pytest.approx(-13213.64, abs=0.01)


def test_dois_agrupados_no_mesmo_dia_somam():
    df = _df([("10/06/2026", 300.00, "A"), ("10/06/2026", 700.00, "B")])
    agrupados = [
        _det("10/06/2026", 300.00, "SISPAG FORNECEDORES"),
        _det("10/06/2026", 700.00, "SISPAG FORNECEDORES"),
    ]
    resultado = validar_fechamento(df, nao_casados=[0, 1], agrupados=agrupados,
                                   dias_extrato={"10/06/2026"})
    assert resultado["10/06/2026"]["status"] == "ok"


def test_dia_fora_do_extrato():
    df = _df([("12/06/2026", 50.00, "A")])
    resultado = validar_fechamento(df, nao_casados=[0], agrupados=[],
                                   dias_extrato={"10/06/2026"})
    assert resultado["12/06/2026"]["status"] == "fora_do_extrato"
