import pandas as pd
from xls_to_ofx import filtrar_tributos


def test_remove_linhas_de_tributo():
    df = pd.DataFrame([
        {"tipo": "Boleto outros bancos", "favorecido": "A"},
        {"tipo": "DARF", "favorecido": "B"},
        {"tipo": "Tributo GNRE", "favorecido": "C"},
    ])
    df_limpo, n = filtrar_tributos(df)
    assert n == 2
    assert list(df_limpo["favorecido"]) == ["A"]


def test_sem_tributos_nada_muda():
    df = pd.DataFrame([{"tipo": "TED outro titular", "favorecido": "A"}])
    df_limpo, n = filtrar_tributos(df)
    assert n == 0 and len(df_limpo) == 1
