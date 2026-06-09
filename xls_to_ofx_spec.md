# xls_to_ofx — Spec do Projeto

**Goal:** Converter o relatório XLS de pagamentos SISPAG do Itaú em arquivo OFX válido, para importação direta no Omie e eliminação do retrabalho de conciliação bancária causado pelo agrupamento nativo do OFX do Itaú.

**Problema resolvido:** O OFX gerado pelo Itaú agrupa todos os lançamentos SISPAG em entradas únicas de alto valor (ex: R$ 393.405,74), impossibilitando a conciliação automática no Omie. O relatório XLS de "Consulta de Pagamentos" do Itaú traz os mesmos lançamentos individualizados por favorecido.

---

## Contexto de Negócio

- **Empresa:** Telbra-Ex (Grupo Telbra) — conta Itaú PJ, agência 0263, conta 43929-1
- **Sistema ERP:** Omie
- **Fluxo atual (problemático):**
  1. Itaú gera OFX com SISPAG agrupado
  2. Usuário importa no Omie → lançamentos de R$ 100–400k aparecem como um único item
  3. Usuário precisa baixar PDF/XLS separado para consultar os lançamentos individuais
  4. Conciliação manual, lenta e propensa a erro
- **Fluxo desejado:**
  1. Itaú → exportar XLS de "Consulta de Pagamentos" (SISPAG)
  2. Rodar `xls_to_ofx.py` → gera `.ofx`
  3. Importar OFX no Omie → cada lançamento aparece separado, conciliação direta

---

## Estrutura de Arquivos

```
xls_to_ofx/
├── xls_to_ofx.py          # Script principal (CLI)
├── requirements.txt        # pandas, xlrd
├── tests/
│   ├── test_leitura.py     # Testa leitura e parsing do XLS
│   ├── test_ofx.py         # Testa geração do OFX
│   └── fixtures/
│       └── sample.xls      # Arquivo XLS de amostra (dados anonimizados)
└── README.md
```

---

## Formato de Entrada — XLS Itaú SISPAG

O arquivo é exportado no portal do Itaú em "Consulta de Pagamentos". Estrutura:

| Linhas | Conteúdo |
|--------|----------|
| 0–4 | Cabeçalho vazio |
| 5 | `Agência/conta` \| `0263 / 43929-1` \| ... \| `Nome da empresa:` \| `TELBRA EX I COMERCIO LTDA` |
| 6 | `CNPJ:` \| `13.705.784/0001-03` |
| 7–17 | Filtros da consulta (período, modalidade, status) |
| **18** | **Linha de cabeçalho das colunas** |
| 19+ | Dados de lançamentos |

### Colunas (a partir da linha 18)

| Coluna original | Campo interno | Tipo |
|----------------|---------------|------|
| `favorecido / beneficiário` | `favorecido` | string |
| `CPF/CNPJ` | `cpf_cnpj` | string |
| `tipo de pagamento` | `tipo` | string |
| `referência da empresa` | `referencia` | string |
| `data do pagamento` | `data` | string `DD/MM/AAAA` |
| `valor (R$)` | `valor` | float |
| `status` | `status` | string: `Efetuado` / `Não efetuado` |

### Tipos de pagamento conhecidos
- `Boleto outros bancos`
- `Boleto Itaú`
- `TED mesmo titular`
- `TED outro titular`
- `Conta Corrente`

---

## Formato de Saída — OFX

Padrão OFX SGML versão 1.02, codificação `cp1252` (coerente com `CHARSET:1252` do header), compatível com Omie.

### Campos mapeados por transação

| Campo OFX | Origem | Observação |
|-----------|--------|------------|
| `TRNTYPE` | fixo | sempre `DEBIT` (saídas) |
| `DTPOSTED` | `data` | formato `AAAAMMDD100000[-03:EST]` |
| `TRNAMT` | `valor` | negativo (débito): `-3756.20` |
| `FITID` | gerado | MD5(data\|favorecido\|referência\|tipo\|valor\|ocorrência)[:16] — estável e único; independe da posição na planilha |
| `MEMO` | `favorecido` + `referencia` + `tipo` | ex: `OMIEXPERIENCE LTDA \| Ref: FATURA \| (Boleto outros bancos)` |

### Metadados da conta no OFX

| Campo OFX | Origem |
|-----------|--------|
| `BANKID` | fixo `0341` (código ITAÚ) |
| `BRANCHID` | extraído da linha 5 do XLS (antes do `/`) |
| `ACCTID` | extraído da linha 5 do XLS (depois do `/`) |
| `ACCTTYPE` | fixo `CHECKING` |

---

## Comportamento Esperado

### Filtragem
- **Padrão:** incluir apenas lançamentos com `status == "Efetuado"`
- **Flag `--incluir-nao-efetuado`:** inclui todos os status

### Deduplicação
- FITID gerado por hash garante que reimportar o mesmo arquivo não cria duplicatas no Omie (o sistema descarta FITIDs já vistos)

### Datas do extrato
- `DTSTART` = data do primeiro lançamento
- `DTEND` = data do último lançamento

---

## CLI — Interface de Linha de Comando

```bash
# Uso básico (gera arquivo com mesmo nome, extensão .ofx)
python xls_to_ofx.py ConsultaPagamentos.xls

# Especificar arquivo de saída
python xls_to_ofx.py ConsultaPagamentos.xls saida.ofx

# Incluir lançamentos não efetuados
python xls_to_ofx.py ConsultaPagamentos.xls --incluir-nao-efetuado
```

**Output esperado no terminal:**
```
Lendo: ConsultaPagamentos.xls
  Lançamentos encontrados: 55
  Agência: 0263 | Conta: 43929-1
OFX gerado: ConsultaPagamentos.ofx
  Total em débitos: R$ 544.301,10
```

---

## Ajustes e Melhorias Pendentes

Estes são os pontos identificados para evolução do script atual:

### P1 — Crítico

- [x] **Suporte a múltiplas datas no mesmo arquivo**
  `DTSTART`/`DTEND` usam `min()`/`max()` das datas — correto mesmo com o XLS desordenado ou com vários dias.

- [x] **Validação de colunas obrigatórias**
  `ler_xls` verifica todas as colunas esperadas e levanta `ValueError` com mensagem clara (lista o que faltou e o que está presente) antes de processar.

- [x] **Tratamento de valores nulos/inválidos**
  Linhas com `valor` vazio/não numérico são ignoradas com aviso por linha. `to_float` também trata formato BR (`3.756,20`, `R$`) e NaN sem crash.

- [x] **Datas exportadas como tipo data do Excel**
  `normalizar_data` converte `datetime`/`Timestamp` para `DD/MM/AAAA`. Sem isso, todas as linhas seriam descartadas e o OFX sairia vazio.

### P2 — Importante

- [ ] **Suporte a XLS com múltiplas abas**
  Verificar se o Itaú pode exportar arquivos com mais de uma aba e garantir que o script lê a aba correta.

- [x] **Log de lançamentos ignorados**
  A CLI exibe contagem de sem-data, não efetuados e inválidos, com o motivo de cada descarte.

- [x] **Compatibilidade com XLSX além de XLS**
  `detectar_engine` escolhe `xlrd` (.xls) ou `openpyxl` (.xlsx) automaticamente pela extensão.

- [x] **FITID estável para reimportação**
  O FITID independe da posição da linha (usa ocorrência entre dados idênticos), evitando duplicatas no Omie ao reexportar períodos maiores.

### P3 — Desejável

- [x] **Interface gráfica simples**
  Implementada em CustomTkinter (`gui.py`): seleção de arquivo, opção de incluir não efetuados, conversão em thread e status. Empacotável em `.exe` via `build.bat`.

- [ ] **Validação do OFX gerado**
  Checar se o arquivo gerado é bem-formado antes de salvar (tags abertas/fechadas, encoding correto).

- [ ] **Teste de importação no Omie**
  Validar que o OFX gerado é aceito pelo Omie sem erros e que os lançamentos aparecem corretamente na tela de conciliação bancária.

- [ ] **Revisar `CHECKNUM`**
  Hoje `CHECKNUM` recebe o mesmo hash do `FITID`, o que é semanticamente incorreto (deveria ser nº de documento). Avaliar omiti-lo — requer reteste no Omie antes de alterar.

---

## Dependências

```
pandas>=1.5.0
xlrd>=2.0.1
```

Instalação:
```bash
pip install pandas xlrd
```

Python mínimo: 3.8

---

## Limitações Conhecidas

1. **Cobre apenas saídas (SISPAG).** Entradas (PIX recebido, TED entrada, boletos recebidos) continuam precisando do OFX nativo do Itaú. O fluxo completo de conciliação usa os dois arquivos.

2. **Layout fixo do XLS.** O cabeçalho está na linha 18 (base 0). Se o Itaú alterar o layout do relatório, o script para de funcionar e precisa de ajuste manual no `HEADER_ROW`.

3. **MEMO truncado.** O campo `favorecido` no XLS é truncado em ~32 caracteres pelo Itaú. Isso pode dificultar a identificação de favorecidos com nomes similares.
