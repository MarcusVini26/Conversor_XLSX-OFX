# GUI Conversor SISPAG → OFX — Design

**Data:** 2026-06-09  
**Status:** Aprovado

---

## Objetivo

Interface gráfica para que o responsável pela conciliação bancária converta o relatório XLS/XLSX do SISPAG Itaú em arquivo OFX sem precisar usar o terminal.

O executável final deve rodar em PCs **sem Python instalado** (distribuído como `.exe` único via PyInstaller).

---

## Arquitetura

- **`gui.py`** — arquivo único com a interface, na mesma pasta de `xls_to_ofx.py`
- Importa `ler_xls`, `extrair_metadados`, `gerar_ofx` de `xls_to_ofx.py` (sem duplicar lógica)
- Classe `App(ctk.CTk)` — janela principal
- Conversão roda em **thread separada** para não travar a UI

---

## Layout

Janela fixa **500 × 300 px**, não redimensionável, tema escuro (CustomTkinter).

```
┌──────────────────────────────────────────────┐
│   Conversor SISPAG → OFX                     │
│   Itaú · Omie                                │
├──────────────────────────────────────────────┤
│  Arquivo XLS/XLSX                            │
│  [____________________________] [Selecionar] │
│                                              │
│  ☐ Incluir lançamentos não efetuados         │
│                                              │
│           [ Converter ]                      │
│                                              │
│  ✅ 55 lançamentos · R$ 544.301,10           │
│     Salvo em: C:\...\saida.ofx              │
└──────────────────────────────────────────────┘
```

---

## Fluxo de uso

1. Usuário clica **Selecionar arquivo** → `filedialog.askopenfilename` (filtro `.xls`, `.xlsx`)
2. Caminho aparece no campo de texto (read-only)
3. (Opcional) marca checkbox "Incluir lançamentos não efetuados"
4. Clica **Converter**:
   - Abre `filedialog.asksaveasfilename` para escolher destino do `.ofx`
   - Se cancelar: conversão abortada silenciosamente
   - Botão desabilitado + texto "Convertendo…" durante o processo
   - **Sucesso:** ✅ verde com contagem de lançamentos e total R$, mais caminho do arquivo
   - **Erro:** ❌ vermelho com mensagem legível (sem stack trace)
5. Selecionar novo arquivo limpa o resultado anterior automaticamente

---

## Distribuição

- Empacotado com **PyInstaller** como `.exe` único (`--onefile`)
- Embute Python + customtkinter + pandas + xlrd + openpyxl
- Usuário final só precisa copiar e abrir o `.exe`
- Tempo de abertura estimado: 8–15 segundos (descompactação interna)

### Comando de build

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "Conversor SISPAG-OFX" gui.py
```

O executável gerado fica em `dist/Conversor SISPAG-OFX.exe`.

---

## Dependências

```
customtkinter
pandas
xlrd
openpyxl
pyinstaller  # apenas para o build
```
