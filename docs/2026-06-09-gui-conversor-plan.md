# GUI Conversor SISPAG → OFX — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Criar `gui.py` com interface CustomTkinter para converter XLS/XLSX SISPAG em OFX, e empacotá-lo como `.exe` único via PyInstaller.

**Architecture:** Arquivo único `gui.py` na mesma pasta de `xls_to_ofx.py`. A classe `App(ctk.CTk)` monta a janela e delega toda lógica de conversão às funções já existentes em `xls_to_ofx.py`. A conversão roda em thread separada para não travar a UI.

**Tech Stack:** Python 3.8+, customtkinter, tkinter (filedialog), threading, PyInstaller

---

## Mapa de arquivos

| Arquivo | Ação | Responsabilidade |
|---|---|---|
| `gui.py` | Criar | Interface gráfica completa |
| `xls_to_ofx.py` | Sem alteração | Lógica de conversão (já existe) |
| `requirements.txt` | Atualizar | Adicionar customtkinter |
| `dist/Conversor SISPAG-OFX.exe` | Gerado pelo build | Executável final |

---

## Task 1: Instalar dependência e verificar imports

**Files:**
- Modify: `requirements.txt`

- [ ] **Step 1: Instalar customtkinter**

```bash
pip install customtkinter
```

Esperado: `Successfully installed customtkinter-x.x.x`

- [ ] **Step 2: Verificar se o import funciona**

```bash
python -c "import customtkinter; print(customtkinter.__version__)"
```

Esperado: número de versão impresso sem erro.

- [ ] **Step 3: Atualizar requirements.txt**

Abrir `requirements.txt` e adicionar:
```
customtkinter>=5.2.0
```

---

## Task 2: Janela base e estrutura da classe

**Files:**
- Create: `gui.py`

- [ ] **Step 1: Criar gui.py com janela mínima abrindo corretamente**

```python
"""
Interface gráfica para o Conversor SISPAG → OFX.
Uso: python gui.py
Dependências: pip install customtkinter pandas xlrd openpyxl
"""

import customtkinter as ctk
from tkinter import filedialog
import threading
import os
import sys

# Garante que xls_to_ofx.py é encontrado mesmo ao rodar como .exe (PyInstaller)
if getattr(sys, "frozen", False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from xls_to_ofx import ler_xls, extrair_metadados, gerar_ofx

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Conversor SISPAG → OFX")
        self.geometry("500x300")
        self.resizable(False, False)
        self._build_ui()

    def _build_ui(self):
        pass  # implementado nas próximas tasks


if __name__ == "__main__":
    app = App()
    app.mainloop()
```

- [ ] **Step 2: Executar e confirmar que a janela abre e fecha sem erro**

```bash
python gui.py
```

Esperado: janela vazia 500×300 px abre, fecha normalmente.

---

## Task 3: Título, campo de arquivo e botão Selecionar

**Files:**
- Modify: `gui.py` — substituir `_build_ui`

- [ ] **Step 1: Implementar _build_ui com título e seleção de arquivo**

Substituir o método `_build_ui` e adicionar `_selecionar_arquivo`:

```python
def _build_ui(self):
    self.grid_columnconfigure(0, weight=1)

    # Título
    ctk.CTkLabel(
        self, text="Conversor SISPAG → OFX",
        font=ctk.CTkFont(size=18, weight="bold")
    ).grid(row=0, column=0, padx=24, pady=(20, 2), sticky="w")

    ctk.CTkLabel(
        self, text="Itaú · Omie",
        font=ctk.CTkFont(size=12),
        text_color="gray"
    ).grid(row=1, column=0, padx=24, pady=(0, 14), sticky="w")

    # Linha de seleção de arquivo
    frame_arquivo = ctk.CTkFrame(self, fg_color="transparent")
    frame_arquivo.grid(row=2, column=0, padx=24, sticky="ew")
    frame_arquivo.grid_columnconfigure(0, weight=1)

    self.entry_arquivo = ctk.CTkEntry(
        frame_arquivo, placeholder_text="Nenhum arquivo selecionado",
        state="disabled"
    )
    self.entry_arquivo.grid(row=0, column=0, sticky="ew", padx=(0, 8))

    ctk.CTkButton(
        frame_arquivo, text="Selecionar arquivo",
        width=140, command=self._selecionar_arquivo
    ).grid(row=0, column=1)

    # Checkbox
    self.var_nao_efetuado = ctk.BooleanVar(value=False)
    ctk.CTkCheckBox(
        self, text="Incluir lançamentos não efetuados",
        variable=self.var_nao_efetuado
    ).grid(row=3, column=0, padx=24, pady=(14, 0), sticky="w")

    # Botão converter
    self.btn_converter = ctk.CTkButton(
        self, text="Converter", width=160,
        state="disabled", command=self._iniciar_conversao
    )
    self.btn_converter.grid(row=4, column=0, pady=(16, 0))

    # Status
    self.lbl_status = ctk.CTkLabel(
        self, text="", wraplength=450,
        font=ctk.CTkFont(size=12)
    )
    self.lbl_status.grid(row=5, column=0, padx=24, pady=(10, 0), sticky="w")


def _selecionar_arquivo(self):
    caminho = filedialog.askopenfilename(
        title="Selecionar relatório SISPAG",
        filetypes=[("Planilha Excel", "*.xls *.xlsx"), ("Todos os arquivos", "*.*")]
    )
    if not caminho:
        return
    self._arquivo = caminho
    # Atualiza entry (reabilita temporariamente para inserir texto)
    self.entry_arquivo.configure(state="normal")
    self.entry_arquivo.delete(0, "end")
    self.entry_arquivo.insert(0, caminho)
    self.entry_arquivo.configure(state="disabled")
    # Habilita botão e limpa status anterior
    self.btn_converter.configure(state="normal")
    self.lbl_status.configure(text="")
```

Adicionar também `self._arquivo = ""` no `__init__`, logo antes de `self._build_ui()`:

```python
self._arquivo = ""
```

- [ ] **Step 2: Executar e testar seleção de arquivo**

```bash
python gui.py
```

Verificar:
- Botão "Selecionar arquivo" abre o diálogo de arquivos
- Caminho aparece no campo após seleção
- Botão "Converter" fica habilitado
- Selecionar outro arquivo limpa o status e atualiza o campo

---

## Task 4: Conversão em thread + diálogo de salvamento + status

**Files:**
- Modify: `gui.py` — adicionar `_iniciar_conversao` e `_executar_conversao`

- [ ] **Step 1: Implementar os métodos de conversão**

Adicionar à classe `App`:

```python
def _iniciar_conversao(self):
    caminho_saida = filedialog.asksaveasfilename(
        title="Salvar arquivo OFX",
        defaultextension=".ofx",
        initialfile=os.path.splitext(os.path.basename(self._arquivo))[0] + ".ofx",
        filetypes=[("Arquivo OFX", "*.ofx")]
    )
    if not caminho_saida:
        return  # usuário cancelou o diálogo

    self.btn_converter.configure(state="disabled", text="Convertendo…")
    self.lbl_status.configure(text="", text_color="gray")

    thread = threading.Thread(
        target=self._executar_conversao,
        args=(self._arquivo, caminho_saida, self.var_nao_efetuado.get()),
        daemon=True
    )
    thread.start()


def _executar_conversao(self, caminho_entrada, caminho_saida, incluir_nao_efetuado):
    try:
        meta = extrair_metadados(caminho_entrada)
        df, stats = ler_xls(caminho_entrada, apenas_efetuados=not incluir_nao_efetuado)

        if stats["incluidos"] == 0:
            raise ValueError("Nenhum lançamento válido encontrado no arquivo.")

        ofx = gerar_ofx(df, agencia=meta["agencia"], conta=meta["conta"])

        with open(caminho_saida, "w", encoding="latin-1", errors="replace") as f:
            f.write(ofx)

        total = df["valor"].apply(lambda v: float(str(v).replace(",", ".")) if str(v).replace(",", ".").replace(".", "", 1).lstrip("-").isdigit() or True else 0).sum()
        total = df["valor"].apply(lambda v: __import__("xls_to_ofx").to_float(v) or 0.0).sum()

        msg = (
            f"✅  {stats['incluidos']} lançamentos  ·  "
            f"R$ {total:,.2f}\n"
            f"Salvo em: {caminho_saida}"
        )
        self.after(0, self._mostrar_sucesso, msg)

    except Exception as e:
        self.after(0, self._mostrar_erro, str(e))


def _mostrar_sucesso(self, msg):
    self.lbl_status.configure(text=msg, text_color="#4CAF50")
    self.btn_converter.configure(state="normal", text="Converter")


def _mostrar_erro(self, msg):
    self.lbl_status.configure(text=f"❌  {msg}", text_color="#F44336")
    self.btn_converter.configure(state="normal", text="Converter")
```

> **Nota sobre `to_float`:** a função `to_float` está em `xls_to_ofx.py`. O cálculo do total na thread deve usar:
> ```python
> from xls_to_ofx import to_float
> total = df["valor"].apply(lambda v: to_float(v) or 0.0).sum()
> ```
> Adicionar esse import no topo do arquivo junto com os demais.

- [ ] **Step 2: Corrigir o cálculo do total no método `_executar_conversao`**

O trecho com o cálculo duplicado/incorreto deve ficar assim (limpo):

```python
total = df["valor"].apply(lambda v: to_float(v) or 0.0).sum()
msg = (
    f"✅  {stats['incluidos']} lançamentos  ·  "
    f"R$ {total:,.2f}\n"
    f"Salvo em: {caminho_saida}"
)
```

- [ ] **Step 3: Atualizar o import no topo de gui.py**

```python
from xls_to_ofx import ler_xls, extrair_metadados, gerar_ofx, to_float
```

- [ ] **Step 4: Testar o fluxo completo**

```bash
python gui.py
```

Verificar:
- Selecionar `ConsultaPagamentos.xls`
- Clicar Converter → diálogo de salvamento abre com nome sugerido
- Escolher destino → status mostra "Convertendo…" brevemente
- Status final: ✅ verde com número de lançamentos e total
- Abrir o `.ofx` gerado e confirmar que tem conteúdo válido

- [ ] **Step 5: Testar caminho de erro**

Renomear temporariamente uma coluna no XLS ou passar um arquivo inválido.
Esperado: status ❌ vermelho com mensagem legível, sem crash.

---

## Task 5: Build do executável com PyInstaller

**Files:**
- Gera: `dist/Conversor SISPAG-OFX.exe`

- [ ] **Step 1: Instalar PyInstaller**

```bash
pip install pyinstaller
```

- [ ] **Step 2: Rodar o build**

```bash
cd "C:\Users\Marcus Suporte\Documents\Conversor"
pyinstaller --onefile --windowed --name "Conversor SISPAG-OFX" --add-data "xls_to_ofx.py;." gui.py
```

> `--windowed` suprime o terminal preto ao abrir no Windows.  
> `--add-data "xls_to_ofx.py;."` inclui o módulo de conversão no bundle.

Esperado: pasta `dist/` criada com `Conversor SISPAG-OFX.exe` (~30–60 MB).

- [ ] **Step 3: Testar o .exe gerado**

Abrir `dist/Conversor SISPAG-OFX.exe` (aguardar 8–15 s na primeira vez).
Realizar uma conversão completa pelo executável.
Esperado: mesmo comportamento da versão Python.

- [ ] **Step 4: Verificar que o .exe roda sem Python instalado**

Se possível, copiar o `.exe` para uma pasta limpa ou outra máquina e confirmar que abre sem instalar nada.

---

## Arquivo gui.py final (referência completa)

```python
"""
Interface gráfica para o Conversor SISPAG → OFX.
Uso: python gui.py  |  ou abrir Conversor SISPAG-OFX.exe
Dependências: pip install customtkinter pandas xlrd openpyxl
"""

import customtkinter as ctk
from tkinter import filedialog
import threading
import os
import sys

if getattr(sys, "frozen", False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from xls_to_ofx import ler_xls, extrair_metadados, gerar_ofx, to_float

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Conversor SISPAG → OFX")
        self.geometry("500x300")
        self.resizable(False, False)
        self._arquivo = ""
        self._build_ui()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self, text="Conversor SISPAG → OFX",
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=0, column=0, padx=24, pady=(20, 2), sticky="w")

        ctk.CTkLabel(
            self, text="Itaú · Omie",
            font=ctk.CTkFont(size=12), text_color="gray"
        ).grid(row=1, column=0, padx=24, pady=(0, 14), sticky="w")

        frame_arquivo = ctk.CTkFrame(self, fg_color="transparent")
        frame_arquivo.grid(row=2, column=0, padx=24, sticky="ew")
        frame_arquivo.grid_columnconfigure(0, weight=1)

        self.entry_arquivo = ctk.CTkEntry(
            frame_arquivo, placeholder_text="Nenhum arquivo selecionado",
            state="disabled"
        )
        self.entry_arquivo.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        ctk.CTkButton(
            frame_arquivo, text="Selecionar arquivo",
            width=140, command=self._selecionar_arquivo
        ).grid(row=0, column=1)

        self.var_nao_efetuado = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            self, text="Incluir lançamentos não efetuados",
            variable=self.var_nao_efetuado
        ).grid(row=3, column=0, padx=24, pady=(14, 0), sticky="w")

        self.btn_converter = ctk.CTkButton(
            self, text="Converter", width=160,
            state="disabled", command=self._iniciar_conversao
        )
        self.btn_converter.grid(row=4, column=0, pady=(16, 0))

        self.lbl_status = ctk.CTkLabel(
            self, text="", wraplength=450,
            font=ctk.CTkFont(size=12)
        )
        self.lbl_status.grid(row=5, column=0, padx=24, pady=(10, 0), sticky="w")

    def _selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar relatório SISPAG",
            filetypes=[("Planilha Excel", "*.xls *.xlsx"), ("Todos os arquivos", "*.*")]
        )
        if not caminho:
            return
        self._arquivo = caminho
        self.entry_arquivo.configure(state="normal")
        self.entry_arquivo.delete(0, "end")
        self.entry_arquivo.insert(0, caminho)
        self.entry_arquivo.configure(state="disabled")
        self.btn_converter.configure(state="normal")
        self.lbl_status.configure(text="")

    def _iniciar_conversao(self):
        caminho_saida = filedialog.asksaveasfilename(
            title="Salvar arquivo OFX",
            defaultextension=".ofx",
            initialfile=os.path.splitext(os.path.basename(self._arquivo))[0] + ".ofx",
            filetypes=[("Arquivo OFX", "*.ofx")]
        )
        if not caminho_saida:
            return
        self.btn_converter.configure(state="disabled", text="Convertendo…")
        self.lbl_status.configure(text="", text_color="gray")
        thread = threading.Thread(
            target=self._executar_conversao,
            args=(self._arquivo, caminho_saida, self.var_nao_efetuado.get()),
            daemon=True
        )
        thread.start()

    def _executar_conversao(self, caminho_entrada, caminho_saida, incluir_nao_efetuado):
        try:
            meta = extrair_metadados(caminho_entrada)
            df, stats = ler_xls(caminho_entrada, apenas_efetuados=not incluir_nao_efetuado)
            if stats["incluidos"] == 0:
                raise ValueError("Nenhum lançamento válido encontrado no arquivo.")
            ofx = gerar_ofx(df, agencia=meta["agencia"], conta=meta["conta"])
            with open(caminho_saida, "w", encoding="latin-1", errors="replace") as f:
                f.write(ofx)
            total = df["valor"].apply(lambda v: to_float(v) or 0.0).sum()
            msg = (
                f"✅  {stats['incluidos']} lançamentos  ·  R$ {total:,.2f}\n"
                f"Salvo em: {caminho_saida}"
            )
            self.after(0, self._mostrar_sucesso, msg)
        except Exception as e:
            self.after(0, self._mostrar_erro, str(e))

    def _mostrar_sucesso(self, msg):
        self.lbl_status.configure(text=msg, text_color="#4CAF50")
        self.btn_converter.configure(state="normal", text="Converter")

    def _mostrar_erro(self, msg):
        self.lbl_status.configure(text=f"❌  {msg}", text_color="#F44336")
        self.btn_converter.configure(state="normal", text="Converter")


if __name__ == "__main__":
    app = App()
    app.mainloop()
```
