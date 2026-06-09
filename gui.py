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

# Garante que xls_to_ofx.py é encontrado mesmo ao rodar como .exe (PyInstaller)
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
        self.title("Conversor de Planilha → OFX")
        self.geometry("500x300")
        self.resizable(False, False)
        self._arquivo = ""
        self._build_ui()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)

        # Título
        ctk.CTkLabel(
            self, text="Conversor SISPAG → OFX",
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=0, column=0, padx=24, pady=(20, 2), sticky="w")

        ctk.CTkLabel(
            self, text="Itaú · Omie",
            font=ctk.CTkFont(size=12), text_color="gray"
        ).grid(row=1, column=0, padx=24, pady=(0, 14), sticky="w")

        # Linha de seleção de arquivo
        frame_arquivo = ctk.CTkFrame(self, fg_color="transparent")
        frame_arquivo.grid(row=2, column=0, padx=24, sticky="ew")
        frame_arquivo.grid_columnconfigure(0, weight=1)

        self.entry_arquivo = ctk.CTkEntry(
            frame_arquivo,
            placeholder_text="Nenhum arquivo selecionado",
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
            return  # usuário cancelou o diálogo — sem mensagem de erro

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
