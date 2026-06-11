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

from xls_to_ofx import ler_xls, extrair_metadados, gerar_ofx
from ofx_reader import ler_ofx_banco
from matching import casar_lancamentos, validar_fechamento

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

VERDE, VERMELHO, AMARELO = "#4CAF50", "#F44336", "#F5C800"


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Conversor SISPAG → OFX")
        self.geometry("720x620")
        self.minsize(720, 520)
        self._xls = ""
        self._ofx_banco = ""
        self._analise = None
        self._build_ui()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)

        ctk.CTkLabel(self, text="Conversor SISPAG → OFX",
                     font=ctk.CTkFont(size=18, weight="bold")
                     ).grid(row=0, column=0, padx=24, pady=(18, 0), sticky="w")
        ctk.CTkLabel(self, text="Itaú · Omie — matching contra o extrato",
                     font=ctk.CTkFont(size=12), text_color="gray"
                     ).grid(row=1, column=0, padx=24, pady=(0, 10), sticky="w")

        self.entry_xls, _ = self._linha_arquivo(
            row=2, rotulo="Planilha SISPAG (obrigatório)",
            comando=self._sel_xls)
        self.entry_ofx, _ = self._linha_arquivo(
            row=3, rotulo="OFX do banco (opcional — ativa o matching)",
            comando=self._sel_ofx)

        self.txt = ctk.CTkTextbox(self, font=ctk.CTkFont(family="Consolas", size=12),
                                  state="disabled", wrap="none")
        self.txt.grid(row=4, column=0, padx=24, pady=(12, 8), sticky="nsew")

        frame_botoes = ctk.CTkFrame(self, fg_color="transparent")
        frame_botoes.grid(row=5, column=0, pady=(0, 16))

        self.var_nao_efetuado = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(frame_botoes, text="Incluir não efetuados",
                        variable=self.var_nao_efetuado
                        ).grid(row=0, column=0, padx=(0, 16))

        self.btn_analisar = ctk.CTkButton(frame_botoes, text="Analisar",
                                          width=140, state="disabled",
                                          command=self._analisar)
        self.btn_analisar.grid(row=0, column=1, padx=(0, 12))

        self.btn_gerar = ctk.CTkButton(frame_botoes, text="Gerar OFX",
                                       width=200, state="disabled",
                                       command=self._gerar)
        self.btn_gerar.grid(row=0, column=2)

    def _linha_arquivo(self, row, rotulo, comando):
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.grid(row=row, column=0, padx=24, pady=(4, 0), sticky="ew")
        frame.grid_columnconfigure(0, weight=1)
        entry = ctk.CTkEntry(frame, placeholder_text=rotulo, state="disabled")
        entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        btn = ctk.CTkButton(frame, text="Selecionar", width=110, command=comando)
        btn.grid(row=0, column=1)
        return entry, btn

    def _set_entry(self, entry, texto):
        entry.configure(state="normal")
        entry.delete(0, "end")
        entry.insert(0, texto)
        entry.configure(state="disabled")

    def _set_txt(self, texto):
        self.txt.configure(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.insert("1.0", texto)
        self.txt.configure(state="disabled")

    def _sel_xls(self):
        c = filedialog.askopenfilename(
            title="Selecionar planilha SISPAG",
            filetypes=[("Planilha Excel", "*.xls *.xlsx"), ("Todos", "*.*")])
        if not c:
            return
        self._xls = c
        self._set_entry(self.entry_xls, c)
        self.btn_analisar.configure(state="normal")
        self.btn_gerar.configure(state="disabled", text="Gerar OFX")
        self._analise = None
        self._set_txt("")

    def _sel_ofx(self):
        c = filedialog.askopenfilename(
            title="Selecionar OFX do banco",
            filetypes=[("Arquivo OFX", "*.ofx"), ("Todos", "*.*")])
        if not c:
            return
        self._ofx_banco = c
        self._set_entry(self.entry_ofx, c)
        self.btn_gerar.configure(state="disabled", text="Gerar OFX")
        self._analise = None

    def _analisar(self):
        self.btn_analisar.configure(state="disabled", text="Analisando…")
        self.btn_gerar.configure(state="disabled")
        threading.Thread(target=self._executar_analise, daemon=True).start()

    def _executar_analise(self):
        try:
            if not os.path.isfile(self._xls):
                raise ValueError("Planilha não encontrada. Selecione novamente.")

            meta = extrair_metadados(self._xls)
            df, stats = ler_xls(
                self._xls, apenas_efetuados=not self.var_nao_efetuado.get())
            if stats["incluidos"] == 0:
                raise ValueError("Nenhum lançamento válido na planilha.")

            linhas = ["═══ PLANILHA ═══",
                      f"Lidos: {stats['total_lido']}  |  "
                      f"Sem data: {stats['sem_data']}  |  "
                      f"Não efetuados: {stats['nao_efetuados']}  |  "
                      f"Inválidos: {stats['invalidos']}  |  "
                      f"Tributos removidos: {stats.get('tributos', 0)}",
                      f"Válidos: {stats['incluidos']}", ""]
            if stats.get("tributos", 0):
                linhas.insert(1, "⚠ A planilha continha tributos — foram removidos. "
                                 "Exporte sem tributos para evitar divergência.")

            divergencias = False
            if self._ofx_banco:
                banco = ler_ofx_banco(self._ofx_banco)
                casados, nao_casados = casar_lancamentos(df, banco["detalhados"])
                fech = validar_fechamento(df, nao_casados,
                                          banco["agrupados"], banco["dias"])
                indices_gerar = sorted(nao_casados)

                linhas += ["═══ MATCHING (OFX do banco) ═══",
                           f"Detalhados no extrato: {len(banco['detalhados'])}  |  "
                           f"Agrupados SISPAG: {len(banco['agrupados'])}  |  "
                           f"Tributos: {len(banco['tributos'])}  |  "
                           f"Créditos ignorados: {banco['n_creditos']}",
                           f"Já conciliados via banco (excluídos): {len(casados)}",
                           f"Entram no OFX gerado: {len(nao_casados)}", "",
                           "═══ FECHAMENTO POR DIA ═══"]
                for dia, r in fech.items():
                    icone = {"ok": "✅", "divergente": "❌",
                             "fora_do_extrato": "⚠"}[r["status"]]
                    linhas.append(
                        f"{icone} {dia}  não-casados R$ {r['soma_nao_casados']:>12,.2f}"
                        f"  |  agrupados R$ {r['soma_agrupados']:>12,.2f}"
                        f"  |  dif R$ {r['diferenca']:,.2f}")
                    if r["status"] == "divergente":
                        divergencias = True
                if any(r["status"] == "fora_do_extrato" for r in fech.values()):
                    linhas.append("⚠ Dias fora do período do extrato não podem ser "
                                  "validados — exporte XLS e OFX com o mesmo período.")
                linhas += ["", "── Excluídos (já no extrato) ──"]
                for idx in sorted(casados):
                    r = df.loc[idx]
                    linhas.append(f"  {r['data']}  R$ {r['valor_num']:>12,.2f}  "
                                  f"{str(r['favorecido'])[:40]}")
                linhas += ["", "── Entram no OFX ──"]
            else:
                indices_gerar = list(df.index)
                linhas += ["(Sem OFX do banco — todos os lançamentos entram. "
                           "Use o matching para evitar duplicidade no Omie.)", "",
                           "── Entram no OFX ──"]

            for idx in indices_gerar:
                r = df.loc[idx]
                linhas.append(f"  {r['data']}  R$ {r['valor_num']:>12,.2f}  "
                              f"{str(r['favorecido'])[:40]}")
            total = df.loc[indices_gerar, "valor_num"].sum()
            linhas += ["", f"TOTAL A GERAR: {len(indices_gerar)} lançamentos  ·  "
                           f"R$ {total:,.2f}"]

            self._analise = {"df": df, "indices": indices_gerar, "meta": meta,
                             "divergencias": divergencias}
            self.after(0, self._fim_analise, "\n".join(linhas))
        except Exception as e:
            self.after(0, self._erro_analise, str(e))

    def _fim_analise(self, relatorio):
        self._set_txt(relatorio)
        self.btn_analisar.configure(state="normal", text="Analisar")
        n = len(self._analise["indices"])
        if self._analise["divergencias"]:
            self.btn_gerar.configure(state="normal", fg_color=VERMELHO,
                                     hover_color="#B71C1C",
                                     text=f"⚠ Gerar mesmo assim ({n})")
        else:
            self.btn_gerar.configure(state="normal",
                                     text=f"Gerar OFX ({n} lançamentos)")

    def _erro_analise(self, msg):
        self._set_txt(f"❌ {msg}")
        self.btn_analisar.configure(state="normal", text="Analisar")

    def _gerar(self):
        saida = filedialog.asksaveasfilename(
            title="Salvar arquivo OFX", defaultextension=".ofx",
            initialfile=os.path.splitext(os.path.basename(self._xls))[0] + ".ofx",
            filetypes=[("Arquivo OFX", "*.ofx")])
        if not saida:
            return
        a = self._analise
        df_gerar = a["df"].loc[a["indices"]].reset_index(drop=True)
        ofx = gerar_ofx(df_gerar, agencia=a["meta"]["agencia"],
                        conta=a["meta"]["conta"])
        with open(saida, "w", encoding="cp1252", errors="replace") as f:
            f.write(ofx)
        total = df_gerar["valor_num"].sum()
        self._set_txt(f"✅ OFX gerado: {len(df_gerar)} lançamentos · "
                      f"R$ {total:,.2f}\nSalvo em: {saida}")
        self.btn_gerar.configure(state="disabled", text="Gerar OFX",
                                 fg_color=["#3B8ED0", "#1F6AA5"])


if __name__ == "__main__":
    app = App()
    app.mainloop()
