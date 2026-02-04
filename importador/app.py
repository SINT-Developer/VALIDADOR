"""
Interface grafica do Importador de Planilhas.
Uso: python app.py
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import threading
import time

# Paths
script_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(script_dir)
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

from motor import PlanilhaImportador
from mapeamento import ORDEM_IMPORTACAO
from planilha_validator import PlanilhaValidator

APP_VERSION = "1.0.0"


class ImportadorApp:

    def __init__(self, root):
        self.root = root
        self.root.title(f"Importador de Planilhas - SINT v{APP_VERSION}")
        self.root.geometry("720x750")
        self.root.resizable(False, False)

        # Variaveis
        self.file_path = tk.StringVar()
        self.servidor = tk.StringVar(value="127.0.0.1")
        self.banco = tk.StringVar(value="SRPP")
        self.usuario = tk.StringVar(value="sa")
        self.senha = tk.StringVar(value="M4573R")
        self.modo_importacao = tk.IntVar(value=2)
        self.excluir_tudo_var = tk.BooleanVar(value=False)
        self.limpar_aux_var = tk.BooleanVar(value=False)
        self.validar_antes = tk.BooleanVar(value=True)
        self.versao_srppwin = tk.StringVar(value="19.1.5")
        self.progress_var = tk.DoubleVar(value=0)
        self.status_var = tk.StringVar(value="Selecione uma planilha para iniciar")

        self.abas_vars = {}
        for aba in ORDEM_IMPORTACAO:
            self.abas_vars[aba] = tk.BooleanVar(value=True)

        self._build_ui()

    # ----------------------------------------------------------
    # UI
    # ----------------------------------------------------------

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main, text="Importador de Planilhas", font=("Arial", 14, "bold")).pack(pady=(0, 8))

        # Arquivo
        f = ttk.LabelFrame(main, text="Planilha", padding=5)
        f.pack(fill=tk.X, pady=(0, 5))
        ttk.Entry(f, textvariable=self.file_path, state="readonly", width=70).pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        ttk.Button(f, text="Procurar", command=self._browse).pack(side=tk.LEFT)

        # Conexao
        f = ttk.LabelFrame(main, text="Conexao SQL Server", padding=5)
        f.pack(fill=tk.X, pady=(0, 5))
        g = ttk.Frame(f)
        g.pack(fill=tk.X)
        campos = [("Servidor:", self.servidor), ("Banco:", self.banco), ("Usuario:", self.usuario)]
        for i, (lbl, var) in enumerate(campos):
            ttk.Label(g, text=lbl, width=10).grid(row=i, column=0, sticky=tk.W)
            ttk.Entry(g, textvariable=var, width=35).grid(row=i, column=1, padx=2, pady=1)
        ttk.Label(g, text="Senha:", width=10).grid(row=3, column=0, sticky=tk.W)
        ttk.Entry(g, textvariable=self.senha, width=35, show="*").grid(row=3, column=1, padx=2, pady=1)
        ttk.Button(g, text="Testar Conexao", command=self._testar_conexao).grid(row=0, column=2, rowspan=2, padx=5)

        # Modo
        f = ttk.LabelFrame(main, text="Modo de Importacao", padding=5)
        f.pack(fill=tk.X, pady=(0, 5))
        for val, txt in [(1, "Inserir novos (erro se ja existe)"),
                         (2, "Sobrescrever (atualiza ou insere)"),
                         (3, "Atualizar existentes (erro se nao existe)")]:
            ttk.Radiobutton(f, text=txt, variable=self.modo_importacao, value=val).pack(anchor=tk.W)

        # Abas
        f = ttk.LabelFrame(main, text="Abas para Importar", padding=5)
        f.pack(fill=tk.X, pady=(0, 5))
        g = ttk.Frame(f)
        g.pack(fill=tk.X)
        for i, aba in enumerate(ORDEM_IMPORTACAO):
            ttk.Checkbutton(g, text=aba, variable=self.abas_vars[aba]).grid(
                row=i // 4, column=i % 4, sticky=tk.W, padx=5
            )
        bf = ttk.Frame(f)
        bf.pack(fill=tk.X, pady=(3, 0))
        ttk.Button(bf, text="Marcar Todas", command=lambda: self._set_abas(True)).pack(side=tk.LEFT, padx=2)
        ttk.Button(bf, text="Desmarcar Todas", command=lambda: self._set_abas(False)).pack(side=tk.LEFT, padx=2)

        # Opcoes
        f = ttk.LabelFrame(main, text="Opcoes", padding=5)
        f.pack(fill=tk.X, pady=(0, 5))
        ttk.Checkbutton(f, text="Validar planilha antes de importar", variable=self.validar_antes,
                        command=self._toggle_versao).pack(anchor=tk.W)
        self.frame_versao = ttk.Frame(f)
        self.frame_versao.pack(anchor=tk.W, padx=(20, 0))
        ttk.Label(self.frame_versao, text="Versao SRPPWin:").pack(side=tk.LEFT)
        self.combo_versao = ttk.Combobox(self.frame_versao, textvariable=self.versao_srppwin,
                                         values=["19.1.5", "20.1.0"], state="readonly", width=10)
        self.combo_versao.pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(f, text="Excluir TUDO antes (pedidos + cadastros + config)", variable=self.excluir_tudo_var).pack(anchor=tk.W)
        ttk.Checkbutton(f, text="Limpar tabelas auxiliares (_imp / _pkunica)", variable=self.limpar_aux_var).pack(anchor=tk.W)

        # Botao
        self.btn_importar = ttk.Button(main, text="IMPORTAR", command=self._iniciar)
        self.btn_importar.pack(fill=tk.X, pady=5)

        # Progresso
        ttk.Progressbar(main, variable=self.progress_var, maximum=100).pack(fill=tk.X, pady=(0, 2))
        ttk.Label(main, textvariable=self.status_var, font=("Arial", 9)).pack(fill=tk.X)

        # Log
        f = ttk.LabelFrame(main, text="Log", padding=5)
        f.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        self.log_text = tk.Text(f, height=10, font=("Consolas", 9), state=tk.DISABLED, wrap=tk.WORD)
        sb = ttk.Scrollbar(f, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=sb.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

    # ----------------------------------------------------------
    # Helpers UI
    # ----------------------------------------------------------

    def _toggle_versao(self):
        if self.validar_antes.get():
            self.frame_versao.pack(anchor=tk.W, padx=(20, 0))
        else:
            self.frame_versao.pack_forget()

    def _set_abas(self, val):
        for v in self.abas_vars.values():
            v.set(val)

    def _browse(self):
        path = filedialog.askopenfilename(
            title="Selecionar planilha",
            filetypes=[("Excel", "*.xlsx *.xls *.xlsm *.xlsb"), ("Todos", "*.*")]
        )
        if path:
            self.file_path.set(path)
            self.status_var.set(f"Arquivo: {os.path.basename(path)}")

    def _conn_str(self):
        return (
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={self.servidor.get()};"
            f"DATABASE={self.banco.get()};"
            f"UID={self.usuario.get()};"
            f"PWD={self.senha.get()};"
            f"TrustServerCertificate=yes;"
        )

    def _testar_conexao(self):
        ok, msg = PlanilhaImportador.testar_conexao(self._conn_str())
        if ok:
            messagebox.showinfo("Conexao", msg)
        else:
            messagebox.showerror("Erro de Conexao", msg)

    def _log(self, msg):
        def _append():
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.insert(tk.END, msg + "\n")
            self.log_text.see(tk.END)
            self.log_text.configure(state=tk.DISABLED)
        self.root.after(0, _append)

    def _progresso(self, value, message):
        def _update():
            self.progress_var.set(value)
            self.status_var.set(message)
            self.root.update_idletasks()
        self.root.after(0, _update)

    def _set_widgets(self, parent, state):
        for child in parent.winfo_children():
            if isinstance(child, (ttk.Button, ttk.Entry, ttk.Checkbutton, ttk.Radiobutton,
                                  tk.Button, tk.Entry, tk.Checkbutton, tk.Radiobutton)):
                try:
                    child.configure(state=state)
                except Exception:
                    pass
            self._set_widgets(child, state)

    # ----------------------------------------------------------
    # Importacao
    # ----------------------------------------------------------

    def _iniciar(self):
        arquivo = self.file_path.get()
        if not arquivo or not os.path.isfile(arquivo):
            messagebox.showerror("Erro", "Selecione um arquivo valido.")
            return

        abas = [a for a, v in self.abas_vars.items() if v.get()]
        if not abas:
            messagebox.showerror("Erro", "Selecione pelo menos uma aba.")
            return

        if self.excluir_tudo_var.get():
            if not messagebox.askyesno("ATENCAO",
                    "Voce marcou 'Excluir TUDO'. Isso vai apagar TODOS os dados "
                    "(pedidos, cadastros, configuracoes).\n\nTem certeza?",
                    icon="warning"):
                return

        # Limpar log
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)

        self._set_widgets(self.root, "disabled")
        self.progress_var.set(0)

        threading.Thread(target=self._executar, args=(arquivo, abas), daemon=True).start()

    def _executar(self, arquivo, abas):
        try:
            # Pre-validacao
            if self.validar_antes.get():
                self._progresso(0, "Validando planilha...")
                self._log("=== PRE-VALIDACAO ===")

                validator = PlanilhaValidator(arquivo, progress_callback=self._progresso)
                validator.versao_srppwin = self.versao_srppwin.get()
                excel_data, nome_arquivo, status, resultados = validator.processar("Validacao Local")

                if status == "reprovado":
                    self._log("REPROVADO - Importacao bloqueada!")
                    for r in resultados:
                        if r.get("erros", 0) > 0:
                            self._log(f"  {r['Planilha']}: {r['erros']} erros")

                    # Salvar planilha com resultados visuais
                    pasta = os.path.dirname(arquivo)
                    output_path = os.path.join(pasta, nome_arquivo)
                    with open(output_path, "wb") as f:
                        f.write(excel_data.getbuffer())
                    self._log(f"  Planilha validada salva em: {output_path}")

                    self._progresso(0, "Bloqueado - planilha reprovada")
                    msg_erro = f"A planilha tem erros e a importacao foi bloqueada.\n\nPlanilha com erros destacados salva em:\n{output_path}"
                    self.root.after(0, lambda: messagebox.showerror(
                        "Validacao Reprovada", msg_erro))
                    return

                if status == "advertencias":
                    self._log("APROVADO COM ADVERTENCIAS")
                    resposta = [None]
                    def _ask():
                        resposta[0] = messagebox.askyesno(
                            "Advertencias",
                            "Planilha aprovada com advertencias.\nContinuar importacao?")
                    self.root.after(0, _ask)
                    while resposta[0] is None:
                        time.sleep(0.1)
                    if not resposta[0]:
                        self._log("Cancelado pelo usuario.")
                        self._progresso(0, "Cancelado")
                        return
                else:
                    self._log("APROVADO")

            # Importar
            self._log("")
            self._log("=== IMPORTACAO ===")
            importador = PlanilhaImportador(
                self._conn_str(),
                progress_callback=self._progresso,
                log_callback=self._log
            )

            resultados = importador.importar(
                arquivo_excel=arquivo,
                abas_selecionadas=abas,
                sobreescreve=self.modo_importacao.get(),
                excluir_tudo=self.excluir_tudo_var.get(),
                limpar_auxiliares=self.limpar_aux_var.get()
            )

            # Resumo
            self._log("")
            self._log("=== RESUMO ===")
            total_s, total_e = 0, 0
            for aba, dados in resultados.items():
                if aba == "ERRO_GERAL":
                    self._log(f"ERRO GERAL: {dados}")
                    continue
                s, e = dados["sucesso"], dados["erros"]
                total_s += s
                total_e += e
                marca = "OK" if e == 0 else f"{e} ERROS"
                self._log(f"  {aba}: {s} importados, {marca}")

            self._log(f"\nTOTAL: {total_s} importados, {total_e} erros")

            if total_e == 0:
                self.root.after(0, lambda: messagebox.showinfo(
                    "Concluido", f"Importacao concluida!\n{total_s} registros importados."))
            else:
                self.root.after(0, lambda: messagebox.showwarning(
                    "Concluido com Erros",
                    f"{total_s} importados, {total_e} erros.\nConsulte o log."))

        except Exception as e:
            self._log(f"ERRO FATAL: {e}")
            import traceback
            traceback.print_exc()
            self.root.after(0, lambda: messagebox.showerror("Erro", str(e)))
        finally:
            self.root.after(0, lambda: self._set_widgets(self.root, "normal"))


if __name__ == "__main__":
    root = tk.Tk()
    app = ImportadorApp(root)
    root.mainloop()
