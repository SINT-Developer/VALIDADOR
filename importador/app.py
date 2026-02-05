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
import pyodbc

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

APP_VERSION = "1.1.0"

# Lista de drivers ODBC para tentar (ordem de preferencia)
ODBC_DRIVERS = [
    "ODBC Driver 18 for SQL Server",
    "ODBC Driver 17 for SQL Server",
    "ODBC Driver 13.1 for SQL Server",
    "ODBC Driver 13 for SQL Server",
    "ODBC Driver 11 for SQL Server",
    "SQL Server Native Client 11.0",
    "SQL Server Native Client 10.0",
    "SQL Server Native Client",
    "SQL Server",
]

# Agrupamento de abas por dependencia
GRUPOS_ABAS = [
    {
        "nome": "Configuracao",
        "abas": ["EMPRESA"]
    },
    {
        "nome": "Clientes",
        "abas": ["CLIENTES", "ESTADOS", "REPR", "TRANSP"]
    },
    {
        "nome": "Produtos",
        "abas": ["PRODUTOS", "FILIAL", "FAMILIAS", "ESTILOS"]
    },
    {
        "nome": "Pagamento",
        "abas": ["PAGTO", "PAGTOFILIAL"]
    },
]


class ImportadorApp:

    def __init__(self, root):
        self.root = root
        self.root.title(f"Importador de Planilhas - SINT v{APP_VERSION}")
        self.root.geometry("700x620")
        self.root.minsize(600, 500)  # Tamanho minimo

        # Variaveis
        self.file_path = tk.StringVar()
        self.servidor = tk.StringVar(value="127.0.0.1")
        self.banco = tk.StringVar(value="SRPP")
        self.usuario = tk.StringVar(value="sa")
        self.senha = tk.StringVar(value="M4573R")
        self.modo_importacao = tk.IntVar(value=2)
        self.excluir_tudo_var = tk.BooleanVar(value=False)
        self.validar_antes = tk.BooleanVar(value=False)  # Desmarcado por padrao
        self.versao_srppwin = tk.StringVar(value="19.1.5")
        self.progress_var = tk.DoubleVar(value=0)
        self.status_var = tk.StringVar(value="Selecione uma planilha para iniciar")

        self.abas_vars = {}
        for aba in ORDEM_IMPORTACAO:
            self.abas_vars[aba] = tk.BooleanVar(value=True)

        # Driver ODBC detectado (None = ainda nao detectado)
        self.odbc_driver = None

        # Dados para as abas de resultado
        self.resultado_resumo = {}
        self.resultado_detalhes = []
        self.resultado_erros = []

        self._build_ui()

    # ----------------------------------------------------------
    # UI
    # ----------------------------------------------------------

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=8)
        main.pack(fill=tk.BOTH, expand=True)

        # Configurar para expandir
        main.columnconfigure(0, weight=1)
        #main.rowconfigure(6, weight=1)  # Linha do resultado expande

        ttk.Label(main, text="Importador de Planilhas", font=("Arial", 12, "bold")).grid(row=0, column=0, pady=(0, 5), sticky=tk.W)

        # Arquivo
        f = ttk.LabelFrame(main, text="Planilha", padding=4)
        f.grid(row=1, column=0, sticky=tk.EW, pady=(0, 4))
        f.columnconfigure(0, weight=1)
        ttk.Entry(f, textvariable=self.file_path, state="readonly").grid(row=0, column=0, sticky=tk.EW, padx=(0, 5))
        ttk.Button(f, text="Procurar", command=self._browse).grid(row=0, column=1)

        # Conexao - Layout compacto 2x2
        f = ttk.LabelFrame(main, text="Conexao SQL Server", padding=4)
        f.grid(row=2, column=0, sticky=tk.EW, pady=(0, 4))
        g = ttk.Frame(f)
        g.pack(fill=tk.X)

        ttk.Label(g, text="Servidor:").grid(row=0, column=0, sticky=tk.W, padx=(0, 2))
        ttk.Entry(g, textvariable=self.servidor, width=18).grid(row=0, column=1, padx=(0, 10))
        ttk.Label(g, text="Banco:").grid(row=0, column=2, sticky=tk.W, padx=(0, 2))
        ttk.Entry(g, textvariable=self.banco, width=18).grid(row=0, column=3, padx=(0, 10))

        ttk.Label(g, text="Usuario:").grid(row=1, column=0, sticky=tk.W, padx=(0, 2), pady=(3, 0))
        ttk.Entry(g, textvariable=self.usuario, width=18).grid(row=1, column=1, padx=(0, 10), pady=(3, 0))
        ttk.Label(g, text="Senha:").grid(row=1, column=2, sticky=tk.W, padx=(0, 2), pady=(3, 0))
        ttk.Entry(g, textvariable=self.senha, width=18, show="*").grid(row=1, column=3, padx=(0, 10), pady=(3, 0))

        ttk.Button(g, text="Testar", command=self._testar_conexao).grid(row=0, column=4, rowspan=2, padx=(5, 0))

        # Modo - Apenas 2 opcoes
        f = ttk.LabelFrame(main, text="Modo", padding=4)
        f.grid(row=3, column=0, sticky=tk.EW, pady=(0, 4))
        g = ttk.Frame(f)
        g.pack(fill=tk.X)
        ttk.Radiobutton(g, text="Sobrescrever (atualiza ou insere)", variable=self.modo_importacao, value=2).pack(side=tk.LEFT, padx=(0, 20))
        ttk.Radiobutton(g, text="Apenas inserir novos", variable=self.modo_importacao, value=1).pack(side=tk.LEFT)

        # Abas - Agrupadas por dependencia
        f = ttk.LabelFrame(main, text="Abas para Importar", padding=4)
        f.grid(row=4, column=0, sticky=tk.EW, pady=(0, 4))

        # Grid de grupos
        grupos_frame = ttk.Frame(f)
        grupos_frame.pack(fill=tk.X)

        for idx, grupo in enumerate(GRUPOS_ABAS):
            gf = ttk.LabelFrame(grupos_frame, text=grupo["nome"], padding=2)
            gf.grid(row=0, column=idx, padx=3, pady=2, sticky=tk.N)

            for aba in grupo["abas"]:
                if aba in self.abas_vars:
                    ttk.Checkbutton(gf, text=aba, variable=self.abas_vars[aba]).pack(anchor=tk.W)

        # Botoes marcar/desmarcar
        bf = ttk.Frame(f)
        bf.pack(fill=tk.X, pady=(4, 0))
        ttk.Button(bf, text="Todas", command=lambda: self._set_abas(True), width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(bf, text="Nenhuma", command=lambda: self._set_abas(False), width=8).pack(side=tk.LEFT, padx=2)

        # Opcoes - Compacto com versao que some
        f = ttk.LabelFrame(main, text="Opcoes", padding=4)
        f.grid(row=5, column=0, sticky=tk.EW, pady=(0, 4))
        g = ttk.Frame(f)
        g.pack(fill=tk.X)

        ttk.Checkbutton(g, text="Validar antes", variable=self.validar_antes,
                       command=self._toggle_versao).pack(side=tk.LEFT)

        # Frame da versao (inicialmente oculto)
        self.frame_versao = ttk.Frame(g)
        ttk.Label(self.frame_versao, text="Versao:").pack(side=tk.LEFT, padx=(5, 2))
        ttk.Combobox(self.frame_versao, textvariable=self.versao_srppwin,
                    values=["19.1.5", "20.1.0"], state="readonly", width=8).pack(side=tk.LEFT)
        # Nao mostra inicialmente pois validar_antes = False

        ttk.Checkbutton(g, text="Excluir TUDO antes (backup automatico)",
                       variable=self.excluir_tudo_var).pack(side=tk.LEFT, padx=(20, 0))

        # Botao e Progresso
        bf = ttk.Frame(main, padding=8)
        bf.grid(row=6, column=0, sticky=tk.EW, pady=(8, 6))
        bf.columnconfigure(0, weight=1)

        self.btn_importar = ttk.Button(
            bf,
            text="IMPORTAR",
            command=self._iniciar
        )
        self.btn_importar.grid(row=0, column=0, sticky=tk.EW, ipady=8)


        pf = ttk.Frame(bf)
        pf.grid(row=1, column=0, sticky=tk.EW, pady=(4, 0))
        pf.columnconfigure(0, weight=1)
        ttk.Progressbar(pf, variable=self.progress_var, maximum=100).grid(row=0, column=0, sticky=tk.EW, padx=(0, 5))
        self.lbl_status = ttk.Label(pf, textvariable=self.status_var, font=("Arial", 8), width=35, anchor=tk.W)
        self.lbl_status.grid(row=0, column=1)

        # Resultado com Abas (expande)
        f = ttk.LabelFrame(main, text="Resultado", padding=4)
        f.grid(row=7, column=0, sticky=tk.NSEW, pady=(0, 0))
        f.columnconfigure(0, weight=1)
        f.rowconfigure(0, weight=1)

        self.notebook = ttk.Notebook(f)
        self.notebook.grid(row=0, column=0, sticky=tk.NSEW)

        # Aba Resumo
        self.frame_resumo = ttk.Frame(self.notebook)
        self.notebook.add(self.frame_resumo, text="Resumo")
        self._build_resumo_grid()

        # Aba Detalhes
        self.frame_detalhes = ttk.Frame(self.notebook)
        self.notebook.add(self.frame_detalhes, text="Detalhes")
        self.frame_detalhes.columnconfigure(0, weight=1)
        self.frame_detalhes.rowconfigure(0, weight=1)
        self.text_detalhes = tk.Text(self.frame_detalhes, font=("Consolas", 9), state=tk.DISABLED, wrap=tk.WORD)
        sb_det = ttk.Scrollbar(self.frame_detalhes, orient=tk.VERTICAL, command=self.text_detalhes.yview)
        self.text_detalhes.configure(yscrollcommand=sb_det.set)
        self.text_detalhes.grid(row=0, column=0, sticky=tk.NSEW)
        sb_det.grid(row=0, column=1, sticky=tk.NS)

        # Aba Erros
        self.frame_erros = ttk.Frame(self.notebook)
        self.notebook.add(self.frame_erros, text="Erros (0)")
        self.frame_erros.columnconfigure(0, weight=1)
        self.frame_erros.rowconfigure(0, weight=1)
        self.text_erros = tk.Text(self.frame_erros, font=("Consolas", 9), state=tk.DISABLED, wrap=tk.WORD, fg="#CC0000")
        sb_err = ttk.Scrollbar(self.frame_erros, orient=tk.VERTICAL, command=self.text_erros.yview)
        self.text_erros.configure(yscrollcommand=sb_err.set)
        self.text_erros.grid(row=0, column=0, sticky=tk.NSEW)
        sb_err.grid(row=0, column=1, sticky=tk.NS)

        # Configurar peso das linhas para expansao
        main.rowconfigure(7, weight=1)

    def _toggle_versao(self):
        """Mostra/esconde o seletor de versao baseado no checkbox validar."""
        if self.validar_antes.get():
            self.frame_versao.pack(side=tk.LEFT)
        else:
            self.frame_versao.pack_forget()

    def _build_resumo_grid(self):
        """Constroi o grid de resumo com checkmarks."""
        # Limpar frame
        for w in self.frame_resumo.winfo_children():
            w.destroy()

        self.frame_resumo.columnconfigure(0, weight=1)
        self.frame_resumo.rowconfigure(0, weight=1)

        # Canvas com scroll
        canvas = tk.Canvas(self.frame_resumo, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.frame_resumo, orient="vertical", command=canvas.yview)
        self.resumo_inner = ttk.Frame(canvas)

        self.resumo_inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.resumo_inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.grid(row=0, column=0, sticky=tk.NSEW)
        scrollbar.grid(row=0, column=1, sticky=tk.NS)

        # Labels de resumo (vazios inicialmente)
        self.resumo_labels = {}

    def _atualizar_resumo(self):
        """Atualiza o grid de resumo com os resultados."""
        # Limpar inner frame
        for w in self.resumo_inner.winfo_children():
            w.destroy()

        if not self.resultado_resumo:
            ttk.Label(self.resumo_inner, text="Nenhuma importacao realizada ainda.",
                     font=("Arial", 9), foreground="gray").pack(pady=20)
            return

        # Grid de resultados
        row = 0
        col = 0
        total_ok = 0
        total_erros = 0

        for aba, dados in self.resultado_resumo.items():
            if aba == "ERRO_GERAL":
                continue

            sucesso = dados.get("sucesso", 0)
            erros = dados.get("erros", 0)
            total = sucesso + erros
            total_ok += sucesso
            total_erros += erros

            # Frame para cada aba
            af = ttk.Frame(self.resumo_inner)
            af.grid(row=row, column=col, padx=8, pady=4, sticky=tk.W)

            # Checkmark ou X
            if erros == 0:
                marca = "OK"
                cor = "#228B22"
            else:
                marca = "ERRO"
                cor = "#CC0000"

            ttk.Label(af, text=marca, font=("Arial", 9, "bold"), foreground=cor, width=5).pack(side=tk.LEFT)
            ttk.Label(af, text=f"{aba}", font=("Arial", 9, "bold"), width=12, anchor=tk.W).pack(side=tk.LEFT)
            ttk.Label(af, text=f"{sucesso}/{total}", font=("Arial", 9), width=10).pack(side=tk.LEFT)

            col += 1
            if col >= 2:
                col = 0
                row += 1

        # Linha de total
        row += 1
        sep = ttk.Separator(self.resumo_inner, orient=tk.HORIZONTAL)
        sep.grid(row=row, column=0, columnspan=2, sticky=tk.EW, pady=8, padx=5)

        row += 1
        tf = ttk.Frame(self.resumo_inner)
        tf.grid(row=row, column=0, columnspan=2, pady=5)

        cor_total = "#228B22" if total_erros == 0 else "#CC0000"
        ttk.Label(tf, text=f"TOTAL: {total_ok}/{total_ok + total_erros} importados",
                 font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        if total_erros > 0:
            ttk.Label(tf, text=f" ({total_erros} erros)",
                     font=("Arial", 10, "bold"), foreground="#CC0000").pack(side=tk.LEFT)

    # ----------------------------------------------------------
    # Helpers UI
    # ----------------------------------------------------------

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

    def _detectar_driver_odbc(self):
        """Tenta cada driver ODBC ate encontrar um que funcione."""
        for driver in ODBC_DRIVERS:
            conn_str = (
                f"DRIVER={{{driver}}};"
                f"SERVER={self.servidor.get()};"
                f"DATABASE={self.banco.get()};"
                f"UID={self.usuario.get()};"
                f"PWD={self.senha.get()};"
                f"TrustServerCertificate=yes;"
            )
            try:
                conn = pyodbc.connect(conn_str, timeout=5)
                conn.close()
                return driver
            except Exception:
                continue
        return None

    def _conn_str(self, driver=None):
        """Retorna connection string usando o driver especificado ou o detectado."""
        drv = driver or self.odbc_driver or "ODBC Driver 17 for SQL Server"
        return (
            f"DRIVER={{{drv}}};"
            f"SERVER={self.servidor.get()};"
            f"DATABASE={self.banco.get()};"
            f"UID={self.usuario.get()};"
            f"PWD={self.senha.get()};"
            f"TrustServerCertificate=yes;"
        )

    def _testar_conexao(self):
        """Tenta varios drivers ODBC e descobre o nome real do servidor e do banco."""
        self.status_var.set("Detectando driver ODBC...")
        self.root.update_idletasks()

        driver = self._detectar_driver_odbc()

        if driver is None:
            drivers_testados = "\n".join(f"  - {d}" for d in ODBC_DRIVERS)
            messagebox.showerror(
                "Falha na Conexao",
                f"Nao foi possivel conectar com nenhum driver ODBC.\n\n"
                f"Drivers testados:\n{drivers_testados}\n\n"
                f"Verifique se o SQL Server esta acessivel."
            )
            self.status_var.set("Falha na conexao")
            return

        self.odbc_driver = driver

        try:
            conn_str = self._conn_str()
            conn = pyodbc.connect(conn_str, timeout=5)
            cursor = conn.cursor()
            cursor.execute("SELECT @@SERVERNAME, DB_NAME()")
            row = cursor.fetchone()

            server_real = row[0]
            db_real = row[1]
            conn.close()

            msg = (
                f"CONEXAO OK!\n\n"
                f"Driver: {driver}\n"
                f"Servidor: {server_real}\n"
                f"Banco: {db_real}"
            )
            messagebox.showinfo("Conexao", msg)
            self.status_var.set(f"Conectado via {driver}")

        except Exception as e:
            messagebox.showerror("Falha na Conexao", f"Erro: {str(e)}")
            self.status_var.set("Falha na conexao")

    def _log_detalhe(self, msg):
        """Adiciona mensagem na aba Detalhes."""
        self.resultado_detalhes.append(msg)
        def _append():
            self.text_detalhes.configure(state=tk.NORMAL)
            self.text_detalhes.insert(tk.END, msg + "\n")
            self.text_detalhes.see(tk.END)
            self.text_detalhes.configure(state=tk.DISABLED)
        self.root.after(0, _append)

    def _log_erro(self, aba, linha, pk, mensagem):
        """Adiciona erro formatado na aba Erros."""
        erro_info = {"aba": aba, "linha": linha, "pk": pk, "mensagem": mensagem}
        self.resultado_erros.append(erro_info)

        def _append():
            self.text_erros.configure(state=tk.NORMAL)
            self.text_erros.insert(tk.END, f"{aba} - Linha {linha}\n")
            self.text_erros.insert(tk.END, f"  PK: {pk}\n")
            self.text_erros.insert(tk.END, f"  {mensagem}\n\n")
            self.text_erros.see(tk.END)
            self.text_erros.configure(state=tk.DISABLED)

            # Atualiza contador na aba
            self.notebook.tab(2, text=f"Erros ({len(self.resultado_erros)})")
        self.root.after(0, _append)

    def _log(self, msg):
        """Callback de log - roteia para detalhes e detecta erros."""
        # Log para detalhes
        self._log_detalhe(msg)

        # Detecta erros e extrai informacoes
        if "ERRO" in msg and "|" in msg:
            # Formato: ERRO ABA | DB=xxx | linha N PK=[xxx]: mensagem
            try:
                partes = msg.split("|")
                aba = partes[0].replace("ERRO", "").strip()

                # Extrair linha e PK
                for p in partes:
                    if "linha" in p:
                        import re
                        match_linha = re.search(r'linha\s+(\d+)', p)
                        match_pk = re.search(r'PK=\[(.*?)\]', p)
                        linha = match_linha.group(1) if match_linha else "?"
                        pk = match_pk.group(1) if match_pk else "?"

                # Mensagem e o que vem depois do ultimo :
                if ":" in msg:
                    mensagem = msg.split(":")[-1].strip()
                else:
                    mensagem = msg

                self._log_erro(aba, linha, pk, mensagem)
            except Exception:
                # Se nao conseguir parsear, adiciona como erro generico
                self._log_erro("?", "?", "?", msg)

    def _progresso(self, value, message):
        def _update():
            self.progress_var.set(value)
            self.status_var.set(message[:50])  # Trunca mensagem longa
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

    def _limpar_resultados(self):
        """Limpa os resultados anteriores."""
        self.resultado_resumo = {}
        self.resultado_detalhes = []
        self.resultado_erros = []

        # Limpar textos
        self.text_detalhes.configure(state=tk.NORMAL)
        self.text_detalhes.delete("1.0", tk.END)
        self.text_detalhes.configure(state=tk.DISABLED)

        self.text_erros.configure(state=tk.NORMAL)
        self.text_erros.delete("1.0", tk.END)
        self.text_erros.configure(state=tk.DISABLED)

        # Reset contador erros
        self.notebook.tab(2, text="Erros (0)")

        # Limpar resumo
        self._atualizar_resumo()

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

        # Se o driver ODBC ainda nao foi detectado, detecta agora
        if self.odbc_driver is None:
            self.status_var.set("Detectando driver ODBC...")
            self.root.update_idletasks()
            driver = self._detectar_driver_odbc()
            if driver is None:
                drivers_testados = "\n".join(f"  - {d}" for d in ODBC_DRIVERS)
                messagebox.showerror(
                    "Falha na Conexao",
                    f"Nao foi possivel conectar com nenhum driver ODBC.\n\n"
                    f"Drivers testados:\n{drivers_testados}\n\n"
                    f"Verifique a conexao antes de importar."
                )
                self.status_var.set("Falha na conexao")
                return
            self.odbc_driver = driver

        # Limpar resultados anteriores
        self._limpar_resultados()
        self._log_detalhe(f"Driver ODBC: {self.odbc_driver}")

        self._set_widgets(self.root, "disabled")
        self.progress_var.set(0)

        threading.Thread(target=self._executar, args=(arquivo, abas), daemon=True).start()

    def _executar(self, arquivo, abas):
        try:
            # Pre-validacao
            if self.validar_antes.get():
                self._progresso(0, "Validando planilha...")
                self._log_detalhe("=== PRE-VALIDACAO ===")

                validator = PlanilhaValidator(arquivo, progress_callback=self._progresso)
                validator.versao_srppwin = self.versao_srppwin.get()
                excel_data, nome_arquivo, status, resultados = validator.processar("Validacao Local")

                if status == "reprovado":
                    self._log_detalhe("REPROVADO - Importacao bloqueada!")
                    for r in resultados:
                        if r.get("erros", 0) > 0:
                            self._log_detalhe(f"  {r['Planilha']}: {r['erros']} erros")

                    # Salvar planilha com resultados visuais
                    pasta = os.path.dirname(arquivo)
                    output_path = os.path.join(pasta, nome_arquivo)
                    with open(output_path, "wb") as f:
                        f.write(excel_data.getbuffer())
                    self._log_detalhe(f"  Planilha salva: {output_path}")

                    self._progresso(0, "Bloqueado - planilha reprovada")
                    msg_erro = f"A planilha tem erros e a importacao foi bloqueada.\n\nPlanilha salva em:\n{output_path}"
                    self.root.after(0, lambda: messagebox.showerror(
                        "Validacao Reprovada", msg_erro))
                    return

                if status == "advertencias":
                    self._log_detalhe("APROVADO COM ADVERTENCIAS")
                    resposta = [None]
                    def _ask():
                        resposta[0] = messagebox.askyesno(
                            "Advertencias",
                            "Planilha aprovada com advertencias.\nContinuar importacao?")
                    self.root.after(0, _ask)
                    while resposta[0] is None:
                        time.sleep(0.1)
                    if not resposta[0]:
                        self._log_detalhe("Cancelado pelo usuario.")
                        self._progresso(0, "Cancelado")
                        return
                else:
                    self._log_detalhe("APROVADO")

            # Importar
            self._log_detalhe("")
            self._log_detalhe("=== IMPORTACAO ===")
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
                limpar_auxiliares=False
            )

            # Guardar resultados para o resumo
            self.resultado_resumo = resultados

            # Log resumo
            self._log_detalhe("")
            self._log_detalhe("=== CONCLUIDO ===")
            total_s, total_e = 0, 0
            for aba, dados in resultados.items():
                if aba == "ERRO_GERAL":
                    self._log_detalhe(f"ERRO GERAL: {dados}")
                    continue
                s, e = dados["sucesso"], dados["erros"]
                total_s += s
                total_e += e

            self._log_detalhe(f"Total: {total_s} importados, {total_e} erros")

            # Atualizar resumo visual
            self.root.after(0, self._atualizar_resumo)

            if total_e == 0:
                self.root.after(0, lambda: messagebox.showinfo(
                    "Concluido", f"Importacao concluida!\n{total_s} registros importados."))
            else:
                # Muda para aba de erros
                self.root.after(0, lambda: self.notebook.select(2))
                self.root.after(0, lambda: messagebox.showwarning(
                    "Concluido com Erros",
                    f"{total_s} importados, {total_e} erros.\nVeja a aba 'Erros'."))

        except Exception as e:
            self._log_detalhe(f"ERRO FATAL: {e}")
            import traceback
            traceback.print_exc()
            self.root.after(0, lambda: messagebox.showerror("Erro", str(e)))
        finally:
            self.root.after(0, lambda: self._set_widgets(self.root, "normal"))


if __name__ == "__main__":
    root = tk.Tk()
    app = ImportadorApp(root)
    root.mainloop()
