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
import urllib.request
import json
import tempfile
import subprocess
import ssl

# Paths
if getattr(sys, "frozen", False):
    # Rodando no EXE (PyInstaller)
    script_dir = sys._MEIPASS
else:
    # Rodando em Python normal
    script_dir = os.path.dirname(os.path.abspath(__file__))

parent_dir = os.path.dirname(script_dir)

if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)


from motor import PlanilhaImportador
from mapeamento import ORDEM_IMPORTACAO
from planilha_validator import PlanilhaValidator

APP_VERSION = "1.2.29"
VERSION_URL = "https://gist.githubusercontent.com/SINT-Developer/4005d983a51756ce108aef5f064f6c01/raw/importador_version.json"


def comparar_versoes(v1, v2):
    """Compara duas versoes. Retorna 1 se v1 > v2, -1 se v1 < v2, 0 se iguais."""
    def parse(v):
        return [int(x) for x in v.replace("v", "").split(".")]
    p1, p2 = parse(v1), parse(v2)
    for a, b in zip(p1, p2):
        if a > b:
            return 1
        if a < b:
            return -1
    return 0


def verificar_atualizacao():
    """Verifica se ha uma nova versao disponivel. Retorna (nova_versao, download_url) ou (None, None)."""
    try:
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE

        url_com_cache_bust = f"{VERSION_URL}?t={int(time.time())}"
        req = urllib.request.Request(url_com_cache_bust, headers={'User-Agent': 'Mozilla/5.0', 'Cache-Control': 'no-cache'})
        with urllib.request.urlopen(req, timeout=10, context=ctx) as response:
            data = json.loads(response.read().decode('utf-8'))
            versao_remota = data.get("version", "")
            download_url = data.get("download_url", "")

            if versao_remota and comparar_versoes(versao_remota, APP_VERSION) > 0:
                return versao_remota, download_url
    except Exception as e:
        print(f"Erro ao verificar atualizacao: {e}")
    return None, None


def baixar_atualizacao(download_url, callback_progresso=None):
    """Baixa o novo executavel. Retorna o caminho do arquivo baixado ou None."""
    try:
        temp_dir = tempfile.gettempdir()
        temp_file = os.path.join(temp_dir, "Importador_SINT_update.exe")

        def report_progress(block_num, block_size, total_size):
            if callback_progresso and total_size > 0:
                percent = int(block_num * block_size * 100 / total_size)
                callback_progresso(min(percent, 100))

        urllib.request.urlretrieve(download_url, temp_file, report_progress)
        return temp_file
    except Exception as e:
        print(f"Erro ao baixar atualizacao: {e}")
        return None


def aplicar_atualizacao(novo_exe_path):
    """Cria script batch para substituir o exe e reiniciar o app."""
    try:
        exe_atual = sys.executable

        if not getattr(sys, 'frozen', False):
            print("Modo desenvolvimento - atualizacao simulada")
            return False

        batch_path = os.path.join(tempfile.gettempdir(), "update_importador.bat")
        exe_nome = os.path.basename(exe_atual)

        batch_content = f'''@echo off
title Importador SINT - Atualizacao
echo Aguardando o aplicativo fechar...
timeout /t 3 /nobreak >nul

:check_process
tasklist /FI "IMAGENAME eq {exe_nome}" 2>NUL | find /I "{exe_nome}" >NUL
if not errorlevel 1 (
    echo Aplicativo ainda em execucao, aguardando...
    timeout /t 2 /nobreak >nul
    goto check_process
)

echo Aplicativo fechado. Aplicando atualizacao...
timeout /t 1 /nobreak >nul

copy /Y "{novo_exe_path}" "{exe_atual}"
if errorlevel 1 (
    echo Tentando novamente em 3 segundos...
    timeout /t 3 /nobreak >nul
    copy /Y "{novo_exe_path}" "{exe_atual}"
)

if errorlevel 1 (
    echo.
    echo ERRO: Nao foi possivel substituir o arquivo.
    echo Feche qualquer outra instancia do Importador e tente novamente.
    pause
    exit /b 1
)

del "{novo_exe_path}" 2>nul
del "%~f0"

echo.
echo Atualizacao concluida com sucesso! Pode abrir o Importador SINT agora.
pause
'''

        with open(batch_path, 'w', encoding='utf-8') as f:
            f.write(batch_content)

        subprocess.Popen(['cmd', '/c', batch_path])
        return True
    except Exception as e:
        print(f"Erro ao aplicar atualizacao: {e}")
        return False

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
        self.root.geometry("800x600")
        self.root.minsize(700, 500)

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

        # Opcoes de validacao (igual ao validador_standalone)
        self.criar_novo_arquivo = tk.BooleanVar(value=True)
        self.gerar_etiquetas = tk.BooleanVar(value=True)

        # Controle de exibicao da conexao SQL
        self.mostrar_conexao = tk.BooleanVar(value=False)

        self.abas_vars = {}
        for aba in ORDEM_IMPORTACAO:
            self.abas_vars[aba] = tk.BooleanVar(value=True)

        # Driver ODBC detectado (None = ainda nao detectado)
        self.odbc_driver = None

        # Dados de log (integrado na mesma janela)
        self.resultado_resumo = {}
        self.resultado_detalhes = []
        self.resultado_erros = []

        self._build_ui()

        # Verificar atualizacao em background
        threading.Thread(target=self._verificar_atualizacao_background, daemon=True).start()

    # ----------------------------------------------------------
    # Auto-Update
    # ----------------------------------------------------------

    def _verificar_atualizacao_background(self):
        """Verifica atualizacao em background e mostra dialogo se houver."""
        nova_versao, download_url = verificar_atualizacao()
        if nova_versao and download_url:
            self.root.after(0, lambda: self._mostrar_dialogo_atualizacao(nova_versao, download_url))

    def _mostrar_dialogo_atualizacao(self, nova_versao, download_url):
        """Mostra dialogo perguntando se deseja atualizar."""
        resposta = messagebox.askyesno(
            "Atualizacao Disponivel",
            f"Nova versao disponivel: v{nova_versao}\n"
            f"Versao atual: v{APP_VERSION}\n\n"
            "Deseja atualizar agora?"
        )
        if resposta:
            self._executar_atualizacao(download_url)

    def _executar_atualizacao(self, download_url):
        """Executa o processo de atualizacao."""
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Atualizando...")
        progress_window.geometry("300x100")
        progress_window.resizable(False, False)
        progress_window.transient(self.root)
        progress_window.grab_set()

        ttk.Label(progress_window, text="Baixando atualizacao...").pack(pady=10)
        progress_bar = ttk.Progressbar(progress_window, length=250, mode='determinate')
        progress_bar.pack(pady=10)

        def atualizar_progresso(percent):
            progress_bar['value'] = percent
            progress_window.update()

        def fazer_download():
            novo_exe = baixar_atualizacao(download_url, atualizar_progresso)
            if novo_exe:
                progress_window.destroy()
                if aplicar_atualizacao(novo_exe):
                    def _encerrar():
                        messagebox.showinfo(
                            "Atualizando...",
                            "O aplicativo sera fechado agora.\n\n"
                            "Uma janela de atualizacao sera exibida automaticamente — "
                            "aguarde ela concluir antes de abrir o Importador novamente."
                        )
                        os._exit(0)
                    self.root.after(0, _encerrar)
                else:
                    messagebox.showerror("Erro", "Nao foi possivel aplicar a atualizacao.")
            else:
                progress_window.destroy()
                messagebox.showerror("Erro", "Falha ao baixar a atualizacao.")

        threading.Thread(target=fazer_download, daemon=True).start()

    # ----------------------------------------------------------
    # UI
    # ----------------------------------------------------------

    def _build_ui(self):
        # Barra de status fixa na parte inferior (sempre visivel, mesmo na aba Logs)
        pf = ttk.Frame(self.root, padding=(5, 2, 5, 5))
        pf.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Progressbar(pf, variable=self.progress_var, maximum=100).pack(fill=tk.X)
        self.lbl_status = ttk.Label(pf, textvariable=self.status_var, font=("Arial", 9), anchor=tk.W)
        self.lbl_status.pack(fill=tk.X, pady=(2, 0))

        # Notebook principal (abas: Importador / Logs)
        self.notebook_principal = ttk.Notebook(self.root)
        self.notebook_principal.pack(fill=tk.BOTH, expand=True, padx=5, pady=(5, 0))

        # =======================================================
        # ABA 1: IMPORTADOR
        # =======================================================
        self.tab_importador = ttk.Frame(self.notebook_principal)
        self.notebook_principal.add(self.tab_importador, text="Importador")

        main = ttk.Frame(self.tab_importador, padding=8)
        main.pack(fill=tk.BOTH, expand=True)
        main.columnconfigure(0, weight=1)

        ttk.Label(main, text="Importador de Planilhas", font=("Arial", 12, "bold")).grid(row=0, column=0, pady=(0, 5), sticky=tk.W)

        # Arquivo
        f = ttk.LabelFrame(main, text="Planilha", padding=4)
        f.grid(row=1, column=0, sticky=tk.EW, pady=(0, 4))
        f.columnconfigure(0, weight=1)
        ttk.Entry(f, textvariable=self.file_path, state="readonly").grid(row=0, column=0, sticky=tk.EW, padx=(0, 5))
        ttk.Button(f, text="Procurar", command=self._browse).grid(row=0, column=1)

        # Conexao SQL Server - Colapsavel
        f = ttk.LabelFrame(main, text="Conexao SQL Server", padding=4)
        f.grid(row=2, column=0, sticky=tk.EW, pady=(0, 4))

        # Header clicavel para expandir/colapsar
        header_frame = ttk.Frame(f)
        header_frame.pack(fill=tk.X)

        self.btn_toggle_conexao = ttk.Checkbutton(
            header_frame, text="Mostrar configuracoes de conexao",
            variable=self.mostrar_conexao, command=self._toggle_conexao
        )
        self.btn_toggle_conexao.pack(side=tk.LEFT)

        ttk.Button(header_frame, text="Testar", command=self._testar_conexao, width=8).pack(side=tk.RIGHT)

        # Frame com os campos (inicialmente oculto)
        self.frame_conexao_campos = ttk.Frame(f)

        g = ttk.Frame(self.frame_conexao_campos)
        g.pack(fill=tk.X, pady=(5, 0))

        ttk.Label(g, text="Servidor:").grid(row=0, column=0, sticky=tk.W, padx=(0, 2))
        ttk.Entry(g, textvariable=self.servidor, width=18).grid(row=0, column=1, padx=(0, 10))
        ttk.Label(g, text="Banco:").grid(row=0, column=2, sticky=tk.W, padx=(0, 2))
        ttk.Entry(g, textvariable=self.banco, width=18).grid(row=0, column=3, padx=(0, 10))

        ttk.Label(g, text="Usuario:").grid(row=1, column=0, sticky=tk.W, padx=(0, 2), pady=(3, 0))
        ttk.Entry(g, textvariable=self.usuario, width=18).grid(row=1, column=1, padx=(0, 10), pady=(3, 0))
        ttk.Label(g, text="Senha:").grid(row=1, column=2, sticky=tk.W, padx=(0, 2), pady=(3, 0))
        ttk.Entry(g, textvariable=self.senha, width=18, show="*").grid(row=1, column=3, padx=(0, 10), pady=(3, 0))

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

        grupos_frame = ttk.Frame(f)
        grupos_frame.pack(fill=tk.X)

        for idx, grupo in enumerate(GRUPOS_ABAS):
            gf = ttk.LabelFrame(grupos_frame, text=grupo["nome"], padding=2)
            gf.grid(row=0, column=idx, padx=3, pady=2, sticky=tk.N)

            for aba in grupo["abas"]:
                if aba in self.abas_vars:
                    ttk.Checkbutton(gf, text=aba, variable=self.abas_vars[aba]).pack(anchor=tk.W)

        bf = ttk.Frame(f)
        bf.pack(fill=tk.X, pady=(4, 0))
        ttk.Button(bf, text="Todas", command=lambda: self._set_abas(True), width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(bf, text="Nenhuma", command=lambda: self._set_abas(False), width=8).pack(side=tk.LEFT, padx=2)

        # Opcoes de Validacao
        f = ttk.LabelFrame(main, text="Opcoes de Validacao", padding=4)
        f.grid(row=5, column=0, sticky=tk.EW, pady=(0, 4))

        # Linha 1: Validar antes + Versao
        g1 = ttk.Frame(f)
        g1.pack(fill=tk.X)

        ttk.Checkbutton(g1, text="Validar antes de importar", variable=self.validar_antes,
                       command=self._toggle_opcoes_validacao).pack(side=tk.LEFT)

        # Frame das opcoes de validacao (inicialmente oculto)
        self.frame_opcoes_validacao = ttk.Frame(g1)

        ttk.Label(self.frame_opcoes_validacao, text="Versao:").pack(side=tk.LEFT, padx=(15, 2))
        ttk.Combobox(self.frame_opcoes_validacao, textvariable=self.versao_srppwin,
                    values=["19.1.5", "20.1.0"], state="readonly", width=8).pack(side=tk.LEFT)

        # Linha 2: Opcoes de arquivo (aparecem junto com validar)
        self.frame_opcoes_arquivo = ttk.Frame(f)

        ttk.Checkbutton(self.frame_opcoes_arquivo, text="Gerar novo arquivo (desmarcado = sobrescrever original)",
                       variable=self.criar_novo_arquivo).pack(anchor=tk.W)
        ttk.Checkbutton(self.frame_opcoes_arquivo, text="Gerar planilha de etiquetas (se houver QtdeEtiquetas > 0)",
                       variable=self.gerar_etiquetas).pack(anchor=tk.W)

        # Opcoes Gerais
        f = ttk.LabelFrame(main, text="Opcoes Gerais", padding=4)
        f.grid(row=6, column=0, sticky=tk.EW, pady=(0, 4))
        g = ttk.Frame(f)
        g.pack(fill=tk.X)

        ttk.Checkbutton(g, text="Excluir TUDO antes da importacao (backup automatico)",
                       variable=self.excluir_tudo_var).pack(side=tk.LEFT)

        # Botao e Progresso
        bf = ttk.Frame(main, padding=8)
        bf.grid(row=7, column=0, sticky=tk.EW, pady=(8, 6))
        bf.columnconfigure(0, weight=1)

        self.btn_importar = ttk.Button(
            bf,
            text="IMPORTAR",
            command=self._iniciar
        )
        self.btn_importar.grid(row=0, column=0, sticky=tk.EW, ipady=8)

        # Creditos
        ttk.Label(main, text="Desenvolvido por SINT", font=("Arial", 8), foreground="gray").grid(
            row=8, column=0, pady=(5, 0), sticky=tk.E)

        # =======================================================
        # ABA 2: LOGS
        # =======================================================
        self.tab_logs = ttk.Frame(self.notebook_principal)
        self.notebook_principal.add(self.tab_logs, text="Logs")

        self._build_logs_tab()

    def _build_logs_tab(self):
        """Constroi a aba de Logs com 3 sub-abas (Resumo/Detalhes/Erros)."""
        main = ttk.Frame(self.tab_logs, padding=10)
        main.pack(fill=tk.BOTH, expand=True)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(0, weight=1)

        # Notebook interno (sub-abas)
        self.notebook_logs = ttk.Notebook(main)
        self.notebook_logs.grid(row=0, column=0, sticky=tk.NSEW)

        # Sub-aba Resumo
        self.frame_resumo = ttk.Frame(self.notebook_logs)
        self.notebook_logs.add(self.frame_resumo, text="Resumo")
        self._build_resumo_grid()

        # Sub-aba Detalhes
        self.frame_detalhes = ttk.Frame(self.notebook_logs)
        self.notebook_logs.add(self.frame_detalhes, text="Detalhes")
        self.frame_detalhes.columnconfigure(0, weight=1)
        self.frame_detalhes.rowconfigure(0, weight=1)
        self.text_detalhes = tk.Text(self.frame_detalhes, font=("Consolas", 10), state=tk.DISABLED,
                                      wrap=tk.WORD, bg="#1E1E1E", fg="#D4D4D4",
                                      selectbackground="#264F78", insertbackground="#D4D4D4")
        sb_det = ttk.Scrollbar(self.frame_detalhes, orient=tk.VERTICAL, command=self.text_detalhes.yview)
        self.text_detalhes.configure(yscrollcommand=sb_det.set)
        self.text_detalhes.grid(row=0, column=0, sticky=tk.NSEW)
        sb_det.grid(row=0, column=1, sticky=tk.NS)

        # Tags visuais para o log de detalhes
        self.text_detalhes.tag_configure("major_header", foreground="#CE9178", font=("Consolas", 12, "bold"),
                                          spacing1=8, spacing3=4)
        self.text_detalhes.tag_configure("header", foreground="#569CD6", font=("Consolas", 11, "bold"),
                                          spacing1=6, spacing3=2)
        self.text_detalhes.tag_configure("success", foreground="#4EC9B0", spacing1=1)
        self.text_detalhes.tag_configure("error", foreground="#F44747", font=("Consolas", 10, "bold"), spacing1=1)
        self.text_detalhes.tag_configure("result", foreground="#DCDCAA", font=("Consolas", 10, "bold"),
                                          spacing1=4, spacing3=2)
        self.text_detalhes.tag_configure("final", foreground="#4FC1FF", font=("Consolas", 11, "bold"),
                                          spacing1=8)
        self.text_detalhes.tag_configure("step", foreground="#C586C0", spacing1=4)
        self.text_detalhes.tag_configure("debug", foreground="#808080", spacing1=1)
        self.text_detalhes.tag_configure("info", foreground="#D4D4D4", spacing1=1)

        # Sub-aba Erros
        self.frame_erros = ttk.Frame(self.notebook_logs)
        self.notebook_logs.add(self.frame_erros, text="Erros (0)")
        self.frame_erros.columnconfigure(0, weight=1)
        self.frame_erros.rowconfigure(0, weight=1)
        self.text_erros = tk.Text(self.frame_erros, font=("Consolas", 10), state=tk.DISABLED, wrap=tk.WORD, fg="#CC0000")
        sb_err = ttk.Scrollbar(self.frame_erros, orient=tk.VERTICAL, command=self.text_erros.yview)
        self.text_erros.configure(yscrollcommand=sb_err.set)
        self.text_erros.grid(row=0, column=0, sticky=tk.NSEW)
        sb_err.grid(row=0, column=1, sticky=tk.NS)

        # Botao Limpar
        btn_frame = ttk.Frame(main)
        btn_frame.grid(row=1, column=0, pady=(10, 0))
        ttk.Button(btn_frame, text="Limpar Logs", command=self._limpar_logs, width=15).pack()

    def _build_resumo_grid(self):
        """Constroi o grid de resumo com checkmarks."""
        for w in self.frame_resumo.winfo_children():
            w.destroy()

        self.frame_resumo.columnconfigure(0, weight=1)
        self.frame_resumo.rowconfigure(0, weight=1)

        canvas = tk.Canvas(self.frame_resumo, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.frame_resumo, orient="vertical", command=canvas.yview)
        self.resumo_inner = ttk.Frame(canvas)

        self.resumo_inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.resumo_inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.grid(row=0, column=0, sticky=tk.NSEW)
        scrollbar.grid(row=0, column=1, sticky=tk.NS)

    def _toggle_conexao(self):
        """Mostra/esconde os campos de conexao SQL Server."""
        if self.mostrar_conexao.get():
            self.frame_conexao_campos.pack(fill=tk.X)
        else:
            self.frame_conexao_campos.pack_forget()

    def _toggle_opcoes_validacao(self):
        """Mostra/esconde as opcoes de validacao baseado no checkbox validar."""
        if self.validar_antes.get():
            self.frame_opcoes_validacao.pack(side=tk.LEFT)
            self.frame_opcoes_arquivo.pack(fill=tk.X, pady=(5, 0))
        else:
            self.frame_opcoes_validacao.pack_forget()
            self.frame_opcoes_arquivo.pack_forget()

    def _atualizar_resumo(self):
        """Atualiza o grid de resumo com os resultados."""
        for w in self.resumo_inner.winfo_children():
            w.destroy()

        if not self.resultado_resumo:
            ttk.Label(self.resumo_inner, text="Nenhuma importacao realizada ainda.",
                     font=("Arial", 10), foreground="gray").pack(pady=20)
            return

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

            af = ttk.Frame(self.resumo_inner)
            af.grid(row=row, column=col, padx=10, pady=6, sticky=tk.W)

            if erros == 0:
                marca = "OK"
                cor = "#228B22"
            else:
                marca = "ERRO"
                cor = "#CC0000"

            ttk.Label(af, text=marca, font=("Arial", 10, "bold"), foreground=cor, width=6).pack(side=tk.LEFT)
            ttk.Label(af, text=f"{aba}", font=("Arial", 10, "bold"), width=14, anchor=tk.W).pack(side=tk.LEFT)
            ttk.Label(af, text=f"{sucesso}/{total}", font=("Arial", 10), width=10).pack(side=tk.LEFT)

            col += 1
            if col >= 2:
                col = 0
                row += 1

        row += 1
        sep = ttk.Separator(self.resumo_inner, orient=tk.HORIZONTAL)
        sep.grid(row=row, column=0, columnspan=2, sticky=tk.EW, pady=10, padx=5)

        row += 1
        tf = ttk.Frame(self.resumo_inner)
        tf.grid(row=row, column=0, columnspan=2, pady=8)

        ttk.Label(tf, text=f"TOTAL: {total_ok}/{total_ok + total_erros} importados",
                 font=("Arial", 11, "bold")).pack(side=tk.LEFT)
        if total_erros > 0:
            ttk.Label(tf, text=f" ({total_erros} erros)",
                     font=("Arial", 11, "bold"), foreground="#CC0000").pack(side=tk.LEFT)

    def _limpar_logs(self):
        """Limpa todos os logs."""
        self.resultado_resumo = {}
        self.resultado_detalhes = []
        self.resultado_erros = []

        self.text_detalhes.configure(state=tk.NORMAL)
        self.text_detalhes.delete("1.0", tk.END)
        self.text_detalhes.configure(state=tk.DISABLED)

        self.text_erros.configure(state=tk.NORMAL)
        self.text_erros.delete("1.0", tk.END)
        self.text_erros.configure(state=tk.DISABLED)

        self.notebook_logs.tab(2, text="Erros (0)")
        self._atualizar_resumo()

    def _mostrar_aba_logs(self):
        """Muda para a aba de Logs e sub-aba Detalhes."""
        self.notebook_principal.select(1)
        self.notebook_logs.select(1)

    def _mostrar_erros(self):
        """Seleciona a sub-aba de erros."""
        self.notebook_logs.select(2)

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
        """Adiciona mensagem na aba Detalhes com formatacao visual."""
        self.resultado_detalhes.append(msg)

        # Determinar tag baseado no conteudo
        tag = "info"
        stripped = msg.strip()
        if msg.startswith("==="):
            tag = "major_header"
        elif msg.startswith("---"):
            tag = "header"
        elif "ERRO" in msg:
            tag = "error"
        elif "DEBUG" in msg:
            tag = "debug"
        elif stripped.startswith("Total:"):
            tag = "final"
        elif stripped.startswith(("EMPRESA:", "FILIAL:", "REPR:", "PAGTO:", "TRANSP:",
                                  "ESTADOS:", "FAMILIAS:", "ESTILOS:", "CLIENTES:",
                                  "PAGTOFILIAL:", "PRODUTOS:", "Tempo total:")):
            tag = "result"
        elif stripped.startswith(("Fazendo backup", "Limpando tabelas", "Executando ",
                                  "Conectado ao", "Planilha:", "Todos os dados")):
            tag = "step"
        elif "OK:" in msg or "ok," in msg:
            tag = "success"

        def _append():
            self.text_detalhes.configure(state=tk.NORMAL)
            self.text_detalhes.insert(tk.END, msg + "\n", tag)
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
            self.notebook_logs.tab(2, text=f"Erros ({len(self.resultado_erros)})")
        self.root.after(0, _append)

    def _log(self, msg):
        """Callback de log - roteia para detalhes e detecta erros."""
        self._log_detalhe(msg)

        if "ERRO" in msg and "|" in msg:
            try:
                import re
                partes = msg.split("|")
                aba = partes[0].replace("ERRO", "").strip()

                linha = "?"
                pk = "?"
                for p in partes:
                    if "linha" in p:
                        match_linha = re.search(r'linha\s+(\d+)', p)
                        match_pk = re.search(r'PK=\[(.*?)\]', p)
                        linha = match_linha.group(1) if match_linha else "?"
                        pk = match_pk.group(1) if match_pk else "?"

                mensagem = msg.split(":")[-1].strip() if ":" in msg else msg
                self._log_erro(aba, linha, pk, mensagem)
            except Exception:
                self._log_erro("?", "?", "?", msg)

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

        # Limpar logs e ir para aba de logs
        self._limpar_logs()
        self._mostrar_aba_logs()
        self._log_detalhe(f"Driver ODBC: {self.odbc_driver}")

        self._set_widgets(self.root, "disabled")
        self.progress_var.set(0)

        threading.Thread(target=self._executar, args=(arquivo, abas), daemon=True).start()

    def _executar(self, arquivo, abas):
        try:
            # Pre-validacao (se marcada)
            if self.validar_antes.get():
                self._progresso(2, "Carregando planilha em memoria...")
                self._log_detalhe("=== PRE-VALIDACAO ===")
                self._log_detalhe(f"Carregando: {os.path.basename(arquivo)} (aguarde...)")

                validator = PlanilhaValidator(arquivo, progress_callback=self._progresso)
                validator.versao_srppwin = self.versao_srppwin.get()
                excel_data, nome_arquivo, status, resultados = validator.processar("Validacao Local")

                # Determinar onde salvar com base na opcao do usuario
                criar_novo = self.criar_novo_arquivo.get()
                pasta_planilha = os.path.dirname(arquivo)

                if criar_novo:
                    output_path = os.path.join(pasta_planilha, nome_arquivo)
                else:
                    output_path = arquivo

                # Salvar planilha validada
                with open(output_path, "wb") as f:
                    f.write(excel_data.getbuffer())
                self._log_detalhe(f"Planilha salva: {output_path}")

                # Verificar se deve gerar planilha de etiquetas
                gerar_etiq = self.gerar_etiquetas.get()
                etiquetas_path = None
                etiquetas_erro = None

                if gerar_etiq:
                    try:
                        etiquetas_result = validator.gerar_planilha_etiquetas()
                        if etiquetas_result:
                            etiquetas_data, etiquetas_nome = etiquetas_result
                            etiquetas_path = os.path.join(pasta_planilha, etiquetas_nome)
                            with open(etiquetas_path, "wb") as f:
                                f.write(etiquetas_data.getbuffer())
                            self._log_detalhe(f"Planilha de etiquetas: {etiquetas_path}")
                    except Exception as e:
                        etiquetas_erro = str(e)
                        self._log_detalhe(f"ERRO ao gerar etiquetas: {etiquetas_erro}")

                # Status da validacao
                status_text = {
                    "aprovado": "APROVADO",
                    "advertencias": "APROVADO COM ADVERTENCIAS",
                    "reprovado": "REPROVADO",
                }.get(status, "Validacao completa")

                self._log_detalhe(f"Status: {status_text}")

                # Montar mensagem de resultado
                msg_resultado = f"Validacao concluida: {status_text}\n\n"
                if criar_novo:
                    msg_resultado += f"Arquivo salvo em:\n{output_path}"
                else:
                    msg_resultado += f"Arquivo atualizado:\n{output_path}"

                if etiquetas_path:
                    msg_resultado += f"\n\nPlanilha de etiquetas:\n{etiquetas_path}"
                elif gerar_etiq and etiquetas_erro:
                    msg_resultado += f"\n\nERRO ao gerar etiquetas:\n{etiquetas_erro}"
                elif gerar_etiq:
                    msg_resultado += "\n\nPlanilha de etiquetas NAO gerada: nenhuma linha com QtdeEtiquetas > 0."

                # Se REPROVADO, bloqueia importacao
                if status == "reprovado":
                    self._progresso(0, "Bloqueado - planilha reprovada")
                    self.root.after(0, lambda: messagebox.showerror("Validacao Reprovada", msg_resultado))
                    return

                # SEMPRE perguntar se quer importar apos validacao
                resposta = [None]
                def _ask():
                    resposta[0] = messagebox.askyesno(
                        "Validacao Concluida",
                        msg_resultado + "\n\nDeseja importar os dados agora?")
                self.root.after(0, _ask)
                while resposta[0] is None:
                    time.sleep(0.1)

                if not resposta[0]:
                    self._log_detalhe("Importacao cancelada pelo usuario.")
                    self._progresso(100, "Validacao concluida - importacao cancelada")
                    return

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
            def _update_resumo():
                self._atualizar_resumo()
                if total_e > 0:
                    self._mostrar_erros()
            self.root.after(0, _update_resumo)

            if total_e == 0:
                self.root.after(0, lambda: messagebox.showinfo(
                    "Concluido", f"Importacao concluida!\n{total_s} registros importados."))
            else:
                self.root.after(0, lambda: messagebox.showwarning(
                    "Concluido com Erros",
                    f"{total_s} importados, {total_e} erros.\nVeja a janela de Resultado."))

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