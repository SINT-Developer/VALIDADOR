import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import threading
from datetime import datetime
from pathlib import Path
from io import BytesIO
import shutil
import urllib.request
import json
import tempfile
import subprocess

# Versao do aplicativo
APP_VERSION = "1.0.15"
VERSION_URL = "https://gist.githubusercontent.com/SINT-Developer/a38baad856a6149526948d7c0c360ab9/raw/version.json"

# Importar o validador
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from planilha_validator import PlanilhaValidator


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
        import ssl
        import time
        # Criar contexto SSL que ignora verificacao (para evitar problemas de certificado)
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE

        # Adicionar timestamp para evitar cache do GitHub
        url_com_cache_bust = f"{VERSION_URL}?t={int(time.time())}"
        req = urllib.request.Request(url_com_cache_bust, headers={'User-Agent': 'Mozilla/5.0', 'Cache-Control': 'no-cache'})
        with urllib.request.urlopen(req, timeout=10, context=ctx) as response:
            data = json.loads(response.read().decode('utf-8'))
            versao_remota = data.get("version", "")
            download_url = data.get("download_url", "")

            if versao_remota and comparar_versoes(versao_remota, APP_VERSION) > 0:
                return versao_remota, download_url
    except Exception as e:
        # Log do erro para debug (so aparece se rodar com console)
        print(f"Erro ao verificar atualizacao: {e}")
    return None, None


def baixar_atualizacao(download_url, callback_progresso=None):
    """Baixa o novo executavel. Retorna o caminho do arquivo baixado ou None."""
    try:
        temp_dir = tempfile.gettempdir()
        temp_file = os.path.join(temp_dir, "Validador_SINT_update.exe")

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

        # Se estiver rodando como script Python, nao atualiza
        if not getattr(sys, 'frozen', False):
            print("Modo desenvolvimento - atualizacao simulada")
            return False

        # Criar script batch para atualizar
        batch_path = os.path.join(tempfile.gettempdir(), "update_validador.bat")

        # Obter nome do processo para poder verificar se fechou
        exe_nome = os.path.basename(exe_atual)

        batch_content = f'''@echo off
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
    echo Erro ao copiar arquivo. Tentando novamente...
    timeout /t 3 /nobreak >nul
    copy /Y "{novo_exe_path}" "{exe_atual}"
)

if errorlevel 1 (
    echo ERRO: Nao foi possivel atualizar o aplicativo.
    pause
) else (
    echo Atualizacao concluida com sucesso!
)

del "{novo_exe_path}" 2>nul
del "%~f0"
'''

        with open(batch_path, 'w') as f:
            f.write(batch_content)

        # Executar o batch e fechar o app
        subprocess.Popen(['cmd', '/c', batch_path],
                        creationflags=subprocess.CREATE_NO_WINDOW)
        return True
    except Exception as e:
        print(f"Erro ao aplicar atualizacao: {e}")
        return False


class ValidadorApp:
    def __init__(self, root, dev_mode=False):
        self.root = root
        self.dev_mode = dev_mode
        self.root.title(f"Validador de Planilhas - SINT v{APP_VERSION}" + (" [DEV]" if dev_mode else ""))
        self.root.geometry("600x420")
        self.root.resizable(False, False)
        self.setup_ui()

        # Verificar atualizacao em background apos iniciar
        threading.Thread(target=self._verificar_atualizacao_background, daemon=True).start()

    def _verificar_atualizacao_background(self):
        """Verifica atualizacao em background e mostra dialogo se houver."""
        nova_versao, download_url = verificar_atualizacao()
        if nova_versao and download_url:
            # Mostrar dialogo na thread principal
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
        # Criar janela de progresso
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
                    messagebox.showinfo(
                        "Atualização Concluída",
                        "A atualização foi aplicada com sucesso!\n\n"
                        "O aplicativo será fechado. Por favor, abra-o novamente para usar a nova versão."
                    )
                    self.root.destroy()  # Fecha o app para o batch substituir
                else:
                    messagebox.showerror("Erro", "Não foi possível aplicar a atualização.")
            else:
                progress_window.destroy()
                messagebox.showerror("Erro", "Falha ao baixar a atualização.")

        threading.Thread(target=fazer_download, daemon=True).start()

    def setup_ui(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Título
        title_label = ttk.Label(
            main_frame, text="Validador de Planilhas", font=("Arial", 16, "bold")
        )
        title_label.pack(pady=10)

        # Frame para seleção do arquivo
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=10)

        self.file_path = tk.StringVar()
        file_label = ttk.Label(file_frame, text="Planilha:")
        file_label.pack(side=tk.LEFT)

        file_entry = ttk.Entry(file_frame, textvariable=self.file_path, width=50)
        file_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        browse_button = ttk.Button(
            file_frame, text="Procurar", command=self.browse_file
        )
        browse_button.pack(side=tk.LEFT, padx=5)

        # Frame para opções de salvamento
        options_frame = ttk.Frame(main_frame)
        options_frame.pack(fill=tk.X, pady=5)

        self.criar_novo_arquivo = tk.BooleanVar(value=True)
        checkbox_novo_arquivo = ttk.Checkbutton(
            options_frame,
            text="Gerar novo arquivo (desmarcado = sobrescrever original)",
            variable=self.criar_novo_arquivo
        )
        checkbox_novo_arquivo.pack(anchor=tk.W)

        # Frame para versão SRPPWIN
        versao_frame = ttk.Frame(main_frame)
        versao_frame.pack(fill=tk.X, pady=5)

        versao_label = ttk.Label(versao_frame, text="Versão SRPPWIN:")
        versao_label.pack(side=tk.LEFT)

        self.versao_srppwin = tk.StringVar(value="19.1.5")
        radio_v19 = ttk.Radiobutton(
            versao_frame,
            text="19.1.5",
            variable=self.versao_srppwin,
            value="19.1.5"
        )
        radio_v19.pack(side=tk.LEFT, padx=(10, 5))

        radio_v20 = ttk.Radiobutton(
            versao_frame,
            text="20.1.0",
            variable=self.versao_srppwin,
            value="20.1.0"
        )
        radio_v20.pack(side=tk.LEFT, padx=5)

        # Botão de validação
        validate_button = ttk.Button(
            main_frame, text="Validar Planilha", command=self.start_validation
        )
        validate_button.pack(pady=15)

        # Barra de progresso
        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(
            main_frame, variable=self.progress_var, maximum=100
        )
        self.progress.pack(fill=tk.X, pady=10)

        # Status
        self.status_var = tk.StringVar()
        self.status_var.set("Selecione uma planilha para iniciar a validação")
        status_label = ttk.Label(main_frame, textvariable=self.status_var)
        status_label.pack(pady=5)

        # Créditos
        credits_label = ttk.Label(
            main_frame, text="Desenvolvido por SINT © 2025", font=("Arial", 8)
        )
        credits_label.pack(side=tk.BOTTOM, pady=10)

    def browse_file(self):
        file_types = [("Arquivos Excel", "*.xlsx *.xls *.xlsm *.xlsb")]
        file_path = filedialog.askopenfilename(filetypes=file_types)
        if file_path:
            self.file_path.set(file_path)
            self.status_var.set("Planilha selecionada: " + os.path.basename(file_path))

    def start_validation(self):
        file_path = self.file_path.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("Erro", "Selecione uma planilha válida!")
            return

        # Encontrar e desabilitar widgets interativos usando uma função recursiva
        def disable_widgets(parent):
            for child in parent.winfo_children():
                if isinstance(child, (ttk.Button, ttk.Entry, ttk.Checkbutton, ttk.Radiobutton, tk.Button, tk.Entry, tk.Checkbutton, tk.Radiobutton)):
                    try:
                        child.configure(state="disabled")
                    except:
                        pass  # Ignore se o widget não suportar state
                disable_widgets(child)  # Recursão para widgets filhos

        # Desabilitar os controles interativos
        disable_widgets(self.root)

        self.status_var.set("Processando... Por favor, aguarde.")
        self.progress_var.set(10)

        # Iniciar validação em uma thread separada para não travar a interface
        threading.Thread(
            target=self.process_validation, args=(file_path,), daemon=True
        ).start()

    def process_validation(self, file_path):
        try:
            import time

            # Atualizar progresso
            self.update_progress(2, "Carregando planilha...")

            # Medir tempo de carregamento (modo dev)
            t0 = time.perf_counter()

            # Instanciar o validador com callback de progresso
            validator = PlanilhaValidator(file_path, progress_callback=self.update_progress, dev_mode=self.dev_mode)

            tempo_load = time.perf_counter() - t0

            # Habilitar profiling se em modo dev
            if self.dev_mode:
                validator._dev_mode = True
                validator._timings = {}

            # Definir versão SRPPWIN escolhida pelo usuário
            versao_srppwin = self.versao_srppwin.get()
            validator.versao_srppwin = versao_srppwin

            # Processar a validação (progresso é reportado automaticamente pelo validador)
            t0 = time.perf_counter()
            excel_data, nome_arquivo, status, resultados = validator.processar(
                "Validação Local"
            )
            tempo_total = time.perf_counter() - t0

            # Verificar se deve gerar planilha de etiquetas
            etiquetas_result = validator.gerar_planilha_etiquetas()

            # Determinar onde salvar com base na opção do usuário
            criar_novo = self.criar_novo_arquivo.get()

            # Sempre salvar na mesma pasta da planilha original
            pasta_planilha = os.path.dirname(file_path)

            if criar_novo:
                # Salvar novo arquivo na pasta da planilha original
                output_path = os.path.join(pasta_planilha, nome_arquivo)
            else:
                # Sobrescrever o arquivo original
                output_path = file_path

            # Salvar o arquivo principal
            with open(output_path, "wb") as f:
                f.write(excel_data.getbuffer())

            # Salvar o arquivo de etiquetas, se existir (sempre como novo arquivo)
            etiquetas_path = None
            if etiquetas_result:
                etiquetas_data, etiquetas_nome = etiquetas_result
                # Etiquetas sempre na mesma pasta do arquivo de saída
                etiquetas_dir = os.path.dirname(output_path)
                etiquetas_path = os.path.join(etiquetas_dir, etiquetas_nome)
                with open(etiquetas_path, "wb") as f:
                    f.write(etiquetas_data.getbuffer())

            # Mostrar resultado
            status_text = {
                "aprovado": "APROVADO",
                "advertencias": "APROVADO COM ADVERTÊNCIAS",
                "reprovado": "REPROVADO",
            }.get(status, "Validação completa")

            message = f"Validação concluída com status: {status_text}\n\n"
            if criar_novo:
                message += f"Novo arquivo salvo em:\n{output_path}"
            else:
                message += f"Arquivo original atualizado:\n{output_path}"

            if etiquetas_path:
                message += f"\n\nArquivo de etiquetas salvo em:\n{etiquetas_path}"

            # Reabilitar interface usando a mesma abordagem recursiva
            def enable_widgets(parent):
                for child in parent.winfo_children():
                    if isinstance(child, (ttk.Button, ttk.Entry, ttk.Checkbutton, ttk.Radiobutton, tk.Button, tk.Entry, tk.Checkbutton, tk.Radiobutton)):
                        try:
                            child.configure(state="normal")
                        except:
                            pass  # Ignore se o widget não suportar state
                    enable_widgets(child)  # Recursão para widgets filhos

            # Habilitar os controles interativos
            enable_widgets(self.root)

            # Exibir mensagem
            messagebox.showinfo("Validação Concluída", message)

            # Atualizar status
            self.status_var.set(f"Validação concluída com status: {status_text}")

            # Gerar relatório de profiling se em modo dev
            if self.dev_mode:
                self._gerar_relatorio_dev(validator, tempo_load, tempo_total, resultados, status)

        except Exception as e:
            messagebox.showerror(
                "Erro", f"Ocorreu um erro durante a validação:\n{str(e)}"
            )
            self.status_var.set("Erro durante a validação.")

            # Reabilitar interface usando a mesma abordagem recursiva
            def enable_widgets(parent):
                for child in parent.winfo_children():
                    if isinstance(child, (ttk.Button, ttk.Entry, ttk.Checkbutton, ttk.Radiobutton, tk.Button, tk.Entry, tk.Checkbutton, tk.Radiobutton)):
                        try:
                            child.configure(state="normal")
                        except:
                            pass  # Ignore se o widget não suportar state
                    enable_widgets(child)  # Recursão para widgets filhos

            # Habilitar os controles interativos
            enable_widgets(self.root)

    def update_progress(self, value, message):
        # Atualizar a UI a partir de uma thread
        self.root.after(0, lambda: self.progress_var.set(value))
        self.root.after(0, lambda: self.status_var.set(message))
        self.root.update_idletasks()

    def _gerar_relatorio_dev(self, validator, tempo_load, tempo_total, resultados, status):
        """Gera relatório de profiling no console."""
        print("\n" + "=" * 60)
        print("RELATORIO DE PERFORMANCE")
        print("=" * 60)
        print(f"{'Etapa':<35} {'Tempo':>10} {'%':>8}")
        print("-" * 60)

        timings = getattr(validator, '_timings', {})
        for etapa, tempo in sorted(timings.items(), key=lambda x: -x[1]):
            pct = (tempo / tempo_total * 100) if tempo_total > 0 else 0
            print(f"{etapa:<35} {tempo:>9.2f}s {pct:>7.1f}%")

        print("-" * 60)
        print(f"{'Carregamento workbook':<35} {tempo_load:>9.2f}s")
        print(f"{'TOTAL PROCESSAMENTO':<35} {tempo_total:>9.2f}s")
        print("=" * 60)

        print("\nLINHAS POR ABA:")
        print("-" * 40)
        for r in resultados:
            print(f"  {r['Planilha']:<20} {r.get('lidas', 0):>6} linhas")

        print("\n" + "=" * 60)
        print(f"Status final: {status.upper()}")
        print("=" * 60)


if __name__ == "__main__":
    dev_mode = "--dev" in sys.argv
    root = tk.Tk()
    app = ValidadorApp(root, dev_mode=dev_mode)
    root.mainloop()
