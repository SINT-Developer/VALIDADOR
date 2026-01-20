import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import threading
from datetime import datetime
from pathlib import Path
from io import BytesIO
import shutil

# Importar o validador
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from planilha_validator import PlanilhaValidator


class ValidadorApp:
    def __init__(self, root, dev_mode=False):
        self.root = root
        self.dev_mode = dev_mode
        self.root.title("Validador de Planilhas - SINT" + (" [DEV]" if dev_mode else ""))
        self.root.geometry("600x350")
        self.root.resizable(False, False)
        self.setup_ui()

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

        # Botão de validação
        validate_button = ttk.Button(
            main_frame, text="Validar Planilha", command=self.start_validation
        )
        validate_button.pack(pady=20)

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
                if isinstance(child, (ttk.Button, ttk.Entry, tk.Button, tk.Entry)):
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

            # Processar a validação (progresso é reportado automaticamente pelo validador)
            t0 = time.perf_counter()
            excel_data, nome_arquivo, status, resultados = validator.processar(
                "Validação Local"
            )
            tempo_total = time.perf_counter() - t0

            # Verificar se deve gerar planilha de etiquetas
            etiquetas_result = validator.gerar_planilha_etiquetas()

            # Salvar arquivos na pasta do executável
            desktop_path = os.path.dirname(os.path.abspath(sys.argv[0]))

            # Salvar o arquivo principal
            output_path = os.path.join(desktop_path, nome_arquivo)
            with open(output_path, "wb") as f:
                f.write(excel_data.getbuffer())

            # Salvar o arquivo de etiquetas, se existir
            etiquetas_path = None
            if etiquetas_result:
                etiquetas_data, etiquetas_nome = etiquetas_result
                etiquetas_path = os.path.join(desktop_path, etiquetas_nome)
                with open(etiquetas_path, "wb") as f:
                    f.write(etiquetas_data.getbuffer())

            # Mostrar resultado
            status_text = {
                "aprovado": "APROVADO",
                "advertencias": "APROVADO COM ADVERTÊNCIAS",
                "reprovado": "REPROVADO",
            }.get(status, "Validação completa")

            message = f"Validação concluída com status: {status_text}\n\n"
            message += f"Arquivo salvo em:\n{output_path}"

            if etiquetas_path:
                message += f"\n\nArquivo de etiquetas salvo em:\n{etiquetas_path}"

            # Reabilitar interface usando a mesma abordagem recursiva
            def enable_widgets(parent):
                for child in parent.winfo_children():
                    if isinstance(child, (ttk.Button, ttk.Entry, tk.Button, tk.Entry)):
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
                    if isinstance(child, (ttk.Button, ttk.Entry, tk.Button, tk.Entry)):
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
