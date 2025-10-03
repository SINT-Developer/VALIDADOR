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
    def __init__(self, root):
        self.root = root
        self.root.title("Validador de Planilhas - SINT")
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
            # Atualizar progresso
            self.update_progress(20, "Carregando planilha...")

            # Instanciar o validador
            validator = PlanilhaValidator(file_path)

            # Processar a validação
            self.update_progress(30, "Validando dados...")
            excel_data, nome_arquivo, status, resultados = validator.processar(
                "Validação Local"
            )

            # Verificar se deve gerar planilha de etiquetas
            self.update_progress(70, "Gerando relatórios...")
            etiquetas_result = validator.gerar_planilha_etiquetas()

            # Salvar arquivos na pasta do executável
            self.update_progress(80, "Salvando arquivos...")
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

            # Finalizar
            self.update_progress(100, "Validação concluída!")

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


if __name__ == "__main__":
    root = tk.Tk()
    app = ValidadorApp(root)
    root.mainloop()
