import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

class ExcelProcessor:
    @staticmethod
    def process_excel(file_path, output_path):
        """
        Processa o arquivo Excel, adicionando e preenchendo colunas específicas.
        """
        # Ler a planilha Excel
        df = pd.read_excel(file_path)

        # Inserir nova coluna na posição D (índice 3)
        df.insert(3, 'Nova Coluna D', None)

        # Preencher a coluna D com a função INT aplicada à coluna C
        df['Nova Coluna D'] = df['C'].apply(lambda x: int(x) if pd.notnull(x) else None)

        # Substituir a coluna D pelos valores calculados
        df['Nova Coluna D'] = df['Nova Coluna D'].astype(object)

        # Inserir nova coluna na posição AB (índice 27)
        df.insert(27, 'Nova Coluna AB', None)

        # Preencher a coluna AB com a concatenação das colunas AC e AD
        df['Nova Coluna AB'] = df.apply(
            lambda row: f"{row['AC']} {row['AD']}" if pd.notnull(row['AC']) and pd.notnull(row['AD']) else None,
            axis=1
        )

        # Substituir a coluna AB pelos valores calculados
        df['Nova Coluna AB'] = df['Nova Coluna AB'].astype(object)

        # Salvar a planilha modificada
        df.to_excel(output_path, index=False)

class GUI:
    def __init__(self, root):
        """
        Inicializa a interface gráfica.
        """
        self.root = root
        self.root.title("Processador de Excel")

        # Criar widgets
        self.label_file_path = tk.Label(root, text="Arquivo de entrada:")
        self.label_file_path.pack(pady=10)

        self.entry_file_path = tk.Entry(root, width=50)
        self.entry_file_path.pack(pady=5)

        self.button_browse = tk.Button(root, text="Procurar", command=self.select_file)
        self.button_browse.pack(pady=5)

        self.button_process = tk.Button(root, text="Executar Processo", command=self.run_process)
        self.button_process.pack(pady=20)

    def select_file(self):
        """
        Abre o diálogo para selecionar um arquivo e insere o caminho no campo de entrada.
        """
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.entry_file_path.delete(0, tk.END)  # Limpa a entrada
            self.entry_file_path.insert(0, file_path)  # Insere o caminho do arquivo selecionado

    def run_process(self):
        """
        Executa o processamento do arquivo Excel.
        """
        input_file = self.entry_file_path.get()
        output_file = input_file.replace('.xlsx', '_modificado.xlsx').replace('.xls', '_modificado.xls')

        try:
            ExcelProcessor.process_excel(input_file, output_file)
            messagebox.showinfo("Sucesso", f"Processo concluído! Arquivo salvo como: {output_file}")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = GUI(root)
    root.mainloop()