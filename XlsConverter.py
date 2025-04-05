import os
import pandas as pd

class ExcelConverter:
    @staticmethod
    def convert_xls_to_xlsx(input_file, output_file):
        """
        Converte um arquivo Excel .xls para .xlsx sem perder o conteúdo.
        O arquivo original permanece intacto.
        
        :param input_file: Caminho do arquivo .xls de entrada.
        :param output_file: Caminho do arquivo .xlsx de saída.
        """
        # Verificar se o arquivo de entrada tem a extensão correta
        if not input_file.lower().endswith('.xls'):
            raise ValueError("O arquivo de entrada deve ter a extensão .xls")

        # Garantir que o arquivo de saída não sobrescreva o arquivo de entrada
        if os.path.abspath(input_file) == os.path.abspath(output_file):
            raise ValueError("O arquivo de saída não pode ter o mesmo nome ou caminho do arquivo de entrada.")

        # Ler o arquivo .xls usando pandas
        try:
            df = pd.read_excel(input_file, engine='xlrd')
        except Exception as e:
            raise ValueError(f"Erro ao ler o arquivo .xls: {e}")

        # Salvar como .xlsx usando openpyxl
        try:
            df.to_excel(output_file, index=False, engine='openpyxl')
            print(f"Arquivo convertido com sucesso: {output_file}")
        except Exception as e:
            raise ValueError(f"Erro ao salvar o arquivo .xlsx: {e}")

# Exemplo de uso
if __name__ == "__main__":
    input_path = "caminho_para_seu_arquivo.xls"  # Substitua pelo caminho do arquivo .xls
    output_path = "caminho_para_seu_arquivo_convertido.xlsx"  # Substitua pelo caminho do arquivo .xlsx
    ExcelConverter.convert_xls_to_xlsx(input_path, output_path)