import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Inicializar a interface gráfica e ocultar a janela principal
root = Tk()
root.withdraw()

# Função para converter string percentual para número
def convert_percentage(value):
    if pd.isnull(value):
        return 0
    if isinstance(value, str):  # Apenas aplicar replace se for string
        try:
            return float(value.replace(',', '.').replace('%', '')) / 100
        except ValueError:
            return 0
    return float(value) / 100  # Para valores float

# Função para formatar o número como porcentagem
def format_as_percentage(value):
    return f"{value * 100:.2f}%"

# Abrir janela de seleção de arquivo
arquivo_selecionado = askopenfilename(title="Selecione sua planilha", filetypes=[("Arquivos Excel", "*.xlsx")])

# Ler o arquivo Excel
xls = pd.ExcelFile(arquivo_selecionado)


def convert_percentage(value):
    # Sua função de conversão aqui
    return float(value.strip('%')) / 100 if isinstance(value, str) and '%' in value else float(value)

def format_as_percentage(value):
    # Sua função de formatação aqui
    return f"{value * 100:.2f}%"

def process_excel_with_tu_columns(xls):
    result_dfs = {}
    
     # Nomes das colunas relacionadas a TU 1 e TU 2
    tu1_columns = ['TU 1 adm', 'TU 1 dcdf', 'TU 1 dcsp', 'TU 1 dcrj']
    tu2_columns = ['TU 2 adm', 'TU 2 dcdf', 'TU 2 dcsp', 'TU 2 dcrj']
    
    # Iterar sobre todas as abas do arquivo
    for sheet in xls.sheet_names:
        try:
            # Ler a aba atual sem cabeçalho
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
            
            # Encontrar a linha onde começa a tabela (após '###Tabela 3###')
            start_row = df[df[0].str.contains('###Tabela 3###', na=False)].index[0] + 1
            
            # Encontrar a linha onde termina a tabela (após '###Tabela 4###')
            end_row = df[df[0].str.contains('###Tabela 4###', na=False)].index[0]
            
            # Ler novamente a aba, começando a partir da linha correta até o final da tabela
            df = pd.read_excel(xls, sheet_name=sheet, header=start_row, nrows=end_row - start_row)  # Ajusta o cabeçalho
            
            # Verificar se o DataFrame está vazio
            if df.empty:
                print(f"A aba '{sheet}' está vazia. Pulando...")
                continue
            
            # Verificar se as colunas necessárias existem no DataFrame
            if not all(col in df.columns for col in tu1_columns + tu2_columns):
                print(f"A aba '{sheet}' não contém todas as colunas de TU 1 e TU 2. Pulando...")
                continue
            
            # Converter as colunas de percentuais para valores numéricos
            for col in tu1_columns + tu2_columns:
                df[col] = df[col].apply(convert_percentage)
            
            # Encontrar o menor valor de 'TU 1' e 'TU 2' separadamente
            df['Min_TU_1'] = df[tu1_columns].min(axis=1)
            df['Min_TU_2'] = df[tu2_columns].min(axis=1)
            
            # Aplicar a condição para verificar se o menor valor está entre 2% e 5%
            df['Resultado_TU_1'] = df['Min_TU_1'].apply(lambda x: 'Verdadeiro' if 0.02 <= x <= 0.05 else 'Falso')
            df['Resultado_TU_2'] = df['Min_TU_2'].apply(lambda x: 'Verdadeiro' if 0.02 <= x <= 0.05 else 'Falso')
            
            # Formatar os mínimos como porcentagem para a saída
            df['Min_TU_1'] = df['Min_TU_1'].apply(format_as_percentage)
            df['Min_TU_2'] = df['Min_TU_2'].apply(format_as_percentage)
            


            # Armazenar resultados processados
            result_dfs[sheet] = df[['ID', 'Host', 'Min_TU_1', 'Min_TU_2', 'Resultado_TU_1','Resultado_TU_2']]
        
        except Exception as e:
            print(f"Erro ao processar a aba '{sheet}': {e}")
    
    return result_dfs

# Processar todas as abas do arquivo
processed_data = process_excel_with_tu_columns(xls)

# Se houver dados processados, salvar em um arquivo Excel com abas separadas
if processed_data:
    with pd.ExcelWriter('resultado_processado.xlsx') as writer:
        for sheet_name, df in processed_data.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    print("Arquivo processado e salvo como 'resultado_processado.xlsx'.")
else:
    print("Nenhum dado foi processado.")


