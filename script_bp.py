import pandas as pd
import numpy as np
import os

# Mensagem inicial
print("Olá! Essa planilha tem duas saídas, uma individual e uma que acumula todas as planilhas anteriores.")
print()

# Função para obter o caminho do arquivo Excel
def get_excel_file_path():
    while True:
        print("Agora digite o caminho completo do arquivo Excel que quer tratar ou o arraste até aqui (com extensão .xlsx): ")
        file_name = input()
        if file_name.lower().endswith('.xlsx') and os.path.isfile(file_name):
            return file_name
        else:
            print("Erro: O arquivo deve ter a extensão '.xlsx' e deve existir. Tente novamente.")

# Função genérica para limpeza de dataframes
def clean_dataframe(df, status_col=None, status_value=None, id_col='ID'):
    if status_col and status_value:
        df = df[df[status_col].str.strip() == status_value]
    df = df.dropna(subset=[id_col]).drop_duplicates(subset=[id_col], keep='last').reset_index(drop=True)
    return df

# Função genérica para verificar colunas ausentes
def verificar_colunas(df, colunas):
    colunas_ausentes = [col for col in colunas if col not in df.columns]
    if colunas_ausentes:
        print(f"Colunas ausentes: {colunas_ausentes}")
    return df[colunas] if not colunas_ausentes else df

# Função para merge dos dataframes
def merge_dataframes(df1, df2, on_col='ID'):
    df2_reindexed = df2.set_index(on_col).reindex(df1[on_col]).reset_index()
    return pd.concat([df1.set_index(on_col), df2_reindexed.set_index(on_col)], axis=1).reset_index()

# Função para salvar e acumular dados
def salvar_acumular_dados(df_final, caminho_geral, output_individual):
    if os.path.exists(caminho_geral):
        df_acumulado = pd.read_excel(caminho_geral)
        df_acumulado = pd.concat([df_acumulado, df_final], ignore_index=True)
        df_acumulado['DATA'] = pd.to_datetime(df_acumulado['DATA'], dayfirst=True)
        df_acumulado['DATA'] = df_acumulado['DATA'].dt.strftime('%d/%m/%Y')
    else:
        df_acumulado = df_final
    df_acumulado.to_excel(caminho_geral, index=False)
    df_final.to_excel(output_individual, index=False)
    return df_acumulado

# Situação da potência usando função nomeada
def situacao_potencia(potencia):
    if pd.isna(potencia):
        return "FALHA DE LEITURA"
    elif potencia < -26.1:
        return "ATENUADO"
    elif potencia >= 0:
        return "FORA DO PADRÃO"
    else:
        return "OK"

def main():
    # Obtendo o caminho do arquivo Excel
    df_ = get_excel_file_path()

    # Exemplo de processamento
    print(f"O arquivo foi lido com sucesso.")

    # Carregamento de dados
    df1 = pd.read_excel(df_, sheet_name='Planilha1')
    df2 = pd.read_excel(df_, sheet_name='Planilha2')

    # Limpeza dos DataFrames
    df1_limpo = clean_dataframe(df1, status_col='Status', status_value='Concluída')
    df2.rename(columns={'ID': 'ID'}, inplace=True)
    df2_limpo = clean_dataframe(df2)

    # Colunas a serem verificadas
    cols_df1 = ['ID', 'Razão completamento 1', 'Razão completamento 2', 
             'Razão completamento 3', 'Zona de Trabalho']
    cols_df2 = ['ID', 'FREQUENCIA1', 'FREQUENCIA2']

    # Selecionando colunas
    df1_selected = verificar_colunas(df1_limpo, cols_df1)
    df2_selected = verificar_colunas(df2_limpo, cols_df2)

    # Merge dos dataframes
    df_merged = merge_dataframes(df1_selected, df2_selected)

    # Conversão de valores da coluna FREQUENCIA2 e FREQEUNCIA1 para numérico
    df_merged[['FREQUENCIA2', 'FREQUENCIA1']] = df_merged[['FREQUENCIA2', 'FREQUENCIA1']].apply(pd.to_numeric, errors='coerce')

    # Ajuste de valores de FREQUENCIA2 e criação de 'FREQUENCIA REAL'
    df_merged['FREQUENCIA1'] = df_merged['FREQUENCIA2'].apply(lambda valor: np.ceil(valor / 1000 * 10) / 10 if valor < -100 else np.ceil(valor * 10) / 10)
    df_merged['FREQUENCIA REAL'] = df_merged[['FREQUENCIA1', 'FREQUENCIA2']].max(axis=1)

    # Aplicando a função de situação de potência
    df_merged['SITUACAO'] = df_merged['FREQUENCIA REAL'].apply(situacao_potencia)
    df_merged['DATA'] = (pd.Timestamp.now() - pd.Timedelta(days=1)).strftime('%d/%m/%Y')

    # Ajuste da zona de trabalho
    df_merged['Zona de Trabalho'] = df_merged['Zona de Trabalho'].apply(lambda local: "EMPRESA1" if local in ["Zona 1", "Zona 2"] else "EMPRESA2")

    # Preparando o DataFrame final
    df_final = df_merged[['ID', 'DATA', 'Razão completamento 1', 'Razão completamento 2', 
                      'Razão completamento 3', 'Zona de Trabalho', 'FREQUENCIA1', 'FREQUENCIA2', 'FREQUENCIA REAL', 'SITUACAO']]

    # Salvando em Excel
    print("Criando o arquivo tratado...")
    print("Digite um nome para o arquivo ser salvo (sem extensão):")
    output_1 = input()
    output_2 = f"{output_1}.xlsx"

    output_individual = os.path.join("C:/python/diario", output_2)  # Usando barras normais para o caminho
    caminho_geral = 'c:/python/Baixa_potencia_geral.xlsx'

    df_acumulado = salvar_acumular_dados(df_final, caminho_geral, output_individual)

    # Mensagem de finalização
    print(f"A planilha acumulada foi salva com sucesso e tem {len(df_acumulado)} linhas agora.")
    print(f"Tudo certo! \n\nArquivo individual foi salvo no caminho com sucesso, com o nome {output_2}.")
    print(f"E o arquivo geral foi acumulado na planilha Baixa_potencia_geral no caminho {caminho_geral}.")
    print("Até amanhã!")
if __name__=="__main__":
    main()
