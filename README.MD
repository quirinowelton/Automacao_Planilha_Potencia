# Automatizando Planilhas de Baixa Frequência

Este projeto fornece um script em Python para automatizar o tratamento de planilhas de baixa frequência de sinal, mesclando diferentes planilhas, limpando os dados e classificando a potência com base em regras específicas.


## Funcionalidades

- **Limpeza de Dados:** Remove duplicatas, filtra registros com base no status e limpa valores nulos.
- **Verificação de Colunas:** Garante que todas as colunas necessárias estejam presentes.
- **Merge de Planilhas:** Mescla os dataframes com base na coluna 'ID'.
- **Classificação de Potência:** Classifica a situação da potência conforme os valores de FREQUENCIA1 e FREQUENCIA2.
- **Criação de Arquivos de Saída:** Gera dois arquivos Excel: um acumulado e outro individual.

## Ferramentas Utilizadas

- **Python:** Linguagem de programação principal.
- **pandas:** Biblioteca para manipulação de dados tabulares.
- **numpy:** Biblioteca para cálculos numéricos.
- **openpyxl:** Biblioteca para leitura e escrita de arquivos Excel.
- **Git:** Controle de versionamento do projeto.

## Instalação

Certifique-se de ter o Python instalado. Para instalar as dependências necessárias, use o `pip`:

```bash
pip install -r requirements.txt
```

## Uso

Para rodar o script, basta executar o comando abaixo:

```bash
python script_bp.py
```

Você será solicitado a fornecer o caminho do arquivo Excel que deseja processar, e o script executará o tratamento e gerará os arquivos de saída.

## Funcionalidades Principais

### `clean_dataframe(df, status_col=None, status_value=None, id_col='ID')`
Limpa o dataframe removendo duplicatas e valores nulos, e opcionalmente filtra por uma coluna de status.

```python
import pandas as pd

def clean_dataframe(df, status_col=None, status_value=None, ID_col='ID'):
    df = df.drop_duplicates(subset=[id_col])
    df = df.dropna()
    if status_col and status_value:
        df = df[df[status_col] == status_value]
    return df
```

### `merge_dataframes(df1, df2, on_col='ID')`
Mescla dois dataframes com base em uma coluna comum, reindexando e concatenando os dados.

```python
def merge_dataframes(df1, df2, on_col='ID'):
    merged_df = pd.merge(df1, df2, on=on_col)
    return merged_df
```

### `situacao_potencia(potencia)`
Classifica a situação da potência de acordo com os valores de FREQUENCIA2.
É considerado Falha de Leitura quando os dados da FREQUENCIA2 está vazio;
É considerado Atenuado quando a FREQUENCIA2 é menor que -15;
É considerado Fora do Padrão quando a potencia é maior que zero;
É considerado Ok quando o valor esta entre -15 e 0.
  
- **Retorna:** Uma string indicando a situação ("FALHA DE LEITURA", "ATENUADO", "FORA DO PADRÃO", "OK").

```python
def situacao_potencia(potencia):
    if pd.isna(potencia):
        return "FALHA DE LEITURA"
    elif potencia < -15:
        return "ATENUADO"
    elif potencia >= 0:
        return "FORA DO PADRÃO"
    else:
        return "OK"
```

## Estrutura de Saída

O script gera dois arquivos Excel:

- Um arquivo acumulado chamado `Baixa_frequencia_geral.xlsx`.
- Um arquivo individual salvo no diretório `C:/python/diario/` com o nome fornecido pelo usuário.

## Licença

Este projeto está licenciado sob a [MIT License](LICENSE).

## Contato

Para mais informações ou dúvidas, entre em contato com: **wpquirino@gmail.com**
Linkedin: www.linkedin.com/in/welton-quirino


