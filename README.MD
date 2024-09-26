# Automatizando Planilhas de Baixa Potência

Este projeto fornece um script em Python para automatizar o tratamento de planilhas de baixa potência de ONT e OLT, mesclando diferentes planilhas, limpando os dados e classificando a potência com base em regras específicas.

## Funcionalidades

- **Limpeza de Dados:** Remove duplicatas, filtra registros com base no status e limpa valores nulos.
- **Verificação de Colunas:** Garante que todas as colunas necessárias estejam presentes.
- **Merge de Planilhas:** Mescla os dataframes com base na coluna 'Designador'.
- **Classificação de Potência:** Classifica a situação da potência conforme os valores de ONT e OLT.
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

### `clean_dataframe(df, status_col=None, status_value=None, designador_col='Designador')`
Limpa o dataframe removendo duplicatas e valores nulos, e opcionalmente filtra por uma coluna de status.

- **Parâmetros:**
  - `df (pd.DataFrame)`: O dataframe de entrada.
  - `status_col (str, opcional)`: Coluna de status para filtragem.
  - `status_value (str, opcional)`: Valor de status a ser filtrado.
  - `designador_col (str, padrão='Designador')`: Coluna para remoção de duplicatas e valores nulos.

### `merge_dataframes(df1, df2, on_col='Designador')`
Mescla dois dataframes com base em uma coluna comum, reindexando e concatenando os dados.

- **Parâmetros:**
  - `df1 (pd.DataFrame)`: Primeiro dataframe.
  - `df2 (pd.DataFrame)`: Segundo dataframe.
  - `on_col (str, padrão='Designador')`: Coluna chave para o merge.

### `situacao_potencia(potencia)`
Classifica a situação da potência de acordo com os valores de ONT.

- **Parâmetros:**
  - `potencia (float)`: Valor da potência a ser classificada.
  
- **Retorna:** Uma string indicando a situação ("FALHA DE LEITURA", "ATENUADO", "FORA DO PADRÃO", "OK").

## Estrutura de Saída

O script gera dois arquivos Excel:

- Um arquivo acumulado chamado `Baixa_potencia_geral.xlsx`.
- Um arquivo individual salvo no diretório `C:/python/diario/` com o nome fornecido pelo usuário.

## Licença

Este projeto está licenciado sob a [MIT License](LICENSE).

## Contato

Para mais informações ou dúvidas, entre em contato com: **wpquirino@gmail.com**
Linkedin: www.linkedin.com/in/welton-quirino


