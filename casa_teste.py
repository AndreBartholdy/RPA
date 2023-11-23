import pandas as pd
import os
import datetime
from colorama import Fore
from openpyxl.styles import Font, PatternFill

# Constantes
diretorio = "C:\\Users\\silas.zimmermann\\Desktop\\Andre\\RPA GILBERTO\\powerbi_employer\\arquivos_employer\\extratos_conta_corrente"
coluna_nomes_empresas = 'Nomes das Empresas'
coluna_saldo_inicial = 'Saldo Inicial'

indice_coluna_alvo = 14
soma_valores = 0.0
dia_atual = datetime.date(2023, 7, 25)
dia = dia_atual.strftime("%d/%m/%y")
# print(dia)
data_limite = dia_atual - datetime.timedelta(days=7)

# DATA PARA ENCONTRAR O DATAFRAME 3 E 4 #####################################################
data_procurada_str = '25/07/2023'
data_procurada_sheet = '25.07.2023'
data_procurada = pd.to_datetime(data_procurada_str, format='%d/%m/%Y')

df_padrao = pd.DataFrame()

def contar_arquivos_por_empresa(diretorio, nomes_empresas):
    contagem_empresas = {empresa: 0 for empresa in nomes_empresas}
    for arquivo in os.listdir(diretorio):
        if arquivo.endswith(".xls"):
            nome_base = arquivo.split(' ')[0]
            for empresa in nomes_empresas:
                if empresa in nome_base:
                    contagem_empresas[empresa] += 1
    return contagem_empresas

def carregar_dados_arquivo(arquivo_path):
    return pd.read_excel(arquivo_path)

def extrair_datas_iguais(df, dia):
    coluna_a = df.iloc[:, 0]
    padrao_data = r'\b(\d{2}/\d{2}/\d{2})\b'
    datas = coluna_a.str.extractall(padrao_data)
    datas = datas[0].unique() if not datas.empty else []
    datas_iguais = [data for data in datas if data == dia]
    return datas_iguais

def encontrar_primeira_linha_com_data(df, data_alvo):
    coluna_a = df.iloc[:, 0]
    padrao_data = r'\b(\d{2}/\d{2}/\d{2})\b'
    datas = coluna_a.str.extractall(padrao_data)
    datas = datas[0].unique() if not datas.empty else []

    # Verifica se a data_alvo está na lista de datas
    if data_alvo in datas:
        mask = (coluna_a == data_alvo)
        primeira_linha_idx = mask.idxmax()
        primeira_linha = df.iloc[primeira_linha_idx]

        # Verifica se a coluna O (índice 14) contém números na primeira linha
        coluna_O = primeira_linha.iloc[14]
        if pd.notna(coluna_O) and isinstance(coluna_O, (int, float)):
            return primeira_linha
    else:
        # Ordena as datas em ordem reversa
        datas_antes = sorted([data for data in datas], reverse=False)
        for data in datas_antes:
            mask = (coluna_a == data)
            primeira_linha_idx = mask.idxmax()
            primeira_linha = df.iloc[primeira_linha_idx]

            # Verifica se a coluna O (índice 14) contém números na primeira linha
            coluna_O = primeira_linha.iloc[14]
            if pd.notna(coluna_O) and isinstance(coluna_O, (int, float)):
                return primeira_linha

    return None

def encontrar_valores_numericos_coluna_O(df, data_alvo):
    # Converta a data_alvo para datetime.date
    data_alvo = datetime.datetime.strptime(data_alvo, '%d/%m/%y').date()
    primeira_linha_com_data = encontrar_primeira_linha_com_data(df, data_alvo)
    if primeira_linha_com_data is not None:
        coluna_O = primeira_linha_com_data.iloc[14]
        if pd.notna(coluna_O) and isinstance(coluna_O, (int, float)):
            return coluna_O
    return None

def encontrar_saldo_anterior(df):
    coluna_D = df.iloc[:, 3]  # Coluna D
    coluna_O = df.iloc[:, 14]  # Coluna O

    for idx, valor_D in enumerate(coluna_D):
        if isinstance(valor_D, str) and "saldo anterior" in valor_D.lower():
            valor_O = coluna_O[idx]
            if pd.notna(valor_O) and isinstance(valor_O, (int, float)):
                return valor_O

    return None

# Lista de nomes das empresas
nomes_empresas = ["8MRS", "ABRE", "AGRO", "BENE", "BNE", "EGMO", "EORH", "ETEC", "ETT", "IEGE", "MARA", "MYRH", "PALM", "PEOPLE", "RCLA", "SAS", "SRH", "STIO", "WAS"]

# Listar todos os arquivos no diretório
mapeamento_empresas = contar_arquivos_por_empresa(diretorio, nomes_empresas)
dados_empresas = {empresa: {'quantidade_arquivos': 0, 'valor_total': 0.0} for empresa in nomes_empresas}

# Criar um DataFrame com os nomes das empresas
df1 = pd.DataFrame(nomes_empresas, columns=[coluna_nomes_empresas])

for arquivo in os.listdir(diretorio):
    if arquivo.endswith(".xls"):
        # Lê o arquivo Excel
        df_arquivo = pd.read_excel(os.path.join(diretorio, arquivo))
        
        # Verifica se o índice da coluna alvo está dentro do range de colunas do DataFrame
        if indice_coluna_alvo < len(df_arquivo.columns):
            # Converte os valores da coluna alvo em valores numéricos e soma
            valores_numericos = pd.to_numeric(df_arquivo.iloc[:, indice_coluna_alvo], errors='coerce')
            valores_numericos = valores_numericos.dropna()
            soma_valores = valores_numericos.sum()
            
            # Obtém o nome da empresa com base no nome do arquivo
            nome_empresa = None
            for empresa in nomes_empresas:
                if empresa in arquivo:
                    nome_empresa = empresa
                    break
            
            # Atualiza a contagem de arquivos e o valor total da empresa
            if nome_empresa:
                dados_empresas[nome_empresa]['quantidade_arquivos'] += 1
                dados_empresas[nome_empresa]['valor_total'] += soma_valores
                
                datas_iguais = extrair_datas_iguais(df_arquivo, dia)
                # print(f"{nome_empresa} = Foi encontrado {len(datas_iguais)} datas iguais a '{dia}': {', '.join(map(str, datas_iguais))}")
                
                # Verifica a regra do dia atual
                datas_iguais = extrair_datas_iguais(df_arquivo, dia)
                if dia in datas_iguais:
                    valor_numerico = encontrar_valores_numericos_coluna_O(df_arquivo, dia)
                    if valor_numerico is not None:
                        novo_dado = pd.DataFrame({'Empresa': [nome_empresa], 'Saldo Inicial': [valor_numerico]})
                        df_padrao = pd.concat([df_padrao, novo_dado], ignore_index=True)
                        # print(Fore.GREEN + f"Valor Numérico da Coluna O para {nome_empresa}: {valor_numerico}" + Fore.RESET)
                    else:
                        novo_dado = pd.DataFrame({'Empresa': [nome_empresa], 'Saldo Inicial': 0.0})
                        df_padrao = pd.concat([df_padrao, novo_dado], ignore_index=True)
                        # print(Fore.RED + f"Nenhum valor numérico encontrado na coluna O para {nome_empresa}." + Fore.RESET)aaaa
                else:
                    # Verifica a regra de datas_antes
                    datas_antes = sorted([data for data in datas_iguais if datetime.datetime.strptime(data, '%d/%m/%y').date() <= data_limite], reverse=True)
                    if datas_antes:
                        data_antes = datas_antes[0]
                        valor_numerico = encontrar_valores_numericos_coluna_O(df_arquivo, data_antes)
                        if valor_numerico is not None:
                            novo_dado = pd.DataFrame({'Empresa': [nome_empresa], 'Saldo Inicial': [valor_numerico]})
                            df_padrao = pd.concat([df_padrao, novo_dado], ignore_index=True)
                            # print(Fore.YELLOW + f"Valor Numérico da Coluna O para {nome_empresa} (Data anterior {data_antes}): {valor_numerico}" + Fore.RESET)
                        else:
                            novo_dado = pd.DataFrame({'Empresa': [nome_empresa], 'Saldo Inicial': 0.0})
                            df_padrao = pd.concat([df_padrao, novo_dado], ignore_index=True)
                            # print(Fore.RED + f"Nenhum valor numérico encontrado na coluna O para {nome_empresa} (Data anterior {data_antes})." + Fore.RESET)
                    else:
                        # Verifica a regra de encontrar_saldo_anterior
                        valor_saldo_anterior = encontrar_saldo_anterior(df_arquivo)
                        if valor_saldo_anterior is not None:
                            novo_dado = pd.DataFrame({'Empresa': [nome_empresa], 'Saldo Inicial': [valor_saldo_anterior]})
                            df_padrao = pd.concat([df_padrao, novo_dado], ignore_index=True)
                            # print(Fore.BLUE + f"Valor Numérico da Coluna O para {nome_empresa} (Saldo Anterior): {valor_saldo_anterior}" + Fore.RESET)
                        else:
                            novo_dado = pd.DataFrame({'Empresa': [nome_empresa], 'Saldo Inicial': 0.0})
                            df_padrao = pd.concat([df_padrao, novo_dado], ignore_index=True)
                            # print(Fore.RED + f"Nenhum valor numérico encontrado na coluna O para {nome_empresa} (Saldo Anterior)." + Fore.RESET)

# CAMINHO PARA ENCONTRAR O DATAFRAME 3 #######################################################################################################
diretorio_receber = r'C:\Users\silas.zimmermann\Desktop\Andre\RPA GILBERTO\powerbi_employer\relatorio_mxm\Movimento Julho 2023.xlsx'

df_terceiro = pd.read_excel(diretorio_receber, header=2)

coluna_empresa = df_terceiro.columns[4]
coluna_valor_pagamento = df_terceiro.columns[3]
empresas_na_data = df_terceiro.loc[df_terceiro['Competência'] == data_procurada]
somas_empresas = []

for empresa in nomes_empresas:
    if empresa in empresas_na_data[coluna_empresa].values:
        valor_total = empresas_na_data[empresas_na_data[coluna_empresa] == empresa][coluna_valor_pagamento].sum()
    else:
        valor_total = 0
    somas_empresas.append(valor_total)

df3 = pd.DataFrame({'Recebimento': somas_empresas})

# CAMINHO PARA ENCONTRAR O DATAFRAME 4 #######################################################################################################
df_quarto = pd.read_excel(diretorio_receber, header=2)

coluna_empresa = df_quarto.columns[4]
coluna_valor_recebimento = df_quarto.columns[5]
coluna_aplicacao = df_quarto.columns[7]
empresas_na_data = df_quarto.loc[df_quarto['Competência'] == data_procurada]
somas_empresas = []

for empresa in nomes_empresas:
    if empresa in empresas_na_data[coluna_empresa].values:
        valor_total = empresas_na_data[empresas_na_data[coluna_empresa] == empresa][coluna_valor_recebimento].sum()
    else:
        valor_total = 0
    somas_empresas.append(valor_total)

df4 = pd.DataFrame({'Pagamento': somas_empresas})
df4 = df4 * (-1)
         
# CAMINHO PARA ENCONTRAR O DATAFRAME 2 #######################################################################################################
df2 = df_padrao.groupby('Empresa')['Saldo Inicial'].sum().reset_index()
df2 = df2.drop(['Empresa'], axis=1)
# print(Fore.YELLOW + f"{df2}" + Fore.RESET)
for empresa, dados in dados_empresas.items():
    quantidade_arquivos = dados['quantidade_arquivos']
    valor_total = dados['valor_total']
    # print(f"{empresa} = {quantidade_arquivos} - {valor_total:.2f}")

# DF5 E DF6 ##################################################################################################################

soma_aplicacao = {}
soma_resgate = {}

def filtro_customizado(texto):
    texto_lower = texto.lower()  # Converter para minúsculas
    return ("aplic" in texto_lower or "resg" in texto_lower) and "aplicativo" not in texto_lower

aplicacao_filtrada = df_quarto[df_quarto['Histórico'].apply(filtro_customizado)]
df_quarto[coluna_valor_pagamento] = pd.to_numeric(df_quarto[coluna_valor_pagamento], errors='coerce')
df_quarto[coluna_valor_recebimento] = pd.to_numeric(df_quarto[coluna_valor_recebimento], errors='coerce')

for empresa in nomes_empresas:
    empresa_filtrada = aplicacao_filtrada[(aplicacao_filtrada['Cd.Empresa'] == empresa) & (aplicacao_filtrada['Competência'] == data_procurada)]
    soma_aplicacao[empresa] = empresa_filtrada[coluna_valor_pagamento].sum()
    soma_resgate[empresa] = empresa_filtrada[coluna_valor_recebimento].sum()

df5 = pd.DataFrame.from_dict(soma_aplicacao, orient='index', columns=['Aplicação'])
df5 = (df5.reset_index(drop=True) * (-1))

df6 = pd.DataFrame.from_dict(soma_resgate, orient='index', columns=['Resgate'])
df6 = df6.reset_index(drop=True)

##############################################################################################################################

df8 = pd.DataFrame(columns=['Saldo Disponível'])
df8['Saldo Disponível'] = df2['Saldo Inicial'] + df3['Recebimento'] + df4['Pagamento'] + df5['Aplicação'] + df6['Resgate']

# Concatenar os DataFrames df1, df2, df3 e df2
df_concatenado = pd.concat([df1, df2, df3, df4, df5, df6, df8], axis=1, ignore_index=False)

# Crie um escritor Excel Pandas
df_concatenado.to_excel('planilha_resultados2.xlsx', sheet_name=(data_procurada_sheet), index=False)