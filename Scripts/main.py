import pandas as pd 
import re

# Função para limpar o CPF/RG e números de contato (remover caracteres não numéricos)
def limpar_numeros(valor):
    return ''.join(filter(str.isdigit, str(valor))) if pd.notna(valor) else ""

# Função para carregar os dados do arquivo CSV
def carregar_dados(arquivo_csv):
    df = pd.read_csv(arquivo_csv, sep=';')
    print("Colunas disponíveis:", df.columns.tolist())  # Imprime os nomes das colunas
    return df

# Função para processar CPFs e RGs, tratar valores ausentes
def processar_dados(df):
    df['CPF_PACIENTE'] = df['CPF_PACIENTE'].apply(limpar_numeros)
    df['RG_PACIENTE'] = df['RG_PACIENTE'].apply(limpar_numeros)
    
    # DataFrame para armazenar os CPFs inválidos
    df_invalidos = df[(df['CPF_PACIENTE'] == "") | (df['RG_PACIENTE'] == "")]
    
    return df, df_invalidos

# Função para salvar CPFs inválidos em um novo arquivo Excel
def salvar_cpfs_invalidos(cpfs_invalidos, nome_arquivo):
    if not cpfs_invalidos.empty:
        cpfs_invalidos.to_excel(nome_arquivo, index=False)
        print(f"Planilha com CPFs/RGs inválidos foi criada: {nome_arquivo}")
    else:
        print("Não há CPFs ou RGs inválidos para salvar.")

# Função para processar endereços
def processar_enderecos(df):
    # Limpar a coluna CEP
    df['CEP'] = df['CEP'].apply(limpar_numeros)

    # Separar ENDERECO e NUMERO
    df['NUMERO'] = df['ENDERECO'].str.extract(r'(\d+)')[0]  # Captura o primeiro número encontrado
    df['ENDERECO'] = df['ENDERECO'].str.replace(r'\d+', '', regex=True)  # Remove números do endereço
    df['ENDERECO'] = df['ENDERECO'].str.replace(',', '', regex=True)  # Remove vírgulas

    return df

# Função para processar contatos
def processar_contatos(df_contatos):
    # Limpar DDD e CONTATO
    df_contatos['DDD'] = df_contatos['DDD'].apply(limpar_numeros)
    df_contatos['CONTATO'] = df_contatos['CONTATO'].apply(limpar_numeros)
    
    # Concatenar DDD e CONTATO
    df_contatos['CONTATO_COMPLETO'] = df_contatos['DDD'] + df_contatos['CONTATO']
    
    # Obter o contato mais recente para cada paciente
    df_contatos['DATA_CADASTRO'] = pd.to_datetime(df_contatos['DATA_CADASTRO'], dayfirst=True)
    df_ultimos_contatos = df_contatos.loc[df_contatos.groupby('ID_PACIENTE')['DATA_CADASTRO'].idxmax()]
    
    # Criar DataFrame para os contatos finais
    df_finais = pd.DataFrame()
    df_finais['ID_PACIENTE'] = df_ultimos_contatos['ID_PACIENTE']
    
    # Criar colunas para Telefone Fixo e Celular
    df_finais['Telefone Fixo'] = ""
    df_finais['Celular'] = ""
    df_finais['Outros Contatos'] = ""

    for paciente in df_ultimos_contatos['ID_PACIENTE'].unique():
        contatos_paciente = df_contatos[df_contatos['ID_PACIENTE'] == paciente]
        telefone_fixo = contatos_paciente[contatos_paciente['TIPO_CONTATO'] == 'fone fixo']
        celular = contatos_paciente[contatos_paciente['TIPO_CONTATO'] == 'celular']
        
        # Obter o contato mais recente
        if not telefone_fixo.empty:
            df_finais.loc[df_finais['ID_PACIENTE'] == paciente, 'Telefone Fixo'] = telefone_fixo.loc[telefone_fixo['DATA_CADASTRO'].idxmax(), 'CONTATO_COMPLETO']
        
        if not celular.empty:
            df_finais.loc[df_finais['ID_PACIENTE'] == paciente, 'Celular'] = celular.loc[celular['DATA_CADASTRO'].idxmax(), 'CONTATO_COMPLETO']

        # Concatenar outros contatos
        outros_contatos = contatos_paciente.loc[~contatos_paciente['DATA_CADASTRO'].isin([telefone_fixo['DATA_CADASTRO'].max(), celular['DATA_CADASTRO'].max()]), 'CONTATO_COMPLETO']
        df_finais.loc[df_finais['ID_PACIENTE'] == paciente, 'Outros Contatos'] = '; '.join(outros_contatos)

    return df_finais

# Função principal
def main():
    # Carregar dados dos pacientes
    arquivo_pacientes = '../Pacientes.csv'  # Nome do arquivo de entrada
    df_pacientes = carregar_dados(arquivo_pacientes)
    
    # Processar dados (manter todos os dados e identificar inválidos)
    df_limpo, cpfs_invalidos = processar_dados(df_pacientes)
    
    # Remover linhas sem CPF e RG para o arquivo limpo
    df_limpo = df_limpo[(df_limpo['CPF_PACIENTE'] != "") | (df_limpo['RG_PACIENTE'] != "")]
    
    # Salvar planilha de CPFs/RGs inválidos
    salvar_cpfs_invalidos(cpfs_invalidos, 'CPF_INVALIDO.xlsx')
    
    # Carregar dados dos endereços
    arquivo_enderecos = '../Enderecos.csv'
    df_enderecos = carregar_dados(arquivo_enderecos)
    
    # Filtrar apenas o último endereço por ID_PACIENTE usando a DATA_CRIACAO
    df_enderecos['DATA_CRIACAO'] = pd.to_datetime(df_enderecos['DATA_CRIACAO'], dayfirst=True)  # Ajustar formato da data
    df_ultimo_endereco = df_enderecos.loc[df_enderecos.groupby('ID_PACIENTE')['DATA_CRIACAO'].idxmax()]

    # Processar os endereços para separar texto e número
    df_ultimo_endereco = processar_enderecos(df_ultimo_endereco)

    # Carregar dados dos contatos
    arquivo_contatos = '../Contatos.csv'
    df_contatos = carregar_dados(arquivo_contatos)
    
    # Processar os contatos
    df_finais_contatos = processar_contatos(df_contatos)

    # Unir os dados dos pacientes com os endereços (apenas o último) e contatos
    df_final = pd.merge(df_limpo, df_ultimo_endereco[['ID_PACIENTE', 'ENDERECO', 'NUMERO', 'CEP', 'BAIRRO', 'CIDADE', 'ESTADO']], 
                        how='left', on='ID_PACIENTE')
    df_final = pd.merge(df_final, df_finais_contatos, how='left', on='ID_PACIENTE')
    
    # Reorganizar as colunas
    df_final = df_final[['ID_PACIENTE', 'NOME_PACIENTE', 'CPF_PACIENTE', 'RG_PACIENTE', 'ENDERECO', 'NUMERO', 'CEP', 'BAIRRO', 'CIDADE', 'ESTADO', 'Telefone Fixo', 'Celular', 'Outros Contatos']]
    
    # Salvar a planilha final
    df_final.to_excel('Pacientes.xlsx', index=False)

# Executar o programa
if __name__ == "__main__":
    main()
