import os
import pandas as pd
import re

nome_cliente = 'Dr Claudio'
id_cliente = 'cliniclaudio'

# Função para limpar CPF/RG e números de contato (remover caracteres não numéricos)
def limpar_numeros(valor):
    return ''.join(filter(str.isdigit, str(valor))) if pd.notna(valor) else ""

# Função para carregar dados do arquivo CSV
def carregar_dados(arquivo_csv):
    df = pd.read_csv(arquivo_csv, sep=';')
    print("Colunas disponíveis:", df.columns.tolist())
    return df

# Função para processar CPFs e RGs e tratar valores ausentes
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
    df['CEP'] = df['CEP'].apply(limpar_numeros)
    df['NUMERO'] = df['ENDERECO'].str.extract(r'(\d+)')[0]
    df['ENDERECO'] = df['ENDERECO'].str.replace(r'\d+', '', regex=True)
    df['ENDERECO'] = df['ENDERECO'].str.replace(',', '', regex=True)
    return df

# Função para processar contatos
def processar_contatos(df_contatos):

    df_contatos['DDD'] = df_contatos['DDD'].apply(limpar_numeros)
    df_contatos['CONTATO'] = df_contatos['CONTATO'].apply(limpar_numeros)
    df_contatos['CONTATO_COMPLETO'] = df_contatos['DDD'] + df_contatos['CONTATO']
    df_contatos['DATA_CADASTRO'] = pd.to_datetime(df_contatos['DATA_CADASTRO'], dayfirst=True)
    
    df_ultimos_contatos = df_contatos.loc[df_contatos.groupby('ID_PACIENTE')['DATA_CADASTRO'].idxmax()]
    
    df_finais = pd.DataFrame()
    df_finais['ID_PACIENTE'] = df_ultimos_contatos['ID_PACIENTE']
    df_finais['Telefone Fixo'] = ""
    df_finais['Celular'] = ""
    df_finais['Outros Contatos'] = ""

    for paciente in df_ultimos_contatos['ID_PACIENTE'].unique():
        contatos_paciente = df_contatos[df_contatos['ID_PACIENTE'] == paciente]
        telefone_fixo = contatos_paciente[contatos_paciente['TIPO_CONTATO'] == 'fone fixo']
        celular = contatos_paciente[contatos_paciente['TIPO_CONTATO'] == 'celular']
        
        if not telefone_fixo.empty:
            df_finais.loc[df_finais['ID_PACIENTE'] == paciente, 'Telefone Fixo'] = telefone_fixo.loc[telefone_fixo['DATA_CADASTRO'].idxmax(), 'CONTATO_COMPLETO']
        
        if not celular.empty:
            df_finais.loc[df_finais['ID_PACIENTE'] == paciente, 'Celular'] = celular.loc[celular['DATA_CADASTRO'].idxmax(), 'CONTATO_COMPLETO']
        
        outros_contatos = contatos_paciente.loc[~contatos_paciente['DATA_CADASTRO'].isin([telefone_fixo['DATA_CADASTRO'].max(), celular['DATA_CADASTRO'].max()]), 'CONTATO_COMPLETO']
        df_finais.loc[df_finais['ID_PACIENTE'] == paciente, 'Outros Contatos'] = '; '.join(outros_contatos)
    
    return df_finais

# Função para processar os horários de agendamentos
def processar_horarios(df_agendamentos):
    df_agendamentos['DATA_AGENDA'] = pd.to_datetime(df_agendamentos['DATA_AGENDA'], format='%d/%m/%Y %H:%M')
    df_agendamentos['Hora Início'] = df_agendamentos['DATA_AGENDA'].dt.strftime('%H:%M')
    df_agendamentos['Hora Final'] = (df_agendamentos['DATA_AGENDA'] + pd.to_timedelta(df_agendamentos['DURACAO_AGENDA'], unit='m')).dt.strftime('%H:%M')
    return df_agendamentos

# Função para criar pasta para salvar as planilhas
def criar_pasta(nome_pasta):
    if not os.path.exists(nome_pasta):
        os.makedirs(nome_pasta)
    return nome_pasta

# Função principal para criar Pacientes.xlsx e Agendamentos.xlsx
def main():

    # Variáveis de caminho dos arquivos
    caminho_pacientes = 'Pacientes.csv'
    caminho_enderecos = 'Enderecos.csv'
    caminho_contatos = 'Contatos.csv'
    caminho_agendamentos = 'Agendamentos.csv'
    
    # Criar a pasta para salvar as planilhas
    nome_pasta = criar_pasta(f"{id_cliente}_{nome_cliente}")

    # Carregar dados dos pacientes
    df_pacientes = carregar_dados(caminho_pacientes)
    df_limpo, cpfs_invalidos = processar_dados(df_pacientes)
    df_limpo = df_limpo[(df_limpo['CPF_PACIENTE'] != "") | (df_limpo['RG_PACIENTE'] != "")]
    salvar_cpfs_invalidos(cpfs_invalidos, os.path.join(nome_pasta, 'CPF_INVALIDO.xlsx'))
    
    # Carregar e processar endereços
    df_enderecos = carregar_dados(caminho_enderecos)
    df_enderecos['DATA_CRIACAO'] = pd.to_datetime(df_enderecos['DATA_CRIACAO'], dayfirst=True)
    df_ultimo_endereco = df_enderecos.loc[df_enderecos.groupby('ID_PACIENTE')['DATA_CRIACAO'].idxmax()]
    df_ultimo_endereco = processar_enderecos(df_ultimo_endereco)
    
    # Carregar e processar contatos
    df_contatos = carregar_dados(caminho_contatos)
    df_finais_contatos = processar_contatos(df_contatos)
    
    # Unir dados de pacientes, endereços e contatos
    df_final = pd.merge(df_limpo, df_ultimo_endereco[['ID_PACIENTE', 'ENDERECO', 'NUMERO', 'CEP', 'BAIRRO', 'CIDADE', 'ESTADO']], on='ID_PACIENTE', how='left', validate='one_to_one')
    df_final = pd.merge(df_final, df_finais_contatos, on='ID_PACIENTE', how='left', validate='one_to_one')
    
    df_final = df_final[['ID_PACIENTE', 'NOME_PACIENTE', 'CPF_PACIENTE', 'RG_PACIENTE', 'ENDERECO', 'NUMERO', 'CEP', 'BAIRRO', 'CIDADE', 'ESTADO', 'Telefone Fixo', 'Celular', 'Outros Contatos']]
    df_final.columns = ['Id Paciente', 'Nome do Paciente', 'CPF', 'RG', 'Endereço', 'Número', 'CEP', 'Bairro', 'Cidade', 'Estado', 'Celular', 'Telefone Fixo', 'Outros Contatos']
    print("Colunas de df_final:", df_final.columns)
    df_final.to_excel(os.path.join(nome_pasta, 'Pacientes.xlsx'), index=False)
    print("Planilha Pacientes.xlsx criada.")
    
    # Verifica se a planilha Pacientes.xlsx foi criada antes de criar Agendamentos.xlsx
    caminho_pacientes_final = os.path.join(nome_pasta, 'Pacientes.xlsx')
    
    if os.path.exists(caminho_pacientes_final):
        # Carregar e processar agendamentos
        df_agendamentos = carregar_dados(caminho_agendamentos)
        df_agendamentos = processar_horarios(df_agendamentos)
        
        # Substituir os valores da coluna Status
        status_mapping = {
            "atendido": "Checkout",
            "confirmado": "Confirmed",
            "desmarcado": "Canceled",
            "faltou": "Missed"
        }
        df_agendamentos['STATUS_AGENDA'] = df_agendamentos['STATUS_AGENDA'].map(status_mapping).fillna(df_agendamentos['STATUS_AGENDA'])
        
        df_agendamentos.rename(columns={'ID_PACIENTE': 'Id Paciente'}, inplace=True)
       
        # Mesclar agendamentos com nome do paciente
        df_agendamentos_final = pd.merge(df_agendamentos, df_final[['Id Paciente', 'Nome do Paciente']], on='Id Paciente', how='left', validate='many_to_one')      
        df_agendamentos_final = df_agendamentos_final[['Id Paciente', 'Nome do Paciente', 'DENTISTA', 'DATA_AGENDA', 'Hora Início', 'Hora Final', 'STATUS_AGENDA', 'PROCEDIMENTO']]       
        df_agendamentos_final.columns = ['Id Paciente', 'Nome do Paciente', 'Dentista', 'Data Agenda', 'Hora Início', 'Hora Final', 'Status', 'Procedimento']


        
        df_agendamentos_final.to_excel(os.path.join(nome_pasta, 'Agendamentos.xlsx'), index=False)
        print("Planilha Agendamentos.xlsx criada.")
    else:
        print("Erro ao criar a planilha Pacientes.xlsx. A criação de Agendamentos.xlsx foi interrompida.")

# Executar o código
if __name__ == "__main__":
    main()
