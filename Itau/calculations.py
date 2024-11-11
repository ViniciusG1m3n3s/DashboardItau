import pandas as pd
import os
import plotly.express as px
import math
import streamlit as st

def load_data(usuario):
    # Definindo o diretório atual ou outro caminho desejado
    excel_file = f'dados_acumulados_{usuario}.xlsx'  # Caminho do arquivo Excel no diretório atual
    if os.path.exists(excel_file):
        try:
            # Tenta ler o arquivo Excel
            df_total = pd.read_excel(excel_file, engine='openpyxl')
        except Exception as e:
            # Caso ocorra algum erro ao ler o arquivo, exibe a mensagem de erro
            print(f"Erro ao carregar o arquivo {excel_file}: {e}")
            df_total = pd.DataFrame(columns=['Protocolo', 'Usuário', 'Status', 'Tempo de Análise', 'Próximo'])
    else:
        # Caso o arquivo não exista, exibe mensagem e retorna DataFrame vazio
        print(f"Arquivo não encontrado: {excel_file}")
        df_total = pd.DataFrame(columns=['Protocolo', 'Usuário', 'Status', 'Tempo de Análise', 'Próximo'])
    
    return df_total

# Função para salvar os dados no Excel do usuário logado
def save_data(df, usuario):
    # Definindo o diretório atual ou outro caminho desejado
    excel_file = f'dados_acumulados_{usuario}.xlsx'  # Caminho do arquivo Excel no diretório atual
    try:
        # Converte a coluna 'Tempo de Análise' para string para evitar erro de tipo
        df['Tempo de Análise'] = df['Tempo de Análise'].astype(str)
        
        # Tenta salvar o DataFrame no arquivo Excel
        with pd.ExcelWriter(excel_file, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, index=False)
        print(f"Dados salvos com sucesso em {excel_file}")
    
    except Exception as e:
        # Caso ocorra algum erro ao salvar, exibe a mensagem de erro
        print(f"Erro ao salvar os dados em {excel_file}: {e}")

def calcular_tmo_por_dia(df):
    df['Dia'] = df['Próximo'].dt.date
    df_finalizados = df[df['Status'] == 'FINALIZADO'].copy()
    df_tmo = df_finalizados.groupby('Dia').agg(
        Tempo_Total=('Tempo de Análise', 'sum'),
        Total_Protocolos=('Tempo de Análise', 'count')
    ).reset_index()
    df_tmo['TMO'] = (df_tmo['Tempo_Total'] / pd.Timedelta(minutes=1)) / df_tmo['Total_Protocolos']
    return df_tmo[['Dia', 'TMO']]

def calcular_produtividade_diaria(df):
    df['Dia'] = df['Próximo'].dt.date
    df_produtividade = df.groupby('Dia').agg(
        Andamento=('Status', lambda x: x[x == 'ANDAMENTO_PRE'].count()),
        Finalizado=('Status', lambda x: x[x == 'FINALIZADO'].count()),
        Reclassificado=('Status', lambda x: x[x == 'RECLASSIFICADO'].count())
    ).reset_index()
    df_produtividade['Produtividade'] = df_produtividade['Andamento'] + df_produtividade['Finalizado'] + df_produtividade['Reclassificado']
    return df_produtividade

def convert_to_timedelta_for_calculations(df):
    df['Tempo de Análise'] = pd.to_timedelta(df['Tempo de Análise'], errors='coerce')
    return df

def convert_to_datetime_for_calculations(df):
    df['Próximo'] = pd.to_datetime(df['Próximo'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
    return df
        
def format_timedelta(td):
    if pd.isnull(td):
        return "0 min"
    total_seconds = int(td.total_seconds())
    minutes, seconds = divmod(total_seconds, 60)
    return f"{minutes} min {seconds}s"

# Função para calcular o TMO por analista
def calcular_tmo(df_total):
    # Calcula o TMO por analista
    users_tmo = df_total['Usuário'].unique()
    df_tmo_analista = df_total[df_total['Usuário'].isin(users_tmo)].groupby('Usuário').agg(
        Tempo_Total=('Tempo de Análise', 'sum'),
        Total_Protocolos=('Tempo de Análise', 'count')
    ).reset_index()
    df_tmo_analista['TMO'] = df_tmo_analista['Tempo_Total'] / df_tmo_analista['Total_Protocolos']
    df_tmo_analista['TMO_Formatado'] = df_tmo_analista['TMO'].apply(lambda x: f"{int(x.total_seconds() // 60)}:{int(x.total_seconds() % 60):02}")

    return df_tmo_analista

# Função para calcular o ranking dinâmico
def calcular_ranking(df_total, selected_users):
    # Filtra o DataFrame com os usuários selecionados
    df_filtered = df_total[df_total['Usuário'].isin(selected_users)]

    df_ranking = df_filtered.groupby('Usuário').agg(
        Andamento=('Status', lambda x: x[x == 'ANDAMENTO_PRE'].count()),
        Finalizado=('Status', lambda x: x[x == 'FINALIZADO'].count()),
        Reclassificado=('Status', lambda x: x[x == 'RECLASSIFICADO'].count())
    ).reset_index()
    df_ranking['Total'] = df_ranking['Andamento'] + df_ranking['Finalizado'] + df_ranking['Reclassificado']
    df_ranking = df_ranking.sort_values(by='Total', ascending=False).reset_index(drop=True)
    df_ranking.index += 1
    df_ranking.index.name = 'Posição'

    # Define o tamanho dos quartis
    num_analistas = len(df_ranking)
    quartil_size = 4 if num_analistas > 12 else math.ceil(num_analistas / 4)

    # Função de estilo para os quartis dinâmicos
    def apply_dynamic_quartile_styles(row):
        if row.name <= quartil_size:
            color = 'rgba(135, 206, 250, 0.4)'  # Azul vibrante translúcido (primeiro quartil)
        elif quartil_size < row.name <= 2 * quartil_size:
            color = 'rgba(144, 238, 144, 0.4)'  # Verde vibrante translúcido (segundo quartil)
        elif 2 * quartil_size < row.name <= 3 * quartil_size:
            color = 'rgba(255, 255, 102, 0.4)'  # Amarelo vibrante translúcido (terceiro quartil)
        else:
            color = 'rgba(255, 99, 132, 0.4)'  # Vermelho vibrante translúcido (quarto quartil)
        return ['background-color: {}'.format(color) for _ in row]

    # Aplicar os estilos e retornar o DataFrame
    styled_df_ranking = df_ranking.style.apply(apply_dynamic_quartile_styles, axis=1).format(
        {'Andamento': '{:.0f}', 'Finalizado': '{:.0f}', 'Reclassificado': '{:.0f}', 'Total': '{:.0f}'}
    )

    return styled_df_ranking

def calcular_tempo_medio_analista(df_analista):
    # Filtra as linhas com status 'FINALIZADO', 'RECLASSIFICADO' ou 'ANDAMENTO_PRE'
    df_filtrado = df_analista[df_analista['Status'].isin(['FINALIZADO', 'RECLASSIFICADO', 'ANDAMENTO_PRE'])]
    
    # Calcula o tempo médio de análise considerando os status filtrados
    tempo_medio_analista = df_filtrado['Tempo de Análise'].mean()
    
    # Se o tempo médio for válido, formate-o
    if pd.notna(tempo_medio_analista):
        return format_timedelta(tempo_medio_analista)  # Supondo que format_timedelta já formate de acordo
    else:
        return 'Nenhum dado encontrado'

#MÉTRICAS INDIVIDUAIS
def calcular_metrica_analista(df_analista):
    # Verifica se a coluna "Carteira" está presente no DataFrame
    if 'Carteira' not in df_analista.columns:
        st.warning("A coluna 'Carteira' não está disponível nos dados. Verifique o arquivo carregado.")
        return None, None, None, None

    # Excluir os registros com "Carteira" como "Desconhecida"
    df_analista_filtrado = df_analista[df_analista['Carteira'] != "Desconhecida"]

    # Filtra os registros com status "FINALIZADO" e "RECLASSIFICADO" (desconsiderando "ANDAMENTO_PRE")
    df_filtrados = df_analista_filtrado[df_analista_filtrado['Status'].isin(['FINALIZADO', 'RECLASSIFICADO'])]

    # Calcula totais conforme os filtros de status
    total_finalizados = len(df_filtrados[df_filtrados['Status'] == 'FINALIZADO'])
    total_reclass = len(df_filtrados[df_filtrados['Status'] == 'RECLASSIFICADO'])
    total_andamento = len(df_analista[df_analista['Status'] == 'ANDAMENTO_PRE'])

    # Calcula o tempo total de análise considerando "FINALIZADO" e "RECLASSIFICADO" apenas
    tempo_total_analista = df_filtrados['Tempo de Análise'].sum()
    total_tarefas = total_finalizados + total_reclass
    tempo_medio_analista = tempo_total_analista / total_tarefas if total_tarefas > 0 else 0

    return total_finalizados, total_reclass, total_andamento, tempo_medio_analista

def calcular_tmo_equipe(df_total):
    return df_total[df_total['Status'] == 'FINALIZADO']['Tempo de Análise'].mean()

def calcular_filas_analista(df_analista):
    if 'Carteira' in df_analista.columns:
        # Filtra apenas os status relevantes para o cálculo (considerando FINALIZADO e RECLASSIFICADO)
        filas_finalizadas_analista = df_analista[
            df_analista['Status'].isin(['FINALIZADO', 'RECLASSIFICADO', 'ANDAMENTO_PRE'])
        ]
        
        # Agrupa por 'Carteira' e calcula a quantidade de FINALIZADO, RECLASSIFICADO e ANDAMENTO_PRE para cada fila
        carteiras_analista = filas_finalizadas_analista.groupby('Carteira').agg(
            Finalizados=('Status', lambda x: (x == 'FINALIZADO').sum()),
            Reclassificados=('Status', lambda x: (x == 'RECLASSIFICADO').sum()),
            Andamento=('Status', lambda x: (x == 'ANDAMENTO_PRE').sum()),
            TMO_médio=('Tempo de Análise', lambda x: x[x.index.isin(df_analista[(df_analista['Status'].isin(['FINALIZADO', 'RECLASSIFICADO']))].index)].mean())
        ).reset_index()

        # Converte o TMO médio para minutos e segundos
        carteiras_analista['TMO_médio'] = carteiras_analista['TMO_médio'].apply(format_timedelta)

        # Renomeia as colunas para exibição
        carteiras_analista = carteiras_analista.rename(
            columns={'Carteira': 'Fila', 'Finalizados': 'Finalizados', 'Reclassificados': 'Reclassificados', 'Andamento': 'Andamento', 'TMO_médio': 'TMO Médio por Fila'}
        )
        
        return carteiras_analista  # Retorna o DataFrame
    
    else:
        # Caso a coluna 'Carteira' não exista
        return pd.DataFrame({'Fila': [], 'Finalizados': [], 'Reclassificados': [], 'Andamento': [], 'TMO Médio por Fila': []})

def calcular_tmo_por_dia(df_analista):
    # Lógica para calcular o TMO por dia
    df_analista['Dia'] = df_analista['Próximo'].dt.date
    tmo_por_dia = df_analista.groupby('Dia').agg(TMO=('Tempo de Análise', 'mean')).reset_index()
    return tmo_por_dia

def calcular_carteiras_analista(df_analista):
    if 'Carteira' in df_analista.columns:
        filas_finalizadas = df_analista[(df_analista['Status'] == 'FINALIZADO') |
                                        (df_analista['Status'] == 'RECLASSIFICADO') |
                                        (df_analista['Status'] == 'ANDAMENTO_PRE')]

        carteiras_analista = filas_finalizadas.groupby('Carteira').agg(
            Quantidade=('Carteira', 'size'),
            TMO_médio=('Tempo de Análise', 'mean')
        ).reset_index()

        # Renomeando a coluna 'Carteira' para 'Fila' para manter consistência
        carteiras_analista = carteiras_analista.rename(columns={'Carteira': 'Fila'})

        return carteiras_analista
    else:
        return pd.DataFrame({'Fila': [], 'Quantidade': [], 'TMO Médio por Fila': []})
    
def get_points_of_attention(df):
    # Verifica se a coluna 'Carteira' existe no DataFrame
    if 'Carteira' not in df.columns:
        return "A coluna 'Carteira' não foi encontrada no DataFrame."
    
    # Filtra os dados para 'JV ITAU BMG' e outras carteiras
    dfJV = df[df['Carteira'] == 'JV ITAU BMG'].copy()
    dfOutras = df[df['Carteira'] != 'JV ITAU BMG'].copy()
    
    # Filtra os pontos de atenção com base no tempo de análise
    pontos_de_atencao_JV = dfJV[dfJV['Tempo de Análise'] > pd.Timedelta(minutes=2)]
    pontos_de_atencao_outros = dfOutras[dfOutras['Tempo de Análise'] > pd.Timedelta(minutes=5)]
    
    # Combina os dados filtrados
    pontos_de_atencao = pd.concat([pontos_de_atencao_JV, pontos_de_atencao_outros])

    # Verifica se o DataFrame está vazio
    if pontos_de_atencao.empty:
        return "Não existem dados a serem exibidos."

    # Cria o dataframe com as colunas 'PROTOCOLO', 'CARTEIRA' e 'TEMPO'
    pontos_de_atencao = pontos_de_atencao[['Protocolo', 'Carteira', 'Tempo de Análise']].copy()

    # Renomeia a coluna 'Tempo de Análise' para 'TEMPO'
    pontos_de_atencao = pontos_de_atencao.rename(columns={'Tempo de Análise': 'TEMPO'})

    # Converte a coluna 'TEMPO' para formato de minutos
    pontos_de_atencao['TEMPO'] = pontos_de_atencao['TEMPO'].apply(lambda x: f"{int(x.total_seconds() // 60)}:{int(x.total_seconds() % 60):02d}")

    # Remove qualquer protocolo com valores vazios ou NaN
    pontos_de_atencao = pontos_de_atencao.dropna(subset=['Protocolo'])

    # Remove as vírgulas e a parte ".0" do protocolo
    pontos_de_atencao['Protocolo'] = pontos_de_atencao['Protocolo'].astype(str).str.replace(',', '', regex=False)
    
    # Garantir que o número do protocolo não tenha ".0"
    pontos_de_atencao['Protocolo'] = pontos_de_atencao['Protocolo'].str.replace(r'\.0$', '', regex=True)

    return pontos_de_atencao

def calcular_tmo_por_carteira(df):
    # Verifica se as colunas 'Carteira' e 'Tempo de Análise' estão no DataFrame
    if 'Carteira' not in df.columns or 'Tempo de Análise' not in df.columns:
        return "As colunas 'Carteira' e/ou 'Tempo de Análise' não foram encontradas no DataFrame."

    # Agrupa os dados por carteira e calcula o tempo médio de análise para cada grupo
    tmo_por_carteira = df.groupby('Carteira')['Tempo de Análise'].mean().reset_index()

    # Converte o tempo médio de análise para minutos e segundos
    tmo_por_carteira['TMO'] = tmo_por_carteira['Tempo de Análise'].apply(
        lambda x: f"{int(x.total_seconds() // 60)}:{int(x.total_seconds() % 60):02d}"
    )

    # Seleciona apenas as colunas de interesse
    tmo_por_carteira = tmo_por_carteira[['Carteira', 'TMO']]

    return tmo_por_carteira
