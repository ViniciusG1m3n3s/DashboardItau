import plotly.express as px
import os
import pandas as pd
import streamlit as st
import plotly.graph_objs as go

def plot_produtividade_diaria(df_produtividade, custom_colors):
    if df_produtividade.empty or 'Dia' not in df_produtividade.columns or 'Produtividade' not in df_produtividade.columns:
        st.warning("Não há dados para exibir no gráfico de produtividade diária.")
        return None
    
    fig_produtividade = px.line(
        df_produtividade,
        x='Dia',
        y='Produtividade',
        color_discrete_sequence=custom_colors,
        labels={'Produtividade': 'Total de Cadastros'},
        line_shape='linear',
        markers=True
    )
    return fig_produtividade

def plot_tmo_por_dia(df_tmo, custom_colors):
    if df_tmo.empty or 'Dia' not in df_tmo.columns or 'TMO' not in df_tmo.columns:
        st.warning("Não há dados para exibir no gráfico de TMO por dia.")
        return None

    if isinstance(df_tmo['TMO'].iloc[0], str):
        df_tmo['TMO'] = pd.to_timedelta(df_tmo['TMO'].apply(lambda x: x.replace(' min', 'm').replace('s', 's')))
    
    df_tmo['TMO_Formatado'] = df_tmo['TMO'].apply(lambda x: f"{int(x.total_seconds() // 60)}:{int(x.total_seconds() % 60):02d}")
    
    fig_tmo_linha = px.line(
        df_tmo,
        x='Dia',
        y='TMO',
        labels={'TMO': 'TMO (min)', 'Dia': 'Data'},
        color_discrete_sequence=custom_colors,
        line_shape='linear',
        markers=True
    )
    
    fig_tmo_linha.update_layout(
        yaxis=dict(
            tickvals=df_tmo['TMO'],
            ticktext=df_tmo['TMO_Formatado']
        )
    )

    fig_tmo_linha.update_traces(
        text=df_tmo['TMO_Formatado'],
        textposition='top center',
        hovertemplate='Data = %{x|%d/%m/%Y}<br>TMO = %{text}'
    )

    return fig_tmo_linha

def plot_status_pie(total_finalizados, total_reclass, total_andamento, custom_colors):
    fig_status = px.pie(
        names=['Finalizado', 'Reclassificado', 'Andamento'],
        values=[total_finalizados, total_reclass, total_andamento],
        color_discrete_sequence=custom_colors
    )
    fig_status.update_traces(
        hovertemplate='Tarefas %{label} = %{value}<extra></extra>',
    )
    fig_status.update_layout(
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.1,
            xanchor="center",
            x=0.5
        )
    )
    return fig_status

# Função para gerar o gráfico de TMO por analista
def grafico_tmo(df_tmo_analista, custom_colors):
    # Gráfico de barras de TMO por analista em minutos
    fig_tmo_analista = px.bar(
        df_tmo_analista,
        x='Usuário',
        y=df_tmo_analista['TMO'].dt.total_seconds() / 60,  # TMO em minutos
        labels={'y': 'Tempo Médio Operacional (min)', 'x': 'Analista', 'Usuário': 'Analista'},
        text=df_tmo_analista['TMO_Formatado'],
        color_discrete_sequence=custom_colors
    )
    fig_tmo_analista.update_traces(
        textposition='outside',  # Exibe o tempo formatado fora das barras
        hovertemplate='Analista = %{x}<br>TMO = %{text}<extra></extra>',
        text=df_tmo_analista['TMO_Formatado']
    )
    return fig_tmo_analista

def grafico_status_analista(total_finalizados_analista, total_reclass_analista, total_andamento_analista, custom_colors):
    fig_status = px.pie(
        names=['FINALIZADO', 'RECLASSIFICADO', 'ANDAMENTO_PRE'],
        values=[total_finalizados_analista, total_reclass_analista, total_andamento_analista],
        color_discrete_sequence=custom_colors
    )
    fig_status.update_traces(
        hovertemplate='Tarefas %{label} = %{value}<extra></extra>',
    )
    fig_status.update_layout(
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.1,
            xanchor="center",
            x=0.5
        )
    )
    return fig_status

def grafico_filas_analista(carteiras_analista, custom_colors):
    fig = px.pie(
        names=carteiras_analista['Fila'],
        values=carteiras_analista['Quantidade'],
        color_discrete_sequence=custom_colors
    )
    fig.update_traces(
        hovertemplate='Fila = %{label}<br>Quantidade = %{value}<extra></extra>',
    )
    fig.update_layout(
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.1,
            xanchor="center",
            x=0.5
        )
    )
    return fig

def grafico_tmo_analista(df_tmo_analista, custom_colors, analista_selecionado):
    # Verifica se o DataFrame está vazio antes de acessar qualquer valor
    if df_tmo_analista.empty:
        st.warning("Não há dados para o analista selecionado.")
        # Retorna um gráfico vazio (ou um gráfico de barras com valores 0, por exemplo)
        return go.Figure()

    # Converte a coluna 'TMO' para o formato de tempo, se for uma string
    if isinstance(df_tmo_analista['TMO'].iloc[0], str):
        # Converte a string para o formato Timedelta
        df_tmo_analista['TMO'] = pd.to_timedelta(df_tmo_analista['TMO'].apply(lambda x: x.replace(' min', 'm').replace('s', 's')))
    
    # Formata a coluna 'TMO' para apenas minutos e segundos
    df_tmo_analista['TMO_formatado'] = df_tmo_analista['TMO'].apply(lambda x: f"{int(x.total_seconds() // 60)}:{int(x.total_seconds() % 60):02d}")

    # Cria o gráfico de barras
    fig = px.bar(
        df_tmo_analista,
        x='Dia',
        y='TMO',
        hover_name='Dia',
        hover_data={'TMO': True},  # Adiciona 'TMO' ao hover
        color_discrete_sequence=custom_colors
    )

    # Atualiza o eixo Y para mostrar em minutos e segundos
    fig.update_layout(
        yaxis_tickformat="M",  # Alinha a formatação para minutos
        yaxis=dict(
            tickvals=df_tmo_analista['TMO'],
            ticktext=df_tmo_analista['TMO_formatado']  # Usamos os valores formatados em minutos e segundos
        )
    )

    # Formata o valor do TMO no gráfico (barra)
    fig.update_traces(
        text=df_tmo_analista['TMO_formatado'],  # Exibe TMO formatado nas barras
        textposition='outside',
        hovertemplate='Data = %{x|%d/%m/%Y}<br>TMO = %{text}'  # Exibe o TMO formatado no hover
    )

    return fig