import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import requests
import pandas as pd
import numpy as np
import datetime
import dash
from datetime import timedelta
from dash import Dash, dcc, html, Input, Output
import plotly.express as px
from plotly import graph_objs as go
from pandas.tseries.offsets import MonthEnd
import streamlit as st
import io
import openpyxl
from openpyxl.chart import LineChart, Reference



series_codes = {
    'All items': 'CUSR0000SA0',
    'All items less food and energy': 'CUSR0000SA0L1E',
    'Food': 'CUSR0000SAF1',
    'Energy': 'CUSR0000SA0E',
    'Apparel': 'CUSR0000SAA',
    'Education and communication': 'CUSR0000SAE',
    'Other goods and services': 'CUSR0000SAG',
    'Medical care': 'CUSR0000SAM',
    'Recreation': 'CUSR0000SAR',
    'Transportation': 'CUSR0000SAT'
}


def gerar_excel_com_graficos(tabela, categorias):
    # Criar um arquivo Excel na memória
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for categoria in categorias:
            # Filtrar dados por categoria
            df_categoria = tabela[tabela['Category'] == categoria]
            # Escrever os dados da categoria em uma aba
            df_categoria.to_excel(writer, sheet_name=categoria, index=False)

            # Carregar a planilha
            workbook = writer.book
            worksheet = workbook[categoria]

            # Criar um gráfico de linha
            chart = LineChart()
            # Os dados do gráfico são da coluna D (índice 4) e as categorias (datas) são da coluna C (índice 3)
            data = Reference(worksheet, min_col=4, min_row=1, max_col=4, max_row=len(df_categoria)+1)
            cats = Reference(worksheet, min_col=3, min_row=2, max_row=len(df_categoria)+1)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)

            # Adicionar o gráfico à planilha
            worksheet.add_chart(chart, "I2")  # Você pode ajustar a posição conforme necessário

    # Retorna o arquivo Excel
    return output.getvalue()

def calcular_inflacao_acumulada(df, start_date, end_date):
    df_periodo = df[(df['datetime'] >= start_date) & (df['datetime'] <= end_date)]
    if not df_periodo.empty:
        valor_inicial = df_periodo.iloc[0]['value']
        valor_final = df_periodo.iloc[-1]['value']
        inflacao_acumulada = ((valor_final - valor_inicial) / valor_inicial) * 100
        return inflacao_acumulada
    else:
        return None

id_para_categoria = {v: k for k, v in series_codes.items()}

def obter_dados_cpi(series_id, start_year='2013', end_year='2023',token = ''):
    endpoint = 'https://api.bls.gov/publicAPI/v2/timeseries/data/'
    params = {
        'seriesid': [series_id],
        'startyear': start_year,
        'endyear': end_year,
        'registrationkey': token,
    }
    
    response = requests.post(endpoint, json=params)

    if response.status_code == 200 and response.json()['status'] == 'REQUEST_NOT_PROCESSED':
        return ("erro", response.json()['message'][0])
    elif response.status_code == 200 and response.json()['status'] == 'REQUEST_SUCCEEDED':
        dados = response.json()['Results']['series'][0]['data']
        df = pd.DataFrame(dados)
        # adicionando coluna com o nome da série
        df['Category'] = id_para_categoria[series_id]
        # Convertendo para datetime e ordenando
        df['date'] = pd.to_datetime(df['year'].astype(str) + df['period'].str[1:], format='%Y%m')
        df = df.sort_values(by='date', ascending=True)
        # Formatando para a abreviação do mês/ano
        df['datetime'] = pd.to_datetime(df['year'].astype(str) + df['period'].str[1:], format='%Y%m')
        df['date'] = df['date'].dt.strftime('%b/%Y')

        df['value'] = df['value'].astype(float)

        df['var_mensal'] = df['value'].pct_change() * 100  # variação percentual mensal

        # Calcular variação percentual anual
        # Considerando que os dados estão ordenados por data, a variação anual pode ser calculada deslocando os valores por 12 meses (1 ano)
        df['var_anual'] = df['value'].pct_change(periods=12) * 100  # variação percentual anual
        
        
        # Removendo colunas desnecessárias e reordenando
        df = df[['datetime','Category', 'date', 'value', 'var_mensal', 'var_anual']]
        # Convertendo valores para float
        
        return ("sucesso", df)

def main():
    st.title('Relatório de CPI')

    token = st.text_input('Digite o token de acesso à API', value='bdaf72c472424c91b17c066cf91ee8cc')

    todas_opcoes = ['All items', 'All items less food and energy', 'Food', 'Energy', 'Apparel', 'Education and communication', 'Other goods and services', 'Medical care', 'Recreation', 'Transportation']

    selecionados = st.multiselect('Selecione as categorias', todas_opcoes)

    if 'dados_baixados' not in st.session_state:
        st.session_state['dados_baixados'] = False
        st.session_state['tabela_concatenada'] = None

    baixar = st.button('Baixar dados')
    if baixar:
        if not token:
            st.warning('Você precisa digitar um token')
            return
        if not selecionados:
            st.warning('Você precisa selecionar pelo menos uma categoria')
            return

        status_message = st.empty()
        status_message.markdown('<span style="color: blue;">Baixando dados...</span>', unsafe_allow_html=True)

        tabelas_temp = []
        erro = False
        for i in selecionados:
            resultado = obter_dados_cpi(series_codes[i], token=token)
            if resultado[0] == 'erro':
                status_message.markdown('<span style="color: red;">Erro ao baixar dados: o Token digitado está inválido</span>', unsafe_allow_html=True)
                erro = True
                break
            elif resultado[0] == 'sucesso' and not resultado[1].empty:
                tabelas_temp.append(resultado[1])

        if not erro:
            # Concatenando as tabelas em uma única DataFrame
            st.session_state['tabela_concatenada'] = pd.concat(tabelas_temp, axis=0)
            st.session_state['dados_baixados'] = True
            status_message.markdown('<span style="color: green;">Dados baixados com sucesso!</span>', unsafe_allow_html=True)

    if st.session_state['dados_baixados']:
        # Seletor de categorias 

        if st.checkbox('Mostrar tabela de dados'):
            categorias_filtro = st.multiselect('Filtrar por categoria', selecionados, default=selecionados)

            tabela_filtrada = st.session_state['tabela_concatenada']

            # Filtrando a tabela
            if categorias_filtro:
                tabela_filtrada = tabela_filtrada[tabela_filtrada['Category'].isin(categorias_filtro)]
                st.dataframe(tabela_filtrada)
        
        categorias_para_graficos = st.multiselect('Escolha as categorias para os gráficos', selecionados, default=selecionados)

        tabela_para_graficos = st.session_state['tabela_concatenada'].copy()
        tabela_para_graficos['datetime'] = pd.to_datetime(tabela_para_graficos['date'], format='%b/%Y')

        # Seletor de período
        periodo = st.selectbox('Escolha o período', ['1 ano', '2 anos', '5 anos', '10 anos'], index=3)
        last_date = tabela_para_graficos['datetime'].max()
        now = last_date + MonthEnd(0)

        if periodo == '1 ano':
            start_date = now - pd.DateOffset(years=1)
        elif periodo == '2 anos':
            start_date = now - pd.DateOffset(years=2)
        elif periodo == '5 anos':
            start_date = now - pd.DateOffset(years=5)
        elif periodo == '10 anos':
            start_date = now - pd.DateOffset(years=10)

        tabela_para_graficos = st.session_state['tabela_concatenada']
        tabela_para_graficos['datetime'] = pd.to_datetime(tabela_para_graficos['date'], format='%b/%Y')

        # Filtrando a tabela para as categorias e intervalo de datas escolhidos
        tabela_filtrada = tabela_para_graficos[
            (tabela_para_graficos['Category'].isin(categorias_para_graficos)) &
            (tabela_para_graficos['datetime'] >= start_date)
        ]

        # Criando um gráfico com múltiplas linhas
        fig = px.line(tabela_filtrada, x='date', y='value', color='Category', title='Gráfico CPI')
        st.plotly_chart(fig)

        print(start_date, now)
        st.markdown("### Inflação Acumulada por Categoria no Período")
        for categoria in selecionados:
            df_categoria = st.session_state['tabela_concatenada'][st.session_state['tabela_concatenada']['Category'] == categoria]
            inflacao_acumulada_categoria = calcular_inflacao_acumulada(df_categoria, start_date, now)
            if inflacao_acumulada_categoria is not None:
                st.write(f"Inflação acumulada para {categoria}: {inflacao_acumulada_categoria:.2f}%")

        if st.button('Gerar arquivo Excel'):
            # Gerar o arquivo Excel
            excel_data = gerar_excel_com_graficos(st.session_state['tabela_concatenada'], selecionados)
            # Disponibilizar para download
            st.download_button(
                label="Baixar arquivo Excel",
                data=excel_data,
                file_name="dados_cpi.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    

if __name__ == '__main__':
    main()


