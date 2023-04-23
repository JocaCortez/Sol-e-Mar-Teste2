import dash
from dash import html, dcc, dash_table
from dash.dependencies import Input, Output

from datetime import date

import pandas as pd
import numpy as np
import re

import plotly.express as px
import plotly.graph_objects as go
import dash_bootstrap_components as dbc
from plotly.subplots import make_subplots

import locale
import plotly.io as pio
from dash_bootstrap_templates import load_figure_template

def formatar_reais(valor):
    return locale.currency(valor, grouping=True, symbol='R$')


# Define o layout do app
load_figure_template("Pulse")


# Extrair os dados
df = pd.read_excel("extrato_financeiro.xls")
# Converter a data para datatime
df['Data movimento'] = pd.to_datetime(df['Data movimento'], dayfirst=True, format='%d/%m/%Y')

# Define a localização para português do Brasil
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')


app = dash.Dash(external_stylesheets=[dbc.themes.PULSE])
app.layout = html.Div(children=([
    
dbc.Row([
    dbc.Col([
        dbc.Card([
            dbc.CardBody([
                dbc.Row([
                    html.H3("DashBoard Sol e Mar")
                    ]),
                dbc.Row([
                    html.H6("Selecione o Periodo de Analise!"),
                    
                    dcc.DatePickerSingle(
                                        id='data-inicial',
                                        min_date_allowed=min(df['Data movimento']),
                                        max_date_allowed=max(df['Data movimento']),
                                        initial_visible_month=max(df['Data movimento']),
                                        date="04/01/2023",
                                       
                                        display_format='DD/MM/YYYY'),
                    
                    # Componente para selecionar a data final
                                    dcc.DatePickerSingle(
                                        id='data-final',
                                        min_date_allowed=min(df['Data movimento']),
                                        max_date_allowed=max(df['Data movimento']),
                                        initial_visible_month=max(df['Data movimento']),
                                        date="04/30/2023",
                                        
                                        display_format='DD/MM/YYYY'),
                ]),
            ])
            
        ])
        
    ], sm=2),
    
    dbc.Col([
        dbc.Row([
            dbc.Card([
                dbc.CardBody([
                    dbc.Row([
                        dbc.Col([dcc.Graph(id='grafico1')]),
                                ], style={"height": "50h", "margin": "1px", "padding": "20px"},),
                        
                    ])
                    
                ])
            ])
            
        ])
    ]),

dbc.Row([
    dbc.Col([
        dbc.Card([
            dbc.CardBody([
                dbc.Row(
                    dbc.Col([    
                        dash_table.DataTable(
                                id='tabela-dados',
                                columns=[{'id': c, 'name': c} for c in df.columns],
                                page_size=10,
                                style_table={'height': '300px', 'overflowY': 'auto', 'textAlign': 'center'})
                    ])
                )
                
                
            ])
        ])
    ])
]),  

dbc.Row([
    dbc.Col([
        dbc.Card([
            dbc.CardBody([
                dbc.Row([
                    dbc.Col([dcc.Graph(id='grafico4')]),
                    dbc.Col([dcc.Graph(id='grafico5')]),
                    
                ])
                
                    
                    
                ])
            ])
        ])
    ]),

dbc.Row([
    dbc.Col([
        dbc.Card([
            dbc.CardBody([
                dbc.Row([
                    dbc.Col([dcc.Graph(id='grafico2')]),
                    dbc.Col([dcc.Graph(id='grafico6')]),
                    
                ])
            ])
        ])
    ])
    
]),

dbc.Row([
    dbc.Col([
        dbc.Card([
            dbc.CardBody([
                dbc.Row(
                    dbc.Col([dcc.Graph(id='grafico7')])
                )
                
                
            ])
        ])
    ])
]),





]))
    





@app.callback(
    Output('grafico1', 'figure'),
    Output('grafico2', 'figure'),
    Output('grafico5', 'figure'),
    Output('grafico4', 'figure'),
    Output('grafico6', 'figure'),
    Output('grafico7', 'figure'),
    Output('tabela-dados', 'data'),
    [Input('data-inicial', 'date'),
     Input('data-final', 'date')]
)






def atualizar_grafico(data_inicial, data_final):
        
    # Filtra as linhas dentro do período desejado
    df_filtrado = df.loc[(df['Data movimento'] >= data_inicial) & (df['Data movimento'] <= data_final)]
    # Filtra o movimento da Sol e Mar
    df_solemar = df_filtrado.loc[(df['Conta bancária'] == 'Sol e Mar - Sicredi')]
    # Transforma os valores negativos em positivos
    df_solemar['Valor (R$)'] = df_solemar['Valor (R$)'].abs()
    # Remove os movimentos com as descrições 'RESG.APLIC.FIN.AVISO PREV' e 'APLICACAO FINANCEIRA'
    df_solemar = df_solemar[~df_solemar['Descrição'].str.contains('RESG.APLIC.FIN.AVISO PREV|APLICACAO FINANCEIRA')]
    # Procura pelos lancamentos dos lotes na descição
    lotes = ['principe', 'mariluz', 'paese', "príncipe"]
    df_loteamentos = df_solemar[df_solemar['Descrição'].str.contains('|'.join(lotes), case=False)]
    df_loteamentos['Nome do Loteamento'] = df_loteamentos['Descrição'].str.extract(f'({"|".join(lotes)})', flags=re.IGNORECASE)[0]
    # Agrupa os dados por nome do loteamento
    df_solemar_filtered_grouped = df_loteamentos.groupby('Nome do Loteamento').sum().reset_index()
    # Calcula a porcentagem de cada loteamento em relação ao valor total
    total_value = df_solemar_filtered_grouped['Valor (R$)'].sum()
    df_solemar_filtered_grouped['Porcentagem'] = df_solemar_filtered_grouped['Valor (R$)'] / total_value
    # Agrupa as datas de movimento
    df_solemar_groupedM = df_solemar.groupby([pd.Grouper(key='Data movimento', freq='M'), 'Tipo da operação']).sum().reset_index()
    
    
    #====================
    df_solemar_debitos = df_solemar.loc[df_solemar['Tipo da operação'] == 'Débito']

    debitos_agrupados = df_solemar_debitos.groupby('Nome do fornecedor/cliente')
    debitos_por_fornecedor = debitos_agrupados['Valor (R$)'].sum()
    debitos_por_fornecedor = debitos_por_fornecedor.sort_values(ascending=False)[:15]
    #====================
    debitos_agrupados2 = df_solemar_debitos.groupby('Categoria 1')
    debitos_por_categoria = debitos_agrupados2['Valor (R$)'].sum()
    debitos_por_categoria = debitos_por_categoria.sort_values(ascending=False)[:15]
    total_value = df_solemar_debitos['Valor (R$)'].sum()
    #====================
    df_pessoal = df_solemar.loc[df_solemar['Centro de Custo 1'] == 'Pessoal']

    
    
    #====================
    df_cash = df_filtrado.loc[(df['Conta bancária'] == 'Cash')]
    
    df_cash = df_cash.drop(df_cash[df_cash['Situação'] == 'Em aberto'].index)
    df_cash = df_cash.drop(df_cash[df_cash['Situação'] == 'Atrasado'].index)
    
    # transforma os valores negativos em positivos
    df_cash['Valor (R$)'] = df_cash['Valor (R$)'].abs()
    #agrupa as datas de movimento
    df_cash = df_cash.groupby([pd.Grouper(key='Data movimento', freq='M'), 'Tipo da operação']).sum().reset_index()
    
    
    mean_values = df_cash.groupby("Data movimento")["Valor (R$)"].mean()
    
    fig7 = px.bar(df_cash, x="Data movimento", 
                  y="Valor (R$)", 
                  color="Tipo da operação", 
                  barmode='group',
                  title="Conta Cash", 
                  text=[locale.currency(val, grouping=True) for val in df_cash['Valor (R$)']], 
                  text_auto='.2s')
    
    
    fig7.add_trace(go.Scatter(x=mean_values.index, 
                           y=mean_values.values, 
                           mode='lines+markers',
                           name='Média',
                           line_color='navy',
                           ))
    
    #fig7.update_traces(text=[locale.currency(val, grouping=True) for val in df_cash['Valor (R$)']], textposition='auto')

    
    #====================
    
    
    
    # Agrupa os dados por tipo de operação e situação agrupada
    df_situacao = df_solemar
    df_situacao['Situação Agrupada'] = df_situacao['Situação'].replace('Conciliado', 'Quitado')
    tabela = df_situacao.groupby(['Tipo da operação', 'Situação'], as_index=False)['Valor (R$)'].sum()


   
    
    #====================
    
    # Grafico Faturamento 
    fig1 = px.bar(df_solemar_groupedM, x="Data movimento", y="Valor (R$)", color="Tipo da operação", barmode='group',
              title="Faturamento Sol e Mar", text=[locale.currency(val, grouping=True) for val in df_solemar_groupedM['Valor (R$)']],
              text_auto=".3s")
    fig1.update_traces(textfont_size=20, textangle=0, textposition="outside", cliponaxis=False)
    fig1.update_layout(margin=dict(l=30, r=30, t=40, b=10))






    

    
    # Grafico Loteamentos
    fig2 = px.pie(df_solemar_filtered_grouped, values='Valor (R$)', names='Nome do Loteamento', title='Faturamento por Loteamento',
                 hover_data={'Valor (R$)': ':,.2f', 'Porcentagem': ':.2%'}, labels={'Valor (R$)': 'Valor'},
                 custom_data=['Valor (R$)', 'Porcentagem'],)
    fig2.update_traces(textposition='inside', textinfo='label+value')
    fig2.add_annotation(text=f"Valor total: {locale.currency(total_value, grouping=True)}", x=0.5, y=-0.2, showarrow=False) 
    
    # Grafico Gasto por Categoria
    fig4 = px.bar(x=debitos_por_categoria.index, y=debitos_por_categoria.values, color=debitos_por_categoria.index, labels={'x': 'Fornecedor', 'y': 'Valor (R$)'},
             title="Gasto por Categoria", text=[locale.currency(val, grouping=True) for val in debitos_por_categoria],
             text_auto=".5s", height=600)
    fig4.update_traces(textfont_size=10, textangle=0, textposition="outside", cliponaxis=False)
    fig4.update_layout(margin=dict(r=100))
    
    # Grafico Gasos por categoria e fornecedor
    fig5 = px.icicle(df_solemar_debitos, path=['Categoria 1', 'Descrição'],
                 values='Valor (R$)',
                 color='Valor (R$)',
                 hover_data={'Valor (R$)': ':.2f', 'Categoria 1': True},
                 branchvalues='total')
    fig5.update_traces(textfont_size=16, textposition='middle center',
                   marker_line_width=1)
    fig5.update_layout(
        margin=dict(l=50, r=50, t=30, b=30))
    fig5.add_annotation(
        x=0.5,
        y=1.15,
        text=f"Valor Total: R$ {total_value:.2f}",
        showarrow=False,
        font=dict(size=16))
    
    fig6 = px.pie(df_pessoal, values='Valor (R$)', names='Categoria 1', title='Gastos pessoais')
    fig6.update_traces(hovertemplate='%{label}: R$ %{value:.2f}', textposition='inside')
    fig6.add_annotation(text=f"Valor total: {locale.currency(total_value, grouping=True)}", x=0.5, y=-0.2, showarrow=False)
    
    return fig1, fig2, fig5, fig4, fig6, fig7, tabela.to_dict('records')










if __name__ == '__main__':
    app.run_server(debug=True,port=8051)