import pandas as pd
import dash
from dash import dcc, html
import plotly.express as px
import plotly.graph_objects as go



# Carregar os dados do arquivo Excel
def load_data(file_path):
    df = pd.read_excel(file_path)
    return df

# Função para criar gráficos de barras
def create_bar_graph(df, x_column, y_column, title):
    fig = px.bar(df, x=x_column, y=y_column, color=x_column, title=title)
    return fig

# Função para criar gráficos de linha
def create_line_graph(df, x_column, y_column, title):
    fig = px.line(df, x=x_column, y=y_column, title=title)
    return fig

# Função para criar gráficos de área
def create_area_graph(df, x_column, y_column, title):
    fig = px.area(df, x=x_column, y=y_column, title=title)
    return fig

# Função para criar gráficos de pizza
def create_pie_chart(df, names_column, values_column, title):
    fig = px.pie(df, names=names_column, values=values_column, title=title)
    fig.update_traces(textinfo='percent+label', pull=[0.1] * len(df[names_column].unique()))  # Destaca as fatias
    return fig

# Função para gerar o layout do dashboard
def serve_layout(df):
    return html.Div([
        dcc.Location(id='url', refresh=False),  # Componente para controlar a navegação

        # Barra lateral
        html.Div([
            html.H2('Navegação', className='sidebar-title'),

            html.Ul([
                html.Li(html.A('Leads Recebidos', href='/leads')),
                html.Li(html.A('Vendas Realizadas', href='/vendas')),
                html.Li(html.A('Tempo Médio de Atendimento', href='/tempo')),
                html.Li(html.A('Ações Realizadas', href='/acoes-realizadas')),
                html.Li(html.A('Ações Planejadas', href='/acoes-planejadas')),
                html.Li(html.A('Resgates de Clientes', href='/resgates')),
                html.Li(html.A('Pesquisa de Satisfação', href='/pesquisa')),
                html.Li(html.A('Meta de Vendas', href='/meta')),
                html.Li(html.A('Insucessos', href='/insucessos')),
                html.Li(html.A('Operador', href='/operador')),
                html.Li(html.A('Unidade', href='/unidade'))
            ], className='sidebar-list'),

        ], className='sidebar'),

        # Conteúdo do dashboard
        html.Div(id='page-content', className='content')
    ])

# Função principal que inicializa o app
def run_dashboard():
    # Caminho do arquivo Excel
    data_file = r'C:\Users\Alan Milan\Documents\automatiz\Base de Dados.xlsx'
    
    # Carregar os dados
    df = load_data(data_file)

    # Inicializar o app Dash
    app = dash.Dash(__name__, suppress_callback_exceptions=True)  # Suprimir exceções de callback

    # Definir o layout do app
    app.layout = serve_layout(df)

    # Callback para atualizar o conteúdo da página dinamicamente com base na URL
    @app.callback(
        dash.dependencies.Output('page-content', 'children'),
        [dash.dependencies.Input('url', 'pathname')]
    )
    def display_page(pathname):
        if pathname == '/leads':
            return generate_graph_layout(df, 'Mês', 'Leads Recebidos', 'Leads Recebidos por Mês')
        elif pathname == '/vendas':
            return generate_graph_layout(df, 'Mês', 'Vendas Realizadas', 'Vendas Realizadas por Mês')
        elif pathname == '/tempo':
            return generate_graph_layout(df, 'Mês', 'Tempo Médio de Atendimento (min)', 'Tempo Médio de Atendimento por Mês')
        elif pathname == '/acoes-realizadas':
            return generate_graph_layout(df, 'Mês', 'Ações Realizadas', 'Ações Realizadas por Mês')
        elif pathname == '/acoes-planejadas':
            return generate_graph_layout(df, 'Mês', 'Ações Planejadas', 'Ações Planejadas por Mês')
        elif pathname == '/resgates':
            return generate_graph_layout(df, 'Mês', 'Resgates de Clientes', 'Resgates de Clientes por Mês')
        elif pathname == '/pesquisa':
            return generate_graph_layout(df, 'Mês', 'Pesquisa de Satisfação', 'Pesquisa de Satisfação por Mês')
        elif pathname == '/meta':
            return generate_graph_layout(df, 'Mês', 'Meta de Vendas', 'Meta de Vendas por Mês')
        elif pathname == '/insucessos':
            return generate_graph_layout(df, 'Mês', 'Insucessos', 'Insucessos por Mês')
        elif pathname == '/operador':
            return generate_operador_layout(df)
        elif pathname == '/unidade':
            return generate_unidade_layout(df)
        else:
            return html.H3('Selecione uma opção na barra lateral.')

    # Função para gerar o layout do gráfico com o dropdown para seleção de tipo de gráfico
    def generate_graph_layout(df, x_column, y_column, title):
        return html.Div([
            html.H3(title),
            dcc.Dropdown(
                id='chart-type-dropdown',
                options=[
                    {'label': 'Gráfico de Barras', 'value': 'bar'},
                    {'label': 'Gráfico de Linha', 'value': 'line'},
                    {'label': 'Gráfico de Área', 'value': 'area'},
                    {'label': 'Gráfico de Pizza', 'value': 'pie'}
                ],
                value='bar',  # Valor inicial padrão
                clearable=False
            ),
            dcc.Graph(id='dynamic-graph')
        ])

    # Função para gerar o layout dos gráficos para a página de "Operador"
    def generate_operador_layout(df):
        return html.Div([
            html.H3('Análise por Operador'),
            dcc.Dropdown(
                id='operador-metric-dropdown',
                options=[
                    {'label': 'Leads Recebidos', 'value': 'Leads Recebidos'},
                    {'label': 'Vendas Realizadas', 'value': 'Vendas Realizadas'},
                    {'label': 'Atendimentos no Dia', 'value': 'Atendimentos no Dia'},
                    {'label': 'Ações Realizadas', 'value': 'Ações Realizadas'},
                    {'label': 'Ações Planejadas', 'value': 'Ações Planejadas'},
                    {'label': 'Resgates de Clientes', 'value': 'Resgates de Clientes'},
                    {'label': 'Pesquisa de Satisfação', 'value': 'Pesquisa de Satisfação'},
                    {'label': 'Meta de Vendas', 'value': 'Meta de Vendas'},
                    {'label': 'Insucessos', 'value': 'Insucessos'},
                    {'label': 'Tempo Médio de Atendimento', 'value': 'Tempo Médio de Atendimento (min)'},
                ],
                value='Leads Recebidos',  # Valor inicial padrão
                clearable=False
            ),
            dcc.Dropdown(
                id='operador-type-dropdown',
                options=[
                    {'label': 'Gráfico de Barras', 'value': 'bar'},
                    {'label': 'Gráfico de Linha', 'value': 'line'},
                    {'label': 'Gráfico de Área', 'value': 'area'},
                    {'label': 'Gráfico de Pizza', 'value': 'pie'}
                ],
                value='bar',  # Valor inicial padrão
                clearable=False
            ),
            dcc.Graph(id='operador-graph')
        ])

    # Função para gerar o layout dos gráficos para a página de "Unidade"
    def generate_unidade_layout(df):
        return html.Div([
            html.H3('Análise por Unidade'),
            dcc.Dropdown(
                id='unidade-metric-dropdown',
                options=[
                    {'label': 'Leads Recebidos', 'value': 'Leads Recebidos'},
                    {'label': 'Vendas Realizadas', 'value': 'Vendas Realizadas'},
                    {'label': 'Atendimentos no Dia', 'value': 'Atendimentos no Dia'},
                    {'label': 'Ações Realizadas', 'value': 'Ações Realizadas'},
                    {'label': 'Ações Planejadas', 'value': 'Ações Planejadas'},
                    {'label': 'Resgates de Clientes', 'value': 'Resgates de Clientes'},
                    {'label': 'Pesquisa de Satisfação', 'value': 'Pesquisa de Satisfação'},
                    {'label': 'Meta de Vendas', 'value': 'Meta de Vendas'},
                    {'label': 'Insucessos', 'value': 'Insucessos'},
                    {'label': 'Tempo Médio de Atendimento', 'value': 'Tempo Médio de Atendimento (min)'},
                ],
                value='Leads Recebidos',  # Valor inicial padrão
                clearable=False
            ),
            dcc.Dropdown(
                id='unidade-type-dropdown',
                options=[
                    {'label': 'Gráfico de Barras', 'value': 'bar'},
                    {'label': 'Gráfico de Linha', 'value': 'line'},
                    {'label': 'Gráfico de Área', 'value': 'area'},
                    {'label': 'Gráfico de Pizza', 'value': 'pie'}
                ],
                value='bar',  # Valor inicial padrão
                clearable=False
            ),
            dcc.Graph(id='unidade-graph')
        ])

    # Callback para atualizar o gráfico com base na seleção de métrica e tipo de gráfico para a página de "Operador"
    @app.callback(
        dash.dependencies.Output('operador-graph', 'figure'),
        [dash.dependencies.Input('operador-metric-dropdown', 'value'),
         dash.dependencies.Input('operador-type-dropdown', 'value')],
        [dash.dependencies.State('url', 'pathname')]
    )
    def update_operador_graph(metric, chart_type, pathname):
        if pathname != '/operador':
            return dash.no_update
        df_operador = df.groupby(['Operador', 'Mês']).agg({metric: 'sum'}).reset_index()
        if chart_type == 'bar':
            return create_bar_graph(df_operador, 'Mês', metric, f'{metric} por Mês e Operador')
        elif chart_type == 'line':
            return create_line_graph(df_operador, 'Mês', metric, f'{metric} por Mês e Operador')
        elif chart_type == 'area':
            return create_area_graph(df_operador, 'Mês', metric, f'{metric} por Mês e Operador')
        elif chart_type == 'pie':
            return create_pie_chart(df_operador, 'Operador', metric, f'{metric} por Operador')

    # Callback para atualizar o gráfico com base na seleção de métrica e tipo de gráfico para a página de "Unidade"
    @app.callback(
        dash.dependencies.Output('unidade-graph', 'figure'),
        [dash.dependencies.Input('unidade-metric-dropdown', 'value'),
         dash.dependencies.Input('unidade-type-dropdown', 'value')],
        [dash.dependencies.State('url', 'pathname')]
    )
    def update_unidade_graph(metric, chart_type, pathname):
        if pathname != '/unidade':
            return dash.no_update
        df_unidade = df.groupby(['Unidade', 'Mês']).agg({metric: 'sum'}).reset_index()
        if chart_type == 'bar':
            return create_bar_graph(df_unidade, 'Mês', metric, f'{metric} por Mês e Unidade')
        elif chart_type == 'line':
            return create_line_graph(df_unidade, 'Mês', metric, f'{metric} por Mês e Unidade')
        elif chart_type == 'area':
            return create_area_graph(df_unidade, 'Mês', metric, f'{metric} por Mês e Unidade')
        elif chart_type == 'pie':
            return create_pie_chart(df_unidade, 'Unidade', metric, f'{metric} por Unidade')

    # Callback para atualizar o gráfico com base no tipo selecionado para páginas de métricas gerais
    @app.callback(
        dash.dependencies.Output('dynamic-graph', 'figure'),
        [dash.dependencies.Input('chart-type-dropdown', 'value')],
        [dash.dependencies.State('url', 'pathname')]
    )
    def update_graph(chart_type, pathname):
        if pathname == '/leads':
            return generate_graph(df, 'Mês', 'Leads Recebidos', 'Leads Recebidos por Mês', chart_type)
        elif pathname == '/vendas':
            return generate_graph(df, 'Mês', 'Vendas Realizadas', 'Vendas Realizadas por Mês', chart_type)
        elif pathname == '/tempo':
            return generate_graph(df, 'Mês', 'Tempo Médio de Atendimento (min)', 'Tempo Médio de Atendimento por Mês', chart_type)
        elif pathname == '/acoes-realizadas':
            return generate_graph(df, 'Mês', 'Ações Realizadas', 'Ações Realizadas por Mês', chart_type)
        elif pathname == '/acoes-planejadas':
            return generate_graph(df, 'Mês', 'Ações Planejadas', 'Ações Planejadas por Mês', chart_type)
        elif pathname == '/resgates':
            return generate_graph(df, 'Mês', 'Resgates de Clientes', 'Resgates de Clientes por Mês', chart_type)
        elif pathname == '/pesquisa':
            return generate_graph(df, 'Mês', 'Pesquisa de Satisfação', 'Pesquisa de Satisfação por Mês', chart_type)
        elif pathname == '/meta':
            return generate_graph(df, 'Mês', 'Meta de Vendas', 'Meta de Vendas por Mês', chart_type)
        elif pathname == '/insucessos':
            return generate_graph(df, 'Mês', 'Insucessos', 'Insucessos por Mês', chart_type)

    # Função para gerar o gráfico com base na seleção de métrica e tipo de gráfico
    def generate_graph(df, x_column, y_column, title, chart_type):
        if chart_type == 'bar':
            return create_bar_graph(df, x_column, y_column, title)
        elif chart_type == 'line':
            return create_line_graph(df, x_column, y_column, title)
        elif chart_type == 'area':
            return create_area_graph(df, x_column, y_column, title)
        elif chart_type == 'pie':
            return create_pie_chart(df, x_column, y_column, title)

    # Executar o app
    app.run_server(debug=True)

if __name__ == '__main__':
    run_dashboard()