# charts.py
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd

def create_abc_chart(abc_data: dict) -> go.Figure:
    if not abc_data: return None
    categories = ['Curva A (80%)', 'Curva B (15%)', 'Curva C (5%)']
    counts = [len(abc_data.get('A', [])), len(abc_data.get('B', [])), len(abc_data.get('C', []))]
    
    fig = px.bar(x=categories, y=counts, title="📦 Curva ABC de Produtos (Qtd de Itens)", color=categories,
                 color_discrete_map={'Curva A (80%)': '#2ca02c', 'Curva B (15%)': '#ff7f0e', 'Curva C (5%)': '#d62728'})
    fig.update_layout(template='plotly_white', xaxis_title="Classificação", yaxis_title="Qtd de Produtos Diferentes")
    return fig

def create_heatmap_chart(heatmap_data: list) -> go.Figure:
    if not heatmap_data: return None
    df = pd.DataFrame(heatmap_data)
    days_order = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo']
    pivot = df.pivot(index='Weekday', columns='Hour', values='Count').reindex(days_order).fillna(0)
    
    fig = px.imshow(pivot, labels=dict(x="Hora do Dia", y="Dia da Semana", color="Vendas"),
                    title="🔥 Mapa de Calor de Vendas (Dias vs Horários)", color_continuous_scale="YlOrRd")
    fig.update_layout(template='plotly_white')
    return fig