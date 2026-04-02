# -*- coding: utf-8 -*-
"""
Analisador de Vendas Web
Aplicação web para análise de dados de vendas com geração de relatórios em PDF.
Com suporte a múltiplas formas de pagamento e taxas diferenciadas.
"""

import io
import json
import os
import logging
from datetime import datetime, timedelta
from typing import Dict, Optional, Tuple

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# Configuração de logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Configuração da página Streamlit
st.set_page_config(
    page_title="Analisador de Vendas",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Tema e estilos
st.markdown("""
<style>
    .main {
        padding: 0rem 1rem;
    }
    .metric-container {
        background-color: #f0f2f6;
        padding: 20px;
        border-radius: 10px;
        margin: 10px 0;
    }
    h1 {
        color: #1f77b4;
        text-align: center;
    }
    .section-title {
        color: #1f77b4;
        font-size: 1.3em;
        font-weight: bold;
        margin-top: 20px;
    }
</style>
""", unsafe_allow_html=True)

# Constantes
DEFAULT_STORE_NAME = "OnFight"
DEFAULT_OVERSIZED_COST = 30.00

# Taxas por método de pagamento (em decimal, ex: 0.0499 = 4.99%)
DEFAULT_TAX_RATES = {
    "Cartão de crédito": 4.99,
    "Pix": 0.99,
    "Boleto": 1.49,
    "Outro": 2.99
}

# Arquivo de histórico
HISTORY_FILE = "analise_history.json"

# Categorias padrão (fallback)
DEFAULT_CATEGORIAS_CONFIG = {
    "Oversized": ["oversized"],
    "Short 2 em 1": ["short", "2 em 1", "2em1"],
    "Dryfit": ["dryfit", "dry fit"],
    "Moletom": ["moletom", "hoodie"],
    "Calça": ["calça", "calca", "pants"],
    "Combo": ["combo", "kit"]
}

# Mapeamento de métodos de pagamento
PAYMENT_METHOD_MAPPING = {
    "cartão": "Cartão de crédito",
    "cartao": "Cartão de crédito",
    "credit card": "Cartão de crédito",
    "pix": "Pix",
    "boleto": "Boleto",
    "transferência": "Outro",
    "transferencia": "Outro"
}


class VendasAnalyzerWeb:
    """Analisador de vendas para interface web."""

    def __init__(self, categorias_config: Dict[str, list] = None):
        """Inicializa o analisador."""
        self.categorias_config = categorias_config or DEFAULT_CATEGORIAS_CONFIG.copy()

    def _identify_payment_method(self, payment_text: str) -> str:
        """
        Identifica o método de pagamento baseado no texto.
        
        Args:
            payment_text: Texto do método de pagamento do CSV
            
        Returns:
            Método de pagamento normalizado
        """
        if not payment_text:
            return "Outro"
        
        payment_lower = str(payment_text).lower().strip()
        
        # Procura por correspondências no mapeamento
        for keyword, method in PAYMENT_METHOD_MAPPING.items():
            if keyword in payment_lower:
                return method
        
        # Se não encontrar, retorna "Outro"
        return "Outro"

    def _identify_category(self, product_name: str) -> str:
        """
        Identifica a categoria do produto baseado em palavras-chave.
        
        Args:
            product_name: Nome do produto
            
        Returns:
            Categoria identificada ou 'Outros'
        """
        product_name_lower = str(product_name).lower().strip()
        
        for category, keywords in self.categorias_config.items():
            for keyword in keywords:
                if keyword.lower() in product_name_lower:
                    return category
        
        return "Outros"

    def add_category(self, category_name: str, keywords: list) -> bool:
        """
        Adiciona uma nova categoria.
        
        Args:
            category_name: Nome da categoria
            keywords: Lista de palavras-chave
            
        Returns:
            True se adicionada com sucesso
        """
        if not category_name.strip():
            return False
        
        # Remove espaços extras e converte para lista se necessário
        clean_keywords = [kw.strip().lower() for kw in keywords if kw.strip()]
        
        if not clean_keywords:
            return False
        
        self.categorias_config[category_name.strip()] = clean_keywords
        return True

    def remove_category(self, category_name: str) -> bool:
        """
        Remove uma categoria.
        
        Args:
            category_name: Nome da categoria a remover
            
        Returns:
            True se removida com sucesso
        """
        if category_name in self.categorias_config:
            del self.categorias_config[category_name]
            return True
        return False

    def update_category(self, old_name: str, new_name: str, keywords: list) -> bool:
        """
        Atualiza uma categoria existente.
        
        Args:
            old_name: Nome atual da categoria
            new_name: Novo nome da categoria
            keywords: Nova lista de palavras-chave
            
        Returns:
            True se atualizada com sucesso
        """
        if old_name not in self.categorias_config:
            return False
        
        if old_name != new_name:
            # Remove a antiga e adiciona a nova
            del self.categorias_config[old_name]
        
        return self.add_category(new_name, keywords)

    def get_categories_list(self) -> list:
        """
        Retorna lista de categorias disponíveis (excluindo 'Outros').
        
        Returns:
            Lista de nomes de categorias
        """
        return [cat for cat in self.categorias_config.keys() if cat != "Outros"]

    @staticmethod
    def _parse_float(value: str) -> float:
        """
        Converte string para float.
        
        Args:
            value: String a converter
            
        Returns:
            Valor em float
            
        Raises:
            ValueError: Se não conseguir converter
        """
        try:
            return float(str(value).strip().replace(',', '.'))
        except ValueError:
            raise ValueError(f"Valor inválido: '{value}'")

    def _calculate_category_stats(
        self,
        df: pd.DataFrame,
        costs_map: Dict[str, float],
        default_cost: float
    ) -> Dict[str, Dict]:
        """Calcula estatísticas por categoria."""
        stats = {}
        
        for _, row in df.iterrows():
            product_name = str(row['Lineitem name'])
            quantity = row['Lineitem quantity']
            price = row['Lineitem price']
            
            category = self._identify_category(product_name)
            unit_cost = costs_map.get(category, default_cost)
            
            if category not in stats:
                stats[category] = {"qty": 0, "value": 0, "cost": 0}
            
            stats[category]["qty"] += quantity
            stats[category]["value"] += (price * quantity)
            stats[category]["cost"] += (unit_cost * quantity)
        
        return stats

    def _calculate_taxes_by_payment_method(self, df: pd.DataFrame, tax_rates: Dict[str, float]) -> Dict:
        """
        Calcula taxas por método de pagamento.
        
        Args:
            df: DataFrame com dados de vendas
            tax_rates: Dicionário com taxas por método de pagamento
            
        Returns:
            Dicionário com análise de taxas por método
        """
        unique_paid = df[df['Financial Status'].str.lower() == 'paid'].drop_duplicates(subset=['Name'])
        
        tax_analysis = {
            "by_method": {},
            "total_by_method": {},
            "tax_by_method": {}
        }
        
        # Tenta encontrar coluna de método de pagamento
        payment_col = None
        possible_cols = ['Payment Method', 'Metodo Pagamento', 'Método Pagamento', 'Payment', 'Pagamento']
        
        for col in possible_cols:
            if col in df.columns:
                payment_col = col
                break
        
        if payment_col is None:
            # Se não encontrar coluna, assume uma taxa média
            return {
                "by_method": {},
                "total_by_method": {},
                "tax_by_method": {},
                "warning": "Coluna de método de pagamento não encontrada"
            }
        
        # Agrupa por método de pagamento
        for _, row in unique_paid.iterrows():
            payment_method = self._identify_payment_method(row.get(payment_col, ""))
            order_total = row['Total']
            
            if payment_method not in tax_analysis['by_method']:
                tax_analysis['by_method'][payment_method] = {
                    'count': 0,
                    'total': 0,
                    'tax_rate': tax_rates.get(payment_method, 2.99) / 100
                }
            
            tax_analysis['by_method'][payment_method]['count'] += 1
            tax_analysis['by_method'][payment_method]['total'] += order_total
        
        # Calcula taxas por método
        for method, data in tax_analysis['by_method'].items():
            tax_amount = data['total'] * data['tax_rate']
            tax_analysis['total_by_method'][method] = data['total']
            tax_analysis['tax_by_method'][method] = tax_amount
        
        return tax_analysis

    def _analyze_repeat_customers(self, df: pd.DataFrame) -> Dict:
        """Analisa clientes que repetiram compra."""
        unique_paid = df[df['Financial Status'].str.lower() == 'paid'].drop_duplicates(subset=['Name'])
        
        if 'Email' not in df.columns:
            return {
                "total_unique_customers": len(unique_paid),
                "repeat_customers": 0,
                "repeat_percentage": 0.0,
                "top_customers": []
            }
        
        email_counts = unique_paid['Email'].value_counts()
        repeat_customers = email_counts[email_counts > 1]
        
        # Top 10 clientes
        top_customers = []
        for email, count in email_counts.head(10).items():
            customer_orders = unique_paid[unique_paid['Email'] == email]
            total_spent = customer_orders['Total'].sum()
            top_customers.append({
                'email': email,
                'orders': count,
                'total_spent': total_spent,
                'avg_order': total_spent / count
            })
        
        return {
            "total_unique_customers": len(email_counts),
            "repeat_customers": len(repeat_customers),
            "repeat_percentage": (len(repeat_customers) / len(email_counts) * 100) if len(email_counts) > 0 else 0.0,
            "top_customers": top_customers
        }

    def _analyze_timeline(self, df: pd.DataFrame) -> Dict:
        """Analisa vendas por período."""
        unique_paid = df[df['Financial Status'].str.lower() == 'paid'].drop_duplicates(subset=['Name'])
        
        if 'Created at' not in df.columns:
            return {
                "daily_sales": {},
                "weekly_sales": {},
                "monthly_sales": {},
                "best_day": None,
                "best_week": None
            }
        
        # Converte para datetime
        df_copy = df.copy()
        df_copy['Created at'] = pd.to_datetime(df_copy['Created at'], errors='coerce')
        unique_paid = df_copy[df_copy['Financial Status'].str.lower() == 'paid'].drop_duplicates(subset=['Name'])
        
        # Por dia
        daily = unique_paid.groupby(unique_paid['Created at'].dt.date)['Total'].agg(['sum', 'count'])
        daily_sales = {str(date): {'value': value, 'count': count} for date, value, count in zip(daily.index, daily['sum'], daily['count'])}
        
        # Por semana
        weekly = unique_paid.groupby(unique_paid['Created at'].dt.isocalendar().week)['Total'].agg(['sum', 'count'])
        weekly_sales = {f"Semana {week}": {'value': value, 'count': count} for week, value, count in zip(weekly.index, weekly['sum'], weekly['count'])}
        
        # Por mês
        monthly = unique_paid.groupby(unique_paid['Created at'].dt.to_period('M'))['Total'].agg(['sum', 'count'])
        monthly_sales = {str(month): {'value': value, 'count': count} for month, value, count in zip(monthly.index, monthly['sum'], monthly['count'])}
        
        best_day = max(daily_sales.items(), key=lambda x: x[1]['value']) if daily_sales else None
        best_week = max(weekly_sales.items(), key=lambda x: x[1]['value']) if weekly_sales else None
        
        return {
            "daily_sales": daily_sales,
            "weekly_sales": weekly_sales,
            "monthly_sales": monthly_sales,
            "best_day": best_day,
            "best_week": best_week
        }

    def _analyze_geographic(self, df: pd.DataFrame) -> Dict:
        """Analisa vendas por localização geográfica."""
        unique_paid = df[df['Financial Status'].str.lower() == 'paid'].drop_duplicates(subset=['Name'])
        
        geo_data = {
            "by_state": {},
            "by_city": {},
            "top_states": [],
            "top_cities": []
        }
        
        if 'Shipping Province' not in df.columns and 'Billing Province' not in df.columns:
            return geo_data
        
        # Usar Shipping Province se disponível, senão Billing Province
        province_col = 'Shipping Province' if 'Shipping Province' in df.columns else 'Billing Province'
        city_col = 'Shipping City' if 'Shipping City' in df.columns else 'Billing City'
        
        # Por estado
        if province_col in unique_paid.columns:
            state_sales = unique_paid.groupby(province_col).agg({
                'Total': ['sum', 'count'],
                'Subtotal': 'sum'
            }).reset_index()
            state_sales.columns = ['state', 'total_value', 'order_count', 'subtotal']
            state_sales = state_sales.sort_values('total_value', ascending=False)
            
            geo_data['by_state'] = state_sales.to_dict('records')
            geo_data['top_states'] = state_sales.head(5).values.tolist()
        
        # Por cidade
        if city_col in unique_paid.columns:
            city_sales = unique_paid.groupby(city_col).agg({
                'Total': ['sum', 'count'],
                'Subtotal': 'sum'
            }).reset_index()
            city_sales.columns = ['city', 'total_value', 'order_count', 'subtotal']
            city_sales = city_sales.sort_values('total_value', ascending=False)
            
            geo_data['by_city'] = city_sales.to_dict('records')
            geo_data['top_cities'] = city_sales.head(5).values.tolist()
        
        return geo_data

    def _analyze_fulfillment(self, df: pd.DataFrame) -> Dict:
        """Analisa status de fulfillment."""
        if 'Fulfillment Status' not in df.columns:
            return {
                "total_orders": 0,
                "fulfilled": 0,
                "unfulfilled": 0,
                "partial": 0,
                "cancelled": 0,
                "fulfillment_rate": 0.0,
                "pending_fulfillment": []
            }
        
        unique_orders = df.drop_duplicates(subset=['Name'])
        
        fulfillment_counts = unique_orders['Fulfillment Status'].value_counts()
        
        total = len(unique_orders)
        fulfilled = fulfillment_counts.get('fulfilled', 0)
        unfulfilled = fulfillment_counts.get('unfulfilled', 0)
        partial = fulfillment_counts.get('partially fulfilled', 0)
        cancelled = fulfillment_counts.get('cancelled', 0)
        
        fulfillment_rate = (fulfilled / total * 100) if total > 0 else 0.0
        
        # Pedidos pendentes (unfulfilled com status paid)
        pending = df[(df['Fulfillment Status'] == 'unfulfilled') & 
                    (df['Financial Status'].str.lower() == 'paid')].drop_duplicates(subset=['Name'])
        
        pending_fulfillment = []
        if 'Created at' in df.columns:
            pending['Created at'] = pd.to_datetime(pending['Created at'], errors='coerce')
            pending_fulfillment = pending[['Name', 'Created at', 'Total']].sort_values('Created at').head(10).values.tolist()
        
        return {
            "total_orders": total,
            "fulfilled": fulfilled,
            "unfulfilled": unfulfilled,
            "partial": partial,
            "cancelled": cancelled,
            "fulfillment_rate": fulfillment_rate,
            "pending_fulfillment": pending_fulfillment
        }

    def _analyze_discount_codes(self, df: pd.DataFrame) -> Dict:
        """Analisa uso e efetividade de cupons."""
        if 'Discount Code' not in df.columns:
            return {
                "total_discounts": 0,
                "total_discount_value": 0.0,
                "discount_codes": [],
                "top_codes": []
            }
        
        unique_paid = df[df['Financial Status'].str.lower() == 'paid'].drop_duplicates(subset=['Name'])
        
        # Somente com código de desconto
        with_discount = unique_paid[unique_paid['Discount Code'].notna() & (unique_paid['Discount Code'] != '')]
        
        discount_stats = with_discount.groupby('Discount Code').agg({
            'Discount Amount': ['sum', 'count'],
            'Total': ['sum', 'mean']
        }).reset_index()
        
        discount_stats.columns = ['code', 'total_discount', 'usage_count', 'total_value', 'avg_order_value']
        discount_stats = discount_stats.sort_values('total_discount', ascending=False)
        
        total_discounts = discount_stats['usage_count'].sum()
        total_discount_value = discount_stats['total_discount'].sum()
        
        discount_codes = discount_stats.to_dict('records')
        top_codes = discount_stats.head(5).to_dict('records')
        
        return {
            "total_discounts": int(total_discounts),
            "total_discount_value": total_discount_value,
            "discount_codes": discount_codes,
            "top_codes": top_codes
        }

    def process_data(
        self,
        df: pd.DataFrame,
        store_name: str,
        costs_map: Dict[str, float],
        default_cost: float,
        platform_tax: float,
        ads_cost: float,
        tax_rates: Dict[str, float]
    ) -> Dict:
        """
        Processa os dados e retorna análise.
        
        Args:
            df: DataFrame com dados de vendas
            store_name: Nome da loja
            costs_map: Mapa de custos por categoria
            default_cost: Custo padrão
            platform_tax: Taxa da plataforma em decimal
            ads_cost: Custo com ADS
            tax_rates: Dicionário com taxas por método de pagamento
            
        Returns:
            Dicionário com análise completa
        """
        df = df.copy()
        
        # Normaliza colunas numéricas
        numeric_cols = ['Subtotal', 'Shipping', 'Total', 'Lineitem quantity', 'Lineitem price']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # Filtra por status financeiro
        paid = df[df['Financial Status'].str.lower() == 'paid']
        cancelled = df[df['Financial Status'].str.lower() == 'cancelled']
        pending = df[df['Financial Status'].str.lower() == 'pending']

        # Dados de pedidos únicos
        unique_paid = paid.drop_duplicates(subset=['Name'])
        total_items = unique_paid['Subtotal'].sum()
        total_shipping = unique_paid['Shipping'].sum()
        total_received = unique_paid['Total'].sum()

        # Análise por categoria
        stats = self._calculate_category_stats(paid, costs_map, default_cost)
        
        total_prod_cost = sum(cat['cost'] for cat in stats.values())
        
        # Análise de taxas por método de pagamento
        tax_analysis = self._calculate_taxes_by_payment_method(df, tax_rates)
        
        # Calcula total de taxas (soma das taxas por método + taxa de plataforma)
        total_payment_taxes = sum(tax_analysis['tax_by_method'].values())
        total_taxes = total_payment_taxes + (total_received * platform_tax)
        
        net_profit = total_received - total_taxes - total_prod_cost - ads_cost

        # ===== ANÁLISE DE FRETE E TRANSPORTADORAS =====
        free_shipping_orders = unique_paid[unique_paid['Shipping'] == 0]
        free_shipping_count = len(free_shipping_orders)
        free_shipping_value = free_shipping_orders['Total'].sum()
        
        # Análise de transportadora
        courier_col = None
        transpose_cols = ['Shipping Method', 'Shipping Name', 'Fulfillment Method', 'Carrier', 'Transportadora']
        
        for col in transpose_cols:
            if col in df.columns:
                courier_col = col
                break
        
        # Contagem de pedidos por transportadora
        correios_pac = 0
        correios_sedex = 0
        transportadoras = 0
        other_couriers = 0
        
        if courier_col is not None:
            for _, order in unique_paid.iterrows():
                shipping_method = str(order[courier_col]).strip().lower() if pd.notna(order[courier_col]) else ""
                
                if 'pac' in shipping_method:
                    correios_pac += 1
                elif 'sedex' in shipping_method:
                    correios_sedex += 1
                elif shipping_method and shipping_method != 'nan' and order['Shipping'] > 0:
                    transportadoras += 1

        # ===== NOVAS ANÁLISES =====
        repeat_customers = self._analyze_repeat_customers(df)
        timeline = self._analyze_timeline(df)
        geographic = self._analyze_geographic(df)
        fulfillment = self._analyze_fulfillment(df)
        discounts = self._analyze_discount_codes(df)

        # ===== PERÍODO DAS VENDAS =====
        sales_period = {"start_date": None, "end_date": None}
        if 'Created at' in df.columns:
            df_dates = df.copy()
            df_dates['Created at'] = pd.to_datetime(df_dates['Created at'], errors='coerce')
            paid_dates = df_dates[df_dates['Financial Status'].str.lower() == 'paid']['Created at'].dropna()
            if len(paid_dates) > 0:
                sales_period["start_date"] = paid_dates.min().strftime('%d/%m/%Y')
                sales_period["end_date"] = paid_dates.max().strftime('%d/%m/%Y')

        return {
            "store_name": store_name,
            "paid_count": len(unique_paid),
            "cancelled_count": len(cancelled.drop_duplicates(subset=['Name'])),
            "pending_count": len(pending.drop_duplicates(subset=['Name'])),
            "total_items": total_items,
            "total_shipping": total_shipping,
            "total_received": total_received,
            "stats": stats,
            "tax_analysis": tax_analysis,
            "platform_tax": platform_tax,
            "total_payment_taxes": total_payment_taxes,
            "total_taxes": total_taxes,
            "total_prod_cost": total_prod_cost,
            "ads_cost": ads_cost,
            "net_profit": net_profit,
            "analysis_date": datetime.now(),
            "free_shipping_count": free_shipping_count,
            "free_shipping_value": free_shipping_value,
            "correios_pac": correios_pac,
            "correios_sedex": correios_sedex,
            "transportadoras": transportadoras,
            "other_couriers": other_couriers,
            "repeat_customers": repeat_customers,
            "timeline": timeline,
            "geographic": geographic,
            "fulfillment": fulfillment,
            "discounts": discounts,
            "sales_period": sales_period
        }

    def _calculate_roi_by_coupon(self, df: pd.DataFrame) -> Dict:
        """Calcula ROI para cada cupom."""
        if 'Discount Code' not in df.columns:
            return {}
        
        unique_paid = df[df['Financial Status'].str.lower() == 'paid'].drop_duplicates(subset=['Name'])
        
        # Média de pedidos sem cupom
        without_discount = unique_paid[unique_paid['Discount Code'].isna() | (unique_paid['Discount Code'] == '')]
        avg_ticket_without = without_discount['Total'].mean() if len(without_discount) > 0 else 0
        
        roi_data = {}
        
        with_discount = unique_paid[unique_paid['Discount Code'].notna() & (unique_paid['Discount Code'] != '')]
        
        for code in with_discount['Discount Code'].unique():
            code_orders = with_discount[with_discount['Discount Code'] == code]
            
            total_discount = code_orders['Discount Amount'].sum()
            total_revenue = code_orders['Total'].sum()
            order_count = len(code_orders)
            avg_ticket_with = code_orders['Total'].mean()
            
            # ROI = (Receita - Desconto) / Desconto
            roi = ((total_revenue - total_discount) / total_discount * 100) if total_discount > 0 else 0
            
            # Aumento de ticket
            ticket_increase = ((avg_ticket_with - avg_ticket_without) / avg_ticket_without * 100) if avg_ticket_without > 0 else 0
            
            roi_data[str(code)] = {
                'code': code,
                'orders': order_count,
                'total_discount': total_discount,
                'total_revenue': total_revenue,
                'roi': roi,
                'avg_ticket': avg_ticket_with,
                'ticket_increase': ticket_increase
            }
        
        return roi_data

    def _compare_periods(self, df: pd.DataFrame, date_from1: str, date_to1: str, date_from2: str, date_to2: str) -> Dict:
        """Compara dois períodos."""
        df_copy = df.copy()
        df_copy['Created at'] = pd.to_datetime(df_copy['Created at'], errors='coerce')
        
        unique_paid = df_copy[df_copy['Financial Status'].str.lower() == 'paid'].drop_duplicates(subset=['Name'])
        
        # Período 1
        p1 = unique_paid[(unique_paid['Created at'] >= date_from1) & (unique_paid['Created at'] <= date_to1)]
        
        # Período 2
        p2 = unique_paid[(unique_paid['Created at'] >= date_from2) & (unique_paid['Created at'] <= date_to2)]
        
        return {
            'period1': {
                'orders': len(p1),
                'revenue': p1['Total'].sum(),
                'avg_ticket': p1['Total'].mean(),
                'items': p1['Subtotal'].sum()
            },
            'period2': {
                'orders': len(p2),
                'revenue': p2['Total'].sum(),
                'avg_ticket': p2['Total'].mean(),
                'items': p2['Subtotal'].sum()
            },
            'growth': {
                'orders_pct': ((len(p2) - len(p1)) / len(p1) * 100) if len(p1) > 0 else 0,
                'revenue_pct': ((p2['Total'].sum() - p1['Total'].sum()) / p1['Total'].sum() * 100) if p1['Total'].sum() > 0 else 0,
                'ticket_pct': ((p2['Total'].mean() - p1['Total'].mean()) / p1['Total'].mean() * 100) if p1['Total'].mean() > 0 else 0
            }
        }

    def generate_excel_report(self, analysis_data: Dict) -> bytes:
        """Gera relatório em Excel."""
        output = io.BytesIO()
        wb = Workbook()
        
        # Estilos
        header_fill = PatternFill(start_color="1F77B4", end_color="1F77B4", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        title_font = Font(bold=True, size=14)
        currency_format = 'R$ #,##0.00'
        
        # ===== RESUMO =====
        ws = wb.active
        ws.title = "Resumo"
        
        ws['A1'] = "RESUMO EXECUTIVO"
        ws['A1'].font = title_font
        
        # Período das vendas
        if analysis_data.get('sales_period') and analysis_data['sales_period'].get('start_date'):
            ws['A2'] = f"Período: {analysis_data['sales_period']['start_date']} até {analysis_data['sales_period']['end_date']}"
        
        row = 4
        metrics = [
            ("Pedidos Pagos", analysis_data['paid_count']),
            ("Pedidos Cancelados", analysis_data['cancelled_count']),
            ("Pedidos Pendentes", analysis_data['pending_count']),
            ("Subtotal", analysis_data['total_items']),
            ("Frete Total", analysis_data['total_shipping']),
            ("Total Recebido", analysis_data['total_received']),
            ("Frete Grátis", analysis_data['free_shipping_count']),
            ("Taxas Totais", analysis_data['total_taxes']),
            ("Lucro Líquido", analysis_data['net_profit']),
        ]
        
        for label, value in metrics:
            ws[f'A{row}'] = label
            ws[f'B{row}'] = value
            row += 1
        
        # ===== TAXAS POR MÉTODO DE PAGAMENTO =====
        ws = wb.create_sheet("Taxas por Pagamento")
        
        ws['A1'] = "MÉTODO DE PAGAMENTO"
        ws['B1'] = "TOTAL VENDAS (R$)"
        ws['C1'] = "TAXA (R$)"
        
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
        
        row = 2
        for method, total in analysis_data['tax_analysis']['total_by_method'].items():
            tax_amount = analysis_data['tax_analysis']['tax_by_method'].get(method, 0)
            ws[f'A{row}'] = method
            ws[f'B{row}'] = total
            ws[f'C{row}'] = tax_amount
            row += 1
        
        # ===== CATEGORIAS =====
        ws = wb.create_sheet("Categorias")
        
        ws['A1'] = "CATEGORIA"
        ws['B1'] = "QUANTIDADE"
        ws['C1'] = "VENDA (R$)"
        ws['D1'] = "CUSTO (R$)"
        ws['E1'] = "MARGEM (R$)"
        
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
        
        row = 2
        for cat, data in sorted(analysis_data['stats'].items(), key=lambda x: x[1]['value'], reverse=True):
            ws[f'A{row}'] = cat
            ws[f'B{row}'] = int(data['qty'])
            ws[f'C{row}'] = data['value']
            ws[f'D{row}'] = data['cost']
            ws[f'E{row}'] = data['value'] - data['cost']
            row += 1
        
        # ===== TOP CLIENTES =====
        ws = wb.create_sheet("Top Clientes")
        
        ws['A1'] = "EMAIL"
        ws['B1'] = "PEDIDOS"
        ws['C1'] = "GASTO TOTAL (R$)"
        ws['D1'] = "TICKET MÉDIO (R$)"
        
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
        
        row = 2
        for customer in analysis_data['repeat_customers'].get('top_customers', []):
            ws[f'A{row}'] = customer['email']
            ws[f'B{row}'] = customer['orders']
            ws[f'C{row}'] = customer['total_spent']
            ws[f'D{row}'] = customer['avg_order']
            row += 1
        
        # ===== GEOGRÁFICA =====
        ws = wb.create_sheet("Geográfica")
        
        ws['A1'] = "ESTADO"
        ws['B1'] = "TOTAL (R$)"
        ws['C1'] = "PEDIDOS"
        
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
        
        row = 2
        for state in analysis_data['geographic'].get('top_states', []):
            ws[f'A{row}'] = state[0]
            ws[f'B{row}'] = state[1]
            ws[f'C{row}'] = int(state[2])
            row += 1
        
        # ===== CUPONS =====
        if analysis_data['discounts'].get('top_codes'):
            ws = wb.create_sheet("Cupons")
            
            ws['A1'] = "CUPOM"
            ws['B1'] = "USOS"
            ws['C1'] = "DESC. TOTAL (R$)"
            ws['D1'] = "TICKET MÉD (R$)"
            
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
            
            row = 2
            for code in analysis_data['discounts']['top_codes']:
                ws[f'A{row}'] = code['code']
                ws[f'B{row}'] = int(code['usage_count'])
                ws[f'C{row}'] = code['total_discount']
                ws[f'D{row}'] = code['avg_order_value']
                row += 1
        
        wb.save(output)
        output.seek(0)
        return output.getvalue()

    @staticmethod
    def generate_pdf(analysis_data: Dict) -> bytes:
        """
        Gera PDF com a análise.
        
        Args:
            analysis_data: Dicionário com dados da análise
            
        Returns:
            Bytes do PDF
        """
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=letter)
        elements = []
        styles = getSampleStyleSheet()
        
        # Estilo customizado
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=16,
            textColor=colors.HexColor('#1f77b4'),
            spaceAfter=30,
            alignment=1  # Center
        )
        
        heading_style = ParagraphStyle(
            'CustomHeading',
            parent=styles['Heading2'],
            fontSize=12,
            textColor=colors.HexColor('#1f77b4'),
            spaceAfter=12,
            spaceBefore=12
        )
        
        # Título
        title = f"Relatório de Vendas - {analysis_data['store_name'].upper()}"
        elements.append(Paragraph(title, title_style))
        elements.append(Spacer(1, 0.2*inch))
        
        # Data
        date_text = f"Data: {analysis_data['analysis_date'].strftime('%d/%m/%Y %H:%M')}"
        elements.append(Paragraph(date_text, styles['Normal']))
        
        # Período das vendas
        if analysis_data.get('sales_period') and analysis_data['sales_period'].get('start_date'):
            period_text = f"Período das Vendas: {analysis_data['sales_period']['start_date']} até {analysis_data['sales_period']['end_date']}"
            elements.append(Paragraph(period_text, styles['Normal']))
        
        elements.append(Spacer(1, 0.3*inch))
        
        # Volume de Pedidos
        elements.append(Paragraph("VOLUME DE PEDIDOS", heading_style))
        data = [
            ['Pagos', str(analysis_data['paid_count'])],
            ['Cancelados', str(analysis_data['cancelled_count'])],
            ['Pendentes', str(analysis_data['pending_count'])]
        ]
        table = Table(data, colWidths=[3*inch, 2*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.beige),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        elements.append(table)
        elements.append(Spacer(1, 0.3*inch))
        
        # Valores Totais
        elements.append(Paragraph("VALORES TOTAIS", heading_style))
        data = [
            ['Subtotal (Produtos)', f"R$ {analysis_data['total_items']:,.2f}"],
            ['Frete Total', f"R$ {analysis_data['total_shipping']:,.2f}"],
            ['Total Recebido', f"R$ {analysis_data['total_received']:,.2f}"]
        ]
        table = Table(data, colWidths=[3*inch, 2*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.lightblue),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        elements.append(table)
        elements.append(Spacer(1, 0.3*inch))
        
        # Taxas por Método de Pagamento
        elements.append(Paragraph("TAXAS POR MÉTODO DE PAGAMENTO", heading_style))
        tax_data = [['Método', 'Total Vendas (R$)', 'Taxa (R$)']]
        
        for method, total in analysis_data['tax_analysis']['total_by_method'].items():
            tax_amount = analysis_data['tax_analysis']['tax_by_method'].get(method, 0)
            tax_data.append([method, f"{total:,.2f}", f"{tax_amount:,.2f}"])
        
        table = Table(tax_data, colWidths=[2*inch, 1.5*inch, 1.5*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        elements.append(table)
        elements.append(Spacer(1, 0.3*inch))
        
        # Análise de Frete e Transportadoras
        elements.append(Paragraph("ANÁLISE DE FRETE E TRANSPORTADORAS", heading_style))
        data = [
            ['Frete Grátis (Qty)', str(analysis_data.get('free_shipping_count', 0))],
            ['Frete Grátis (Valor)', f"R$ {analysis_data.get('free_shipping_value', 0):,.2f}"],
            ['Pedidos - Correios PAC', str(analysis_data.get('correios_pac', 0))],
            ['Pedidos - Correios SEDEX', str(analysis_data.get('correios_sedex', 0))],
            ['Pedidos - Transportadoras', str(analysis_data.get('transportadoras', 0))],
            ['Pedidos - Outros', str(analysis_data.get('other_couriers', 0))]
        ]
        table = Table(data, colWidths=[3*inch, 2*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.lightyellow),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        elements.append(table)
        elements.append(Spacer(1, 0.3*inch))
        
        # ===== NOVAS ANÁLISES AVANÇADAS NO PDF =====
        
        # Clientes
        repeat_customers = analysis_data.get('repeat_customers', {})
        if repeat_customers:
            elements.append(Paragraph("ANÁLISE DE CLIENTES", heading_style))
            data = [
                ['Clientes Únicos', str(repeat_customers.get('total_unique_customers', 0))],
                ['Clientes Repeat', str(repeat_customers.get('repeat_customers', 0))],
                ['% Repeat Customers', f"{repeat_customers.get('repeat_percentage', 0):.1f}%"]
            ]
            table = Table(data, colWidths=[3*inch, 2*inch])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), colors.lightcyan),
                ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            elements.append(table)
            elements.append(Spacer(1, 0.3*inch))
        
        # Fulfillment
        fulfillment = analysis_data.get('fulfillment', {})
        if fulfillment:
            elements.append(Paragraph("STATUS DE FULFILLMENT", heading_style))
            data = [
                ['Total de Pedidos', str(fulfillment.get('total_orders', 0))],
                ['Entregues', str(fulfillment.get('fulfilled', 0))],
                ['Não Entregues', str(fulfillment.get('unfulfilled', 0))],
                ['Taxa de Entrega', f"{fulfillment.get('fulfillment_rate', 0):.1f}%"]
            ]
            table = Table(data, colWidths=[3*inch, 2*inch])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), colors.lightgrey),
                ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            elements.append(table)
            elements.append(Spacer(1, 0.3*inch))
        
        # Cupons
        discounts = analysis_data.get('discounts', {})
        if discounts and discounts.get('total_discounts', 0) > 0:
            elements.append(Paragraph("ANÁLISE DE CUPONS", heading_style))
            data = [
                ['Cupons Utilizados', str(discounts.get('total_discounts', 0))],
                ['Desconto Total', f"R$ {discounts.get('total_discount_value', 0):,.2f}"],
                ['Desconto Médio', f"R$ {discounts.get('total_discount_value', 0) / discounts.get('total_discounts', 1):,.2f}"]
            ]
            table = Table(data, colWidths=[3*inch, 2*inch])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), colors.lightyellow),
                ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            elements.append(table)
            elements.append(Spacer(1, 0.3*inch))
        
        # Categorias
        elements.append(Paragraph("DETALHAMENTO POR CATEGORIA", heading_style))
        sorted_stats = sorted(
            analysis_data['stats'].items(),
            key=lambda x: x[1]['value'],
            reverse=True
        )
        data = [['Categoria', 'Qtd', 'Venda (R$)', 'Custo (R$)']]
        for category, cat_data in sorted_stats:
            data.append([
                category,
                str(int(cat_data['qty'])),
                f"{cat_data['value']:,.2f}",
                f"{cat_data['cost']:,.2f}"
            ])
        
        table = Table(data, colWidths=[2*inch, 1*inch, 1.5*inch, 1.5*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        elements.append(table)
        elements.append(Spacer(1, 0.3*inch))
        
        # Resumo Financeiro
        elements.append(Paragraph("RESUMO FINANCEIRO", heading_style))
        data = [
            [f"(-) Taxas de Pagamento", f"R$ {analysis_data['total_payment_taxes']:,.2f}"],
            [f"(-) Taxa Plataforma ({analysis_data['platform_tax']*100:.2f}%)", 
             f"R$ {analysis_data['total_received'] * analysis_data['platform_tax']:,.2f}"],
            ['(-) Custo Produção', f"R$ {analysis_data['total_prod_cost']:,.2f}"],
            ['(-) Gasto ADS', f"R$ {analysis_data['ads_cost']:,.2f}"],
            ['(=) LUCRO LÍQUIDO', f"R$ {analysis_data['net_profit']:,.2f}"]
        ]
        table = Table(data, colWidths=[3*inch, 2*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -2), colors.lightcyan),
            ('BACKGROUND', (0, -1), (-1, -1), colors.lightgreen),
            ('TEXTCOLOR', (0, -1), (-1, -1), colors.darkgreen),
            ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, -1), (-1, -1), 11),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        elements.append(table)
        
        doc.build(elements)
        buffer.seek(0)
        return buffer.getvalue()


def save_analysis_history(analysis_data: Dict) -> None:
    """Salva análise no histórico."""
    history = []
    
    # Carrega histórico existente
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                history = json.load(f)
        except:
            history = []
    
    # Adiciona nova análise
    analysis_entry = {
        'timestamp': datetime.now().isoformat(),
        'store': analysis_data['store_name'],
        'paid_orders': analysis_data['paid_count'],
        'revenue': float(analysis_data['total_received']),
        'profit': float(analysis_data['net_profit']),
        'shipping': float(analysis_data['total_shipping'])
    }
    
    history.append(analysis_entry)
    
    # Mantém apenas últimas 100 análises
    if len(history) > 100:
        history = history[-100:]
    
    # Salva história
    with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
        json.dump(history, f, indent=2, ensure_ascii=False)


def load_analysis_history() -> list:
    """Carrega histórico de análises."""
    if not os.path.exists(HISTORY_FILE):
        return []
    
    try:
        with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return []


def create_sales_chart(timeline_data: Dict) -> go.Figure:
    """Cria gráfico de vendas ao longo do tempo."""
    if not timeline_data['daily_sales']:
        return None
    
    dates = []
    values = []
    
    for date, data in sorted(timeline_data['daily_sales'].items()):
        dates.append(date)
        values.append(data['value'])
    
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=dates, 
        y=values,
        mode='lines+markers',
        name='Vendas Diárias',
        line=dict(color='#1f77b4', width=2),
        marker=dict(size=6)
    ))
    
    fig.update_layout(
        title="📈 Vendas Diárias",
        xaxis_title="Data",
        yaxis_title="Valor (R$)",
        hovermode='x unified',
        template='plotly_white'
    )
    
    return fig


def create_category_chart(stats: Dict) -> go.Figure:
    """Cria gráfico de vendas por categoria."""
    categories = []
    values = []
    
    for cat, data in sorted(stats.items(), key=lambda x: x[1]['value'], reverse=True):
        categories.append(cat)
        values.append(data['value'])
    
    fig = px.bar(
        x=categories,
        y=values,
        labels={'x': 'Categoria', 'y': 'Vendas (R$)'},
        title="📊 Vendas por Categoria",
        color=values,
        color_continuous_scale='Blues'
    )
    
    fig.update_layout(template='plotly_white')
    
    return fig


def create_geographic_chart(geo_data: Dict) -> go.Figure:
    """Cria gráfico geográfico."""
    if not geo_data['top_states']:
        return None
    
    states = []
    values = []
    
    for state in geo_data['top_states']:
        states.append(state[0])
        values.append(state[1])
    
    fig = px.pie(
        values=values,
        names=states,
        title="🗺️ Distribuição de Vendas por Estado"
    )
    
    fig.update_layout(template='plotly_white')
    
    return fig


def create_fulfillment_chart(fulfillment: Dict) -> go.Figure:
    """Cria gráfico de fulfillment."""
    labels = ['Entregues', 'Não Entregues']
    sizes = [fulfillment['fulfilled'], fulfillment['unfulfilled']]
    colors = ['#4CAF50', '#FF9800']
    
    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=sizes,
        marker=dict(colors=colors)
    )])
    
    fig.update_layout(
        title="📦 Status de Fulfillment",
        template='plotly_white'
    )
    
    return fig


def create_coupon_chart(discounts: Dict) -> go.Figure:
    """Cria gráfico de cupons."""
    if not discounts['top_codes']:
        return None
    
    codes = []
    uses = []
    
    for code in discounts['top_codes']:
        codes.append(code['code'])
        uses.append(code['usage_count'])
    
    fig = px.bar(
        x=codes,
        y=uses,
        labels={'x': 'Cupom', 'y': 'Utilizações'},
        title="💳 Top Cupons",
        color=uses,
        color_continuous_scale='Greens'
    )
    
    fig.update_layout(template='plotly_white')
    
    return fig


def create_payment_method_chart(tax_analysis: Dict) -> go.Figure:
    """Cria gráfico de métodos de pagamento."""
    if not tax_analysis['total_by_method']:
        return None
    
    methods = []
    values = []
    
    for method, total in tax_analysis['total_by_method'].items():
        methods.append(method)
        values.append(total)
    
    fig = px.pie(
        values=values,
        names=methods,
        title="💳 Distribuição de Vendas por Método de Pagamento"
    )
    
    fig.update_layout(template='plotly_white')
    
    return fig


def main():
    """Função principal da aplicação web."""
    st.title("📊 Analisador de Vendas Multi-Loja")
    
    # Inicializar categorias na session_state se não existir
    if 'custom_categories' not in st.session_state:
        st.session_state.custom_categories = DEFAULT_CATEGORIAS_CONFIG.copy()
    
    # Criar instância do analyzer com categorias customizadas
    analyzer = VendasAnalyzerWeb(st.session_state.custom_categories)
    
    # Layout com abas
    tab1, tab2, tab3, tab4 = st.tabs(["📈 Análise Básica", "🔍 Análises Avançadas", "⚙️ Categorias", "ℹ️ Ajuda"])
    
    with tab1:
        # Seção 1: Informações Básicas
        st.markdown("### 1️⃣ Informações Básicas")
        col1, col2 = st.columns([2, 1])
        
        with col1:
            store_name = st.text_input(
                "Nome da Loja",
                value=DEFAULT_STORE_NAME,
                help="Nome da loja para o relatório"
            )
        
        # Seção 2: Upload do CSV
        st.markdown("### 2️⃣ Seleção de Arquivo")
        uploaded_file = st.file_uploader(
            "Selecione seu arquivo CSV de vendas",
            type=['csv'],
            help="Arquivo exportado da Shopify ou plataforma de vendas"
        )
        
        if uploaded_file is not None:
            try:
                # Tenta diferentes encodings
                for encoding in ['utf-8', 'latin-1', 'iso-8859-1']:
                    try:
                        df = pd.read_csv(uploaded_file, encoding=encoding)
                        st.success(f"✅ Arquivo carregado com sucesso ({len(df)} linhas)")
                        break
                    except UnicodeDecodeError:
                        continue
                
                # Seção 3: Custos por Categoria
                st.markdown("### 3️⃣ Custos de Produção (R$)")
                
                col1, col2, col3 = st.columns(3)
                costs = {}
                
                categories = analyzer.get_categories_list()
                
                with col1:
                    for cat in categories[:2]:
                        default = DEFAULT_OVERSIZED_COST if cat == "Oversized" else 0.0
                        costs[cat] = st.number_input(
                            cat,
                            value=float(default),
                            min_value=0.0,
                            step=0.01,
                            key=f"cost_{cat}"
                        )
                
                with col2:
                    for cat in categories[2:4]:
                        costs[cat] = st.number_input(
                            cat,
                            value=0.0,
                            min_value=0.0,
                            step=0.01,
                            key=f"cost_{cat}"
                        )
                
                with col3:
                    for cat in categories[4:]:
                        costs[cat] = st.number_input(
                            cat,
                            value=0.0,
                            min_value=0.0,
                            step=0.01,
                            key=f"cost_{cat}"
                        )
                    
                    default_cost = st.number_input(
                        "Outros (padrão)",
                        value=0.0,
                        min_value=0.0,
                        step=0.01
                    )
                
                # Seção 4: Taxas por Método de Pagamento
                st.markdown("### 4️⃣ Taxas por Método de Pagamento (%)")
                
                col1, col2, col3 = st.columns(3)
                
                tax_rates = DEFAULT_TAX_RATES.copy()
                
                with col1:
                    tax_rates["Cartão de crédito"] = st.number_input(
                        "% Taxa Cartão",
                        value=DEFAULT_TAX_RATES["Cartão de crédito"],
                        min_value=0.0,
                        max_value=100.0,
                        step=0.01
                    )
                
                with col2:
                    tax_rates["Pix"] = st.number_input(
                        "% Taxa Pix",
                        value=DEFAULT_TAX_RATES["Pix"],
                        min_value=0.0,
                        max_value=100.0,
                        step=0.01
                    )
                
                with col3:
                    tax_rates["Boleto"] = st.number_input(
                        "% Taxa Boleto",
                        value=DEFAULT_TAX_RATES["Boleto"],
                        min_value=0.0,
                        max_value=100.0,
                        step=0.01
                    )
                
                # Seção 5: Taxas Adicionais e Marketing
                st.markdown("### 5️⃣ Taxas Adicionais e Marketing")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    platform_tax = st.number_input(
                        "% Taxa Plataforma",
                        value=0.0,
                        min_value=0.0,
                        max_value=100.0,
                        step=0.01
                    )
                
                with col2:
                    ads_cost = st.number_input(
                        "Gasto Total com ADS (R$)",
                        value=0.0,
                        min_value=0.0,
                        step=0.01
                    )
                
                # Botão de Análise
                if st.button("🚀 GERAR ANÁLISE COMPLETA", use_container_width=True):
                    try:
                        analysis_data = analyzer.process_data(
                            df=df,
                            store_name=store_name,
                            costs_map=costs,
                            default_cost=default_cost,
                            platform_tax=platform_tax/100,
                            ads_cost=ads_cost,
                            tax_rates=tax_rates
                        )
                        
                        # Salva na sessão
                        st.session_state.analysis_data = analysis_data
                        st.session_state.df_analysis = df  # Salva DataFrame para comparações
                        st.session_state.show_results = True
                        
                        # Salva no histórico
                        save_analysis_history(analysis_data)
                        
                        st.success("✅ Análise realizada com sucesso!")
                        
                    except Exception as e:
                        st.error(f"❌ Erro na análise: {str(e)}")
                
                # Exibir resultados se existirem
                if st.session_state.get('show_results'):
                    analysis = st.session_state.analysis_data
                    
                    st.markdown("---")
                    st.markdown("## 📊 Resultado da Análise")
                    
                    # Período das vendas
                    if analysis.get('sales_period') and analysis['sales_period'].get('start_date'):
                        period_text = f"📅 **Período das Vendas:** {analysis['sales_period']['start_date']} até {analysis['sales_period']['end_date']}"
                        st.info(period_text)
                    
                    # Métricas principais
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.metric("Pedidos Pagos", analysis['paid_count'])
                    
                    with col2:
                        st.metric("Pedidos Cancelados", analysis['cancelled_count'])
                    
                    with col3:
                        st.metric("Pedidos Pendentes", analysis['pending_count'])
                    
                    with col4:
                        st.metric(
                            "Lucro Líquido",
                            f"R$ {analysis['net_profit']:,.2f}",
                            delta=f"{(analysis['net_profit']/analysis['total_received']*100):.1f}%" if analysis['total_received'] > 0 else "0%"
                        )
                    
                    # Valores Totais
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.metric("Subtotal", f"R$ {analysis['total_items']:,.2f}")
                    
                    with col2:
                        st.metric("Frete", f"R$ {analysis['total_shipping']:,.2f}")
                    
                    with col3:
                        st.metric("Total Recebido", f"R$ {analysis['total_received']:,.2f}")
                    
                    # Análise de Métodos de Pagamento
                    st.markdown("### 💳 Análise de Métodos de Pagamento")
                    
                    if analysis['tax_analysis']['total_by_method']:
                        tax_cols = st.columns(len(analysis['tax_analysis']['total_by_method']))
                        
                        for idx, (method, total) in enumerate(analysis['tax_analysis']['total_by_method'].items()):
                            tax_amount = analysis['tax_analysis']['tax_by_method'].get(method, 0)
                            with tax_cols[idx]:
                                st.metric(
                                    method,
                                    f"R$ {total:,.2f}",
                                    f"Taxa: R$ {tax_amount:,.2f}"
                                )
                    else:
                        st.info("⏳ Coluna de método de pagamento não encontrada. As taxas foram calculadas como média.")
                    
                    # Análise de Frete e Transportadoras
                    st.markdown("### 📦 Análise de Frete e Transportadoras")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.metric(
                            "Frete Grátis",
                            f"{analysis['free_shipping_count']} pedidos",
                            f"R$ {analysis['free_shipping_value']:,.2f}"
                        )
                    
                    with col2:
                        st.metric(
                            "Correios PAC",
                            f"{analysis['correios_pac']} pedidos"
                        )
                    
                    with col3:
                        st.metric(
                            "Correios SEDEX",
                            f"{analysis['correios_sedex']} pedidos"
                        )
                    
                    with col4:
                        st.metric(
                            "Transportadoras",
                            f"{analysis['transportadoras']} pedidos"
                        )
                    
                    if analysis['other_couriers'] > 0:
                        st.info(f"📌 {analysis['other_couriers']} pedido(s) com outras transportadoras")
                    
                    # Resumo Financeiro
                    st.markdown("### Resumo Financeiro")
                    
                    col1, col2 = st.columns([1, 1])
                    
                    with col1:
                        st.write("**Despesas:**")
                        st.write(f"- Taxas Pagamento: R$ {analysis['total_payment_taxes']:,.2f}")
                        st.write(f"- Taxa Plataforma: R$ {analysis['total_received'] * analysis['platform_tax']:,.2f}")
                        st.write(f"- Custo Produção: R$ {analysis['total_prod_cost']:,.2f}")
                        st.write(f"- Gasto ADS: R$ {analysis['ads_cost']:,.2f}")
                    
                    with col2:
                        st.write("**Resultado:**")
                        color = '🟢' if analysis['net_profit'] > 0 else '🔴'
                        st.write(f"{color} **Lucro Líquido: R$ {analysis['net_profit']:,.2f}**")
                        if analysis['total_received'] > 0:
                            margin = (analysis['net_profit'] / analysis['total_received'] * 100)
                            st.write(f"**Margem: {margin:.1f}%**")
                    
                    # Detalhamento por Categoria
                    st.markdown("### Detalhamento por Categoria")
                    
                    category_data = []
                    for cat, data in sorted(
                        analysis['stats'].items(),
                        key=lambda x: x[1]['value'],
                        reverse=True
                    ):
                        category_data.append({
                            'Categoria': cat,
                            'Quantidade': int(data['qty']),
                            'Venda': f"R$ {data['value']:,.2f}",
                            'Custo': f"R$ {data['cost']:,.2f}",
                            'Margem': f"R$ {data['value'] - data['cost']:,.2f}"
                        })
                    
                    st.dataframe(
                        pd.DataFrame(category_data),
                        use_container_width=True,
                        hide_index=True
                    )
                    
                    # Botão de Download PDF e Excel
                    st.markdown("---")
                    st.markdown("### 📥 Downloads")
                    
                    pdf_data = analyzer.generate_pdf(analysis)
                    excel_data = analyzer.generate_excel_report(analysis)
                    filename_pdf = f"Relatorio_{store_name}_{analysis['analysis_date'].strftime('%Y%m%d_%H%M')}.pdf"
                    filename_excel = f"Relatorio_{store_name}_{analysis['analysis_date'].strftime('%Y%m%d_%H%M')}.xlsx"
                    
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.download_button(
                            label="📄 PDF",
                            data=pdf_data,
                            file_name=filename_pdf,
                            mime="application/pdf",
                            use_container_width=True
                        )
                    
                    with col2:
                        st.download_button(
                            label="📊 Excel",
                            data=excel_data,
                            file_name=filename_excel,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    
                    with col3:
                        # Botão para fazer nova análise
                        if st.button("🔄 Nova Análise", use_container_width=True):
                            st.session_state.show_results = False
                            st.rerun()
            
            except Exception as e:
                st.error(f"❌ Erro ao processar arquivo: {str(e)}")
    
    with tab2:
        st.markdown("## 🔍 Análises Avançadas")
        
        if not st.session_state.get('show_results') or not st.session_state.get('analysis_data'):
            st.info("⏳ Execute uma análise na aba '📈 Análise Básica' para ver as análises avançadas.")
        else:
            analysis = st.session_state.analysis_data
            
            # Criar sub-abas para análises
            subab1, subab2, subab3, subab4, subab5, subab6 = st.tabs([
                "💳 Pagamentos", 
                "👥 Clientes", 
                "📅 Timeline", 
                "🗺️ Geográfica",
                "📦 Fulfillment",
                "💳 Cupons"
            ])
            
            # ===== ABA: MÉTODOS DE PAGAMENTO =====
            with subab1:
                st.markdown("### 💳 Análise Detalhada de Métodos de Pagamento")
                
                if analysis['tax_analysis']['total_by_method']:
                    tax_data = []
                    for method, total in analysis['tax_analysis']['total_by_method'].items():
                        tax_amount = analysis['tax_analysis']['tax_by_method'].get(method, 0)
                        percentage = (total / analysis['total_received'] * 100) if analysis['total_received'] > 0 else 0
                        tax_data.append({
                            'Método': method,
                            'Total Vendas': f"R$ {total:,.2f}",
                            'Percentual': f"{percentage:.1f}%",
                            'Taxa Total': f"R$ {tax_amount:,.2f}",
                            'Taxa %': f"{(tax_amount/total*100):.2f}%" if total > 0 else "0%"
                        })
                    
                    st.dataframe(
                        pd.DataFrame(tax_data),
                        use_container_width=True,
                        hide_index=True
                    )
                    
                    # Gráfico de métodos de pagamento
                    payment_chart = create_payment_method_chart(analysis['tax_analysis'])
                    if payment_chart:
                        st.plotly_chart(payment_chart, use_container_width=True)
                else:
                    st.warning("Coluna de método de pagamento não encontrada no CSV")
            
            # ===== ABA: CLIENTES =====
            with subab2:
                st.markdown("### 👥 Análise de Clientes")
                
                repeat = analysis['repeat_customers']
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("Clientes Únicos", repeat['total_unique_customers'])
                
                with col2:
                    st.metric("Clientes Repeat", repeat['repeat_customers'])
                
                with col3:
                    st.metric("% Repeat", f"{repeat['repeat_percentage']:.1f}%")
                
                st.markdown("---")
                st.markdown("### Top 10 Clientes")
                
                if repeat['top_customers']:
                    top_customers_data = []
                    for customer in repeat['top_customers']:
                        top_customers_data.append({
                            'Email': customer['email'],
                            'Pedidos': customer['orders'],
                            'Gasto Total': f"R$ {customer['total_spent']:,.2f}",
                            'Ticket Médio': f"R$ {customer['avg_order']:,.2f}"
                        })
                    
                    st.dataframe(
                        pd.DataFrame(top_customers_data),
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.warning("Nenhum dado de cliente disponível (coluna 'Email' não encontrada)")
            
            # ===== ABA: TIMELINE =====
            with subab3:
                st.markdown("### 📅 Análise de Vendas por Período")
                
                timeline = analysis['timeline']
                
                if timeline['best_day']:
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        day, day_data = timeline['best_day']
                        st.metric(
                            "Melhor Dia",
                            day,
                            f"R$ {day_data['value']:,.2f} ({day_data['count']} pedidos)"
                        )
                    
                    with col2:
                        if timeline['best_week']:
                            week, week_data = timeline['best_week']
                            st.metric(
                                "Melhor Semana",
                                week,
                                f"R$ {week_data['value']:,.2f} ({week_data['count']} pedidos)"
                            )
                
                st.markdown("---")
                
                # Vendas por dia
                if timeline['daily_sales']:
                    st.markdown("#### 📊 Vendas Diárias")
                    daily_data = []
                    for date, data in sorted(timeline['daily_sales'].items(), reverse=True):
                        daily_data.append({
                            'Data': date,
                            'Valor': f"R$ {data['value']:,.2f}",
                            'Pedidos': data['count']
                        })
                    
                    st.dataframe(
                        pd.DataFrame(daily_data),
                        use_container_width=True,
                        hide_index=True
                    )
                
                # Vendas por mês
                if timeline['monthly_sales']:
                    st.markdown("#### 📈 Vendas Mensais")
                    monthly_data = []
                    for month, data in sorted(timeline['monthly_sales'].items()):
                        monthly_data.append({
                            'Mês': month,
                            'Valor': f"R$ {data['value']:,.2f}",
                            'Pedidos': data['count']
                        })
                    
                    st.dataframe(
                        pd.DataFrame(monthly_data),
                        use_container_width=True,
                        hide_index=True
                    )
            
            # ===== ABA: GEOGRÁFICA =====
            with subab4:
                st.markdown("### 🗺️ Análise Geográfica")
                
                geo = analysis['geographic']
                
                # Top Estados
                if geo['top_states']:
                    st.markdown("#### 🏆 Top Estados")
                    states_data = []
                    for state in geo['top_states']:
                        states_data.append({
                            'Estado': state[0],
                            'Total': f"R$ {state[1]:,.2f}",
                            'Pedidos': int(state[2]),
                            'Subtotal': f"R$ {state[3]:,.2f}"
                        })
                    
                    st.dataframe(
                        pd.DataFrame(states_data),
                        use_container_width=True,
                        hide_index=True
                    )
                
                st.markdown("---")
                
                # Top Cidades
                if geo['top_cities']:
                    st.markdown("#### 🏘️ Top Cidades")
                    cities_data = []
                    for city in geo['top_cities']:
                        cities_data.append({
                            'Cidade': city[0],
                            'Total': f"R$ {city[1]:,.2f}",
                            'Pedidos': int(city[2]),
                            'Subtotal': f"R$ {city[3]:,.2f}"
                        })
                    
                    st.dataframe(
                        pd.DataFrame(cities_data),
                        use_container_width=True,
                        hide_index=True
                    )
                
                if not geo['top_states'] and not geo['top_cities']:
                    st.warning("Nenhum dado geográfico disponível (colunas de localização não encontradas)")
            
            # ===== ABA: FULFILLMENT =====
            with subab5:
                st.markdown("### 📦 Status de Fulfillment")
                
                fulfill = analysis['fulfillment']
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("Total de Pedidos", fulfill['total_orders'])
                
                with col2:
                    st.metric("Entregues", fulfill['fulfilled'])
                
                with col3:
                    st.metric("Não Entregues", fulfill['unfulfilled'])
                
                with col4:
                    st.metric("Taxa Entrega", f"{fulfill['fulfillment_rate']:.1f}%")
                
                st.markdown("---")
                
                if fulfill['partial'] > 0:
                    st.info(f"⚠️ {fulfill['partial']} pedido(s) parcialmente entregue(s)")
                
                if fulfill['cancelled'] > 0:
                    st.warning(f"❌ {fulfill['cancelled']} pedido(s) cancelado(s)")
                
                st.markdown("---")
                
                # Pedidos pendentes
                if fulfill['pending_fulfillment']:
                    st.markdown("#### ⏳ Pedidos Pendentes de Entrega")
                    pending_data = []
                    for order in fulfill['pending_fulfillment']:
                        pending_data.append({
                            'Pedido': order[0],
                            'Data': order[1],
                            'Valor': f"R$ {order[2]:,.2f}"
                        })
                    
                    st.dataframe(
                        pd.DataFrame(pending_data),
                        use_container_width=True,
                        hide_index=True
                    )
            
            # ===== ABA: CUPONS =====
            with subab6:
                st.markdown("### 💳 Análise de Cupons")
                
                disc = analysis['discounts']
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("Cupons Utilizados", disc['total_discounts'])
                
                with col2:
                    st.metric("Desconto Total", f"R$ {disc['total_discount_value']:,.2f}")
                
                with col3:
                    avg_discount = disc['total_discount_value'] / disc['total_discounts'] if disc['total_discounts'] > 0 else 0
                    st.metric("Desconto Médio", f"R$ {avg_discount:,.2f}")
                
                st.markdown("---")
                
                if disc['top_codes']:
                    st.markdown("#### 🏆 Top Cupons")
                    codes_data = []
                    for code in disc['top_codes']:
                        codes_data.append({
                            'Cupom': code['code'],
                            'Usos': int(code['usage_count']),
                            'Desconto Total': f"R$ {code['total_discount']:,.2f}",
                            'Ticket Médio': f"R$ {code['avg_order_value']:,.2f}"
                        })
                    
                    st.dataframe(
                        pd.DataFrame(codes_data),
                        use_container_width=True,
                        hide_index=True
                    )
                
                if not disc['top_codes']:
                    st.info("Nenhum cupom foi usado neste período")
            
            # ===== ABA: GRÁFICOS =====
            subab_charts = st.expander("📊 Gráficos & Visualizações", expanded=False)
            with subab_charts:
                st.markdown("### 📈 Visualizações Detalhadas")
                
                # Gráfico de vendas diárias
                sales_chart = create_sales_chart(analysis['timeline'])
                if sales_chart:
                    st.plotly_chart(sales_chart, use_container_width=True)
                
                # Gráfico de categorias
                cat_chart = create_category_chart(analysis['stats'])
                if cat_chart:
                    st.plotly_chart(cat_chart, use_container_width=True)
                
                col1, col2 = st.columns(2)
                
                with col1:
                    # Gráfico geográfico
                    geo_chart = create_geographic_chart(analysis['geographic'])
                    if geo_chart:
                        st.plotly_chart(geo_chart, use_container_width=True)
                
                with col2:
                    # Gráfico fulfillment
                    fulfill_chart = create_fulfillment_chart(analysis['fulfillment'])
                    if fulfill_chart:
                        st.plotly_chart(fulfill_chart, use_container_width=True)
                
                # Gráfico de cupons
                coupon_chart = create_coupon_chart(analysis['discounts'])
                if coupon_chart:
                    st.plotly_chart(coupon_chart, use_container_width=True)
            
            # ===== ABA: ROI POR CUPOM =====
            subab_roi = st.expander("💰 ROI por Cupom", expanded=False)
            with subab_roi:
                st.markdown("### 💹 Análise de ROI")
                
                # Pega os dados já carregados
                df_for_roi = st.session_state.get('df_analysis', pd.DataFrame())
                roi_data = analyzer._calculate_roi_by_coupon(df_for_roi)
                
                if roi_data:
                    roi_list = []
                    for code, data in roi_data.items():
                        roi_list.append({
                            'Cupom': data['code'],
                            'Usos': data['orders'],
                            'Receita Gerada': f"R$ {data['total_revenue']:,.2f}",
                            'Desconto Dado': f"R$ {data['total_discount']:,.2f}",
                            'ROI (%)': f"{data['roi']:.1f}%",
                            'Aumento Ticket': f"{data['ticket_increase']:+.1f}%"
                        })
                    
                    roi_df = pd.DataFrame(roi_list)
                    st.dataframe(roi_df, use_container_width=True, hide_index=True)
                    
                    st.markdown("---")
                    st.markdown("**Interpretação:**")
                    st.text("ROI = (Receita - Desconto) / Desconto × 100%")
                    st.text("Aumento Ticket = (Ticket com cupom - Ticket sem cupom) / Ticket sem cupom × 100%")
                else:
                    st.info("Nenhum dado de cupom disponível para cálculo de ROI")
            
            # ===== ABA: COMPARAÇÃO DE PERÍODOS =====
            subab_comp = st.expander("📊 Comparar Períodos", expanded=False)
            with subab_comp:
                st.markdown("### 📅 Comparação entre Períodos")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.subheader("Período 1")
                    p1_from = st.date_input("De", key="p1_from")
                    p1_to = st.date_input("Até", key="p1_to")
                
                with col2:
                    st.subheader("Período 2")
                    p2_from = st.date_input("De", key="p2_from")
                    p2_to = st.date_input("Até", key="p2_to")
                
                if st.button("▶️ Comparar Períodos"):
                    try:
                        # Re-carregar dados se disponível
                        if 'df_analysis' in st.session_state:
                            comparison = analyzer._compare_periods(
                                st.session_state.df_analysis,
                                pd.Timestamp(p1_from),
                                pd.Timestamp(p1_to),
                                pd.Timestamp(p2_from),
                                pd.Timestamp(p2_to)
                            )
                            
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                st.markdown(f"**Período 1:** {p1_from} a {p1_to}")
                                st.metric("Pedidos", int(comparison['period1']['orders']))
                                st.metric("Receita", f"R$ {comparison['period1']['revenue']:,.2f}")
                                st.metric("Ticket Médio", f"R$ {comparison['period1']['avg_ticket']:,.2f}")
                            
                            with col2:
                                st.markdown(f"**Período 2:** {p2_from} a {p2_to}")
                                st.metric("Pedidos", int(comparison['period2']['orders']), 
                                         f"{comparison['growth']['orders_pct']:+.1f}%")
                                st.metric("Receita", f"R$ {comparison['period2']['revenue']:,.2f}",
                                         f"{comparison['growth']['revenue_pct']:+.1f}%")
                                st.metric("Ticket Médio", f"R$ {comparison['period2']['avg_ticket']:,.2f}",
                                         f"{comparison['growth']['ticket_pct']:+.1f}%")
                    except Exception as e:
                        st.error(f"Erro ao comparar períodos: {str(e)}")
            
            # ===== ABA: HISTÓRICO =====
            subab_hist = st.expander("📚 Histórico de Análises", expanded=False)
            with subab_hist:
                st.markdown("### 📋 Análises Anteriores")
                
                history = load_analysis_history()
                
                if history:
                    history_data = []
                    for entry in reversed(history[-20:]):  # Últimas 20
                        history_data.append({
                            'Data': entry['timestamp'][:10],
                            'Hora': entry['timestamp'][11:16],
                            'Loja': entry['store'],
                            'Pedidos': entry['paid_orders'],
                            'Receita': f"R$ {entry['revenue']:,.2f}",
                            'Lucro': f"R$ {entry['profit']:,.2f}"
                        })
                    
                    st.dataframe(pd.DataFrame(history_data), use_container_width=True, hide_index=True)
                    
                    st.markdown("---")
                    if st.button("🗑️ Limpar Histórico"):
                        if os.path.exists(HISTORY_FILE):
                            os.remove(HISTORY_FILE)
                            st.success("Histórico limpo!")
                            st.rerun()
                else:
                    st.info("Nenhuma análise anterior encontrada")
                
                # Salvar análise atual no histórico
                if st.button("💾 Salvar Análise no Histórico"):
                    save_analysis_history(analysis)
                    st.success("✅ Análise salva no histórico!")
    
    with tab3:
        st.markdown("## ⚙️ Gerenciamento de Categorias")
        
        # Inicializar categorias na session_state se não existir
        if 'custom_categories' not in st.session_state:
            st.session_state.custom_categories = DEFAULT_CATEGORIAS_CONFIG.copy()
        
        # Criar instância do analyzer com categorias customizadas
        analyzer = VendasAnalyzerWeb(st.session_state.custom_categories)
        
        st.markdown("""
        Configure as categorias de produtos para uma análise mais precisa. 
        Cada categoria tem palavras-chave que ajudam o sistema a identificar automaticamente os produtos.
        """)
        
        # Seção: Adicionar nova categoria
        st.markdown("### ➕ Adicionar Nova Categoria")
        
        col1, col2, col3 = st.columns([2, 3, 1])
        
        with col1:
            new_category_name = st.text_input(
                "Nome da Categoria",
                placeholder="Ex: Camiseta Infantil",
                key="new_cat_name"
            )
        
        with col2:
            new_keywords = st.text_input(
                "Palavras-chave (separadas por vírgula)",
                placeholder="Ex: infantil, criança, baby",
                key="new_cat_keywords"
            )
        
        with col3:
            if st.button("➕ Adicionar", use_container_width=True):
                if new_category_name.strip() and new_keywords.strip():
                    keywords_list = [kw.strip() for kw in new_keywords.split(',') if kw.strip()]
                    if analyzer.add_category(new_category_name.strip(), keywords_list):
                        st.session_state.custom_categories = analyzer.categorias_config.copy()
                        st.success(f"✅ Categoria '{new_category_name}' adicionada!")
                        st.rerun()
                    else:
                        st.error("❌ Erro ao adicionar categoria. Verifique os dados.")
                else:
                    st.error("❌ Preencha o nome da categoria e pelo menos uma palavra-chave.")
        
        st.markdown("---")
        
        # Seção: Categorias existentes
        st.markdown("### 📋 Categorias Configuradas")
        
        if not analyzer.get_categories_list():
            st.info("Nenhuma categoria configurada. Adicione uma acima.")
        else:
            # Mostrar categorias em um formato editável
            categories_to_remove = []
            
            for category_name in analyzer.get_categories_list():
                with st.expander(f"📦 {category_name}", expanded=False):
                    col1, col2, col3 = st.columns([2, 3, 1])
                    
                    current_keywords = analyzer.categorias_config[category_name]
                    
                    with col1:
                        edit_name = st.text_input(
                            "Nome",
                            value=category_name,
                            key=f"edit_name_{category_name}"
                        )
                    
                    with col2:
                        edit_keywords = st.text_input(
                            "Palavras-chave",
                            value=", ".join(current_keywords),
                            key=f"edit_keywords_{category_name}"
                        )
                    
                    with col3:
                        col3_1, col3_2 = st.columns(2)
                        
                        with col3_1:
                            if st.button("💾 Salvar", key=f"save_{category_name}"):
                                new_keywords_list = [kw.strip() for kw in edit_keywords.split(',') if kw.strip()]
                                if analyzer.update_category(category_name, edit_name, new_keywords_list):
                                    st.session_state.custom_categories = analyzer.categorias_config.copy()
                                    st.success(f"✅ Categoria atualizada!")
                                    st.rerun()
                                else:
                                    st.error("❌ Erro ao atualizar categoria.")
                        
                        with col3_2:
                            if st.button("🗑️ Remover", key=f"remove_{category_name}"):
                                if analyzer.remove_category(category_name):
                                    st.session_state.custom_categories = analyzer.categorias_config.copy()
                                    st.success(f"✅ Categoria '{category_name}' removida!")
                                    st.rerun()
                                else:
                                    st.error("❌ Erro ao remover categoria.")
            
            # Botão para resetar para padrão
            st.markdown("---")
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("🔄 Resetar para Padrão", use_container_width=True):
                    st.session_state.custom_categories = DEFAULT_CATEGORIAS_CONFIG.copy()
                    st.success("✅ Categorias resetadas para configuração padrão!")
                    st.rerun()
            
            with col2:
                # Mostrar resumo
                total_cats = len(analyzer.get_categories_list())
                total_keywords = sum(len(keywords) for keywords in analyzer.categorias_config.values())
                st.metric("Total de Categorias", total_cats)
                st.metric("Total de Palavras-chave", total_keywords)
        
        # Seção: Teste de categorização
        st.markdown("---")
        st.markdown("### 🧪 Teste de Categorização")
        
        test_product = st.text_input(
            "Digite o nome de um produto para testar:",
            placeholder="Ex: Camiseta Oversized Preta"
        )
        
        if test_product:
            detected_category = analyzer._identify_category(test_product)
            if detected_category == "Outros":
                st.warning(f"⚠️ Produto '{test_product}' foi categorizado como 'Outros'")
                st.info("💡 Adicione palavras-chave relevantes para categorizar este produto automaticamente.")
            else:
                st.success(f"✅ Produto '{test_product}' foi categorizado como: **{detected_category}**")
    
    with tab4:
        st.markdown("""
        ## 📖 Como Usar
        
        ### Passo 1: Prepare seu arquivo CSV
        - Exporte seus dados de vendas da Shopify ou plataforma de vendas
        - O arquivo deve conter as colunas: Name, Financial Status, Lineitem name, Lineitem quantity, Lineitem price, Subtotal, Shipping, Total
        - **Importante:** Para análise de métodos de pagamento, adicione uma coluna com o método de pagamento (Payment Method, Pagamento, etc)
        - **Opcional:** Adicione uma coluna com informações de transportadora (Shipping Name, Shipping Method, etc) para análise de frete
        
        ### Passo 2: Carregue o arquivo
        - Clique em "Selecione seu arquivo CSV de vendas"
        - Escolha o arquivo da sua loja
        
        ### Passo 3: Configure os custos
        - Defina o custo de produção para cada categoria de produto
        - Configure o custo padrão para produtos não categorizados
        
        ### Passo 4: Configure as taxas por método de pagamento
        - **Taxa Cartão**: Taxa cobrada pelo método de pagamento com cartão
        - **Taxa Pix**: Taxa cobrada por transações Pix
        - **Taxa Boleto**: Taxa cobrada por boleto bancário
        
        O sistema identifica automaticamente o método de pagamento no CSV e aplica a taxa correspondente!
        
        ### Passo 5: Configure taxas adicionais
        - Taxa de Plataforma: taxa da plataforma de vendas (exemplo: Shopify)
        - Gasto com ADS: total gasto em publicidade
        
        ### Passo 6: Gere a análise
        - Clique em "GERAR ANÁLISE COMPLETA"
        - Visualize os resultados na tela
        - Baixe o relatório em PDF ou Excel se desejado
        
        ## 💳 Análise de Métodos de Pagamento
        
        **NOVO:** O sistema agora detecta automaticamente o método de pagamento de cada transação!
        
        ### Como funciona:
        1. O sistema procura por uma coluna de método de pagamento no CSV
        2. Identifica se é Cartão, Pix, Boleto ou outro método
        3. Aplica a taxa correta para cada método
        4. Mostra análise detalhada na aba "Análises Avançadas"
        
        ### Mapeamento automático:
        - **Cartão**: qualquer transação com "cartão", "credit card", "mastercard", "visa"
        - **Pix**: qualquer transação com "pix"
        - **Boleto**: qualquer transação com "boleto", "bancário"
        - **Outro**: qualquer outro método não mapeado
        
        ## 📦 Análise de Frete e Transportadoras
        
        O sistema agora analisa automaticamente:
        - **Vendas com Frete Grátis**: Pedidos onde o campo "Shipping" é 0
        - **Correios PAC**: Pedidos identificados como PAC
        - **Correios SEDEX**: Pedidos identificados como SEDEX
        - **Transportadoras**: Pedidos com outras transportadoras (Loggi, JNE, etc)
        
        ℹ️ *Nota: Para que a análise de transportadora funcione corretamente, seu CSV deve incluir uma coluna com a informação da transportadora.*
        
        ## 🔍 Análises Avançadas
        
        A aba **"🔍 Análises Avançadas"** fornece insights profundos sobre seu negócio:
        
        ### 💳 Análise de Métodos de Pagamento
        - Vendas totais por método
        - Taxa cobrada por cada método
        - Percentual de cada método
        - Comparação de rentabilidade por método
        
        ### 👥 Análise de Clientes
        - **Clientes Únicos**: Quantidade total de clientes diferentes
        - **Clientes Repeat**: Quantos clientes compraram mais de uma vez
        - **% Repeat**: Percentual de clientes que repetiram
        - **Top 10 Clientes**: Seus melhores clientes, com ticket médio
        
        ### 📅 Timeline de Vendas
        - **Melhor Dia**: Qual dia teve mais vendas
        - **Melhor Semana**: Qual semana foi mais lucrativa
        - **Vendas Diárias**: Detalhamento dia a dia
        - **Vendas Mensais**: Análise por mês
        
        ### 🗺️ Análise Geográfica
        - **Top Estados**: Qual estado gera mais receita
        - **Top Cidades**: Qual cidade é mais lucrativa
        - **Detalhamento Completo**: Venda e quantidade por local
        
        ### 📦 Status de Fulfillment
        - **Taxa de Entrega**: Percentual de pedidos entregues
        - **Pedidos Pendentes**: Lista de pedidos que não foram entregues
        
        ### 💳 Análise de Cupons
        - **Total de Usos**: Quantas vezes cupons foram usados
        - **Desconto Total**: Quanto você distribuiu em descontos
        - **Top Cupons**: Quais cupons funcionam melhor
        - **Ticket Médio**: Qual cupom traz pedidos maiores
        
        ## 🎯 Sobre as Categorias
        
        O sistema identifica automaticamente categorias de produtos por palavras-chave. 
        Você pode **personalizar completamente as categorias** na aba "⚙️ Categorias".
        
        ### Categorias Padrão Incluídas:
        - **Oversized**: produtos com "oversized" no nome
        - **Short 2 em 1**: produtos com "short" ou "2 em 1" no nome
        - **Dryfit**: produtos com "dryfit" ou "dry fit" no nome
        - **Moletom**: produtos com "moletom" ou "hoodie" no nome
        - **Calça**: produtos com "calça", "calca" ou "pants" no nome
        - **Combo**: produtos com "combo" ou "kit" no nome
        
        ### Como Adicionar Novas Categorias:
        1. Vá na aba **"⚙️ Categorias"**
        2. Digite o nome da categoria (ex: "Camiseta Infantil")
        3. Adicione palavras-chave separadas por vírgula (ex: "infantil, criança, baby")
        4. Clique em **"➕ Adicionar"**
        
        ## 💡 Dicas
        
        ✅ **Personalize as categorias** conforme seu catálogo de produtos
        ✅ Use nomes de produtos consistentes para melhor categorização
        ✅ Adicione variações de palavras (ex: "infantil, criança, baby, kids")
        ✅ Configure as taxas corretas para cada método de pagamento
        ✅ Atualize regularmente os custos de produção por categoria
        ✅ Revise as taxas de gateway conforme suas negociações
        ✅ Acompanhe o gasto com ADS para medir ROI
        ✅ **NOVO:** Monitore as vendas por método de pagamento para otimizar custos
        ✅ **NOVO:** Analise quais métodos de pagamento são mais rentáveis
        """)


if __name__ == "__main__":
    # Inicializa estado da sessão
    if 'show_results' not in st.session_state:
        st.session_state.show_results = False
    if 'analysis_data' not in st.session_state:
        st.session_state.analysis_data = None
    
    main()