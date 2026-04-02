# -*- coding: utf-8 -*-
"""
Analisador de Vendas Web
Aplicação web para análise de dados de vendas com geração de relatórios em PDF.
Versão 2.0 - Com detecção automática de método de pagamento
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

# Taxas de Pagamento Padrão (em decimal)
DEFAULT_PAYMENT_RATES = {
    "Cartão de crédito": {
        "gateway_tax": 0.0499,  # 4.99%
        "platform_tax": 0.0,
        "installment_fee": 0.0129  # 1.29% por parcela
    },
    "Pix": {
        "gateway_tax": 0.0,  # 0%
        "platform_tax": 0.0,
        "installment_fee": 0.0
    },
    "Boleto": {
        "gateway_tax": 0.0,  # 0%
        "platform_tax": 0.0,
        "installment_fee": 0.0
    }
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


class PaymentMethodDetector:
    """Detecta método de pagamento a partir de informações do CSV."""
    
    @staticmethod
    def detect_payment_method(payment_method_text: str, payment_reference: str = "") -> str:
        """
        Detecta o método de pagamento baseado no texto.
        
        Args:
            payment_method_text: Texto do método de pagamento
            payment_reference: Referência adicional do pagamento
            
        Returns:
            Um dos: "Cartão de crédito", "Pix", "Boleto"
        """
        text = str(payment_method_text).lower().strip()
        ref = str(payment_reference).lower().strip()
        
        # Detecta Pix
        if 'pix' in text or 'pix' in ref:
            return "Pix"
        
        # Detecta Boleto
        if 'boleto' in text or 'boleto' in ref:
            return "Boleto"
        
        # Detecta Cartão de Crédito
        if 'cartão' in text or 'cartao' in text or 'credit' in text or 'card' in text:
            return "Cartão de crédito"
        
        # Padrão é cartão
        return "Cartão de crédito"
    
    @staticmethod
    def detect_installments(payment_reference: str, total_value: float) -> Optional[int]:
        """
        Detecta número de parcelas a partir da referência de pagamento.
        
        Args:
            payment_reference: Referência do pagamento
            total_value: Valor total do pedido
            
        Returns:
            Número de parcelas ou None
        """
        ref = str(payment_reference).lower().strip()
        
        # Procura por padrão "x parcelAS" ou "xX" onde X é número
        import re
        
        # Tenta encontrar "parcelAS x" ou similar
        if 'parcel' in ref:
            numbers = re.findall(r'parcel[a-z]*\s*(\d+)', ref)
            if numbers:
                return int(numbers[0])
        
        # Tenta encontrar números seguidos de "x"
        numbers = re.findall(r'(\d+)\s*x', ref)
        if numbers:
            return int(numbers[0])
        
        return None


class VendasAnalyzerWeb:
    """Analisador de vendas para interface web."""

    def __init__(self, categorias_config: Dict[str, list] = None):
        """Inicializa o analisador."""
        self.categorias_config = categorias_config or DEFAULT_CATEGORIAS_CONFIG.copy()
        self.payment_detector = PaymentMethodDetector()

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

    def _analyze_payment_methods(self, df: pd.DataFrame) -> Dict:
        """
        Analisa vendas por método de pagamento.
        
        Args:
            df: DataFrame com dados de vendas
            
        Returns:
            Dicionário com análise de pagamentos
        """
        unique_paid = df[df['Financial Status'].str.lower() == 'paid'].drop_duplicates(subset=['Name'])
        
        payment_analysis = {
            "by_method": {},
            "total_by_method": {},
            "installments_analysis": [],
            "payment_fees_detail": []
        }
        
        if 'Payment Method' not in df.columns:
            return payment_analysis
        
        # Agrupa por método de pagamento
        for _, order in unique_paid.iterrows():
            payment_method_text = str(order.get('Payment Method', ''))
            payment_reference = str(order.get('Payment Reference', ''))
            total = order['Total']
            
            # Detecta método
            method = self.payment_detector.detect_payment_method(payment_method_text, payment_reference)
            
            if method not in payment_analysis["by_method"]:
                payment_analysis["by_method"][method] = {
                    "count": 0,
                    "total": 0,
                    "avg_ticket": 0,
                    "orders": []
                }
            
            payment_analysis["by_method"][method]["count"] += 1
            payment_analysis["by_method"][method]["total"] += total
            payment_analysis["by_method"][method]["orders"].append({
                'order': order['Name'],
                'value': total,
                'reference': payment_reference
            })
        
        # Calcula médias
        for method, data in payment_analysis["by_method"].items():
            data["avg_ticket"] = data["total"] / data["count"] if data["count"] > 0 else 0
            payment_analysis["total_by_method"][method] = data["total"]
        
        # Analisa parcelamentos
        if 'Payment Reference' in df.columns:
            for _, order in unique_paid.iterrows():
                payment_ref = str(order.get('Payment Reference', ''))
                installments = self.payment_detector.detect_installments(payment_ref, order['Total'])
                
                if installments and installments > 1:
                    payment_analysis["installments_analysis"].append({
                        'order': order['Name'],
                        'method': self.payment_detector.detect_payment_method(
                            str(order.get('Payment Method', '')),
                            payment_ref
                        ),
                        'installments': installments,
                        'total': order['Total'],
                        'per_installment': order['Total'] / installments
                    })
        
        return payment_analysis

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
        
        df_copy = df.copy()
        df_copy['Created at'] = pd.to_datetime(df_copy['Created at'], errors='coerce')
        unique_paid = df_copy[df_copy['Financial Status'].str.lower() == 'paid'].drop_duplicates(subset=['Name'])
        
        daily = unique_paid.groupby(unique_paid['Created at'].dt.date)['Total'].agg(['sum', 'count'])
        daily_sales = {str(date): {'value': value, 'count': count} for date, value, count in zip(daily.index, daily['sum'], daily['count'])}
        
        weekly = unique_paid.groupby(unique_paid['Created at'].dt.isocalendar().week)['Total'].agg(['sum', 'count'])
        weekly_sales = {f"Semana {week}": {'value': value, 'count': count} for week, value, count in zip(weekly.index, weekly['sum'], weekly['count'])}
        
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
        
        province_col = 'Shipping Province' if 'Shipping Province' in df.columns else 'Billing Province'
        city_col = 'Shipping City' if 'Shipping City' in df.columns else 'Billing City'
        
        if province_col in unique_paid.columns:
            state_sales = unique_paid.groupby(province_col).agg({
                'Total': ['sum', 'count'],
                'Subtotal': 'sum'
            }).reset_index()
            state_sales.columns = ['state', 'total_value', 'order_count', 'subtotal']
            state_sales = state_sales.sort_values('total_value', ascending=False)
            
            geo_data['by_state'] = state_sales.to_dict('records')
            geo_data['top_states'] = state_sales.head(5).values.tolist()
        
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
        payment_rates: Dict[str, Dict] = None
    ) -> Dict:
        """
        Processa os dados e retorna análise com detecção de método de pagamento.
        
        Args:
            df: DataFrame com dados de vendas
            store_name: Nome da loja
            costs_map: Mapa de custos por categoria
            default_cost: Custo padrão
            payment_rates: Taxas por método de pagamento
            
        Returns:
            Dicionário com análise completa
        """
        if payment_rates is None:
            payment_rates = DEFAULT_PAYMENT_RATES
        
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

        # ===== ANÁLISE DE MÉTODOS DE PAGAMENTO E TAXAS =====
        payment_analysis = self._analyze_payment_methods(df)
        
        # Calcula taxas por método de pagamento
        total_taxes = 0
        payment_taxes_breakdown = {}
        
        for method, rate_info in payment_rates.items():
            method_total = payment_analysis["total_by_method"].get(method, 0)
            
            # Taxa do gateway
            gateway_tax = method_total * rate_info.get("gateway_tax", 0)
            
            # Taxa da plataforma
            platform_tax = method_total * rate_info.get("platform_tax", 0)
            
            # Taxa de parcelamento
            installment_fee = 0
            for inst in payment_analysis["installments_analysis"]:
                if inst['method'] == method and inst['installments'] > 1:
                    # Cobra taxa por parcela (exceto a primeira)
                    installments_count = inst['installments'] - 1
                    installment_fee += inst['total'] * rate_info.get("installment_fee", 0) * installments_count
            
            method_total_tax = gateway_tax + platform_tax + installment_fee
            total_taxes += method_total_tax
            
            payment_taxes_breakdown[method] = {
                "total": method_total,
                "gateway_tax": gateway_tax,
                "platform_tax": platform_tax,
                "installment_fee": installment_fee,
                "total_tax": method_total_tax,
                "effective_tax_rate": (method_total_tax / method_total * 100) if method_total > 0 else 0
            }
        
        ads_cost = 0  # Pode ser adicionado depois na interface
        net_profit = total_received - total_taxes - total_prod_cost - ads_cost

        # Análise de frete
        free_shipping_orders = unique_paid[unique_paid['Shipping'] == 0]
        free_shipping_count = len(free_shipping_orders)
        free_shipping_value = free_shipping_orders['Total'].sum()
        
        courier_col = None
        transpose_cols = ['Shipping Method', 'Shipping Name', 'Fulfillment Method', 'Carrier', 'Transportadora']
        
        for col in transpose_cols:
            if col in df.columns:
                courier_col = col
                break
        
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

        # Análises adicionais
        repeat_customers = self._analyze_repeat_customers(df)
        timeline = self._analyze_timeline(df)
        geographic = self._analyze_geographic(df)
        fulfillment = self._analyze_fulfillment(df)
        discounts = self._analyze_discount_codes(df)

        # Período das vendas
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
            "total_taxes": total_taxes,
            "total_prod_cost": total_prod_cost,
            "ads_cost": ads_cost,
            "net_profit": net_profit,
            "analysis_date": datetime.now(),
            # Métodos de pagamento
            "payment_analysis": payment_analysis,
            "payment_taxes_breakdown": payment_taxes_breakdown,
            # Frete
            "free_shipping_count": free_shipping_count,
            "free_shipping_value": free_shipping_value,
            "correios_pac": correios_pac,
            "correios_sedex": correios_sedex,
            "transportadoras": transportadoras,
            "other_couriers": other_couriers,
            # Análises avançadas
            "repeat_customers": repeat_customers,
            "timeline": timeline,
            "geographic": geographic,
            "fulfillment": fulfillment,
            "discounts": discounts,
            # Período
            "sales_period": sales_period
        }

    def generate_excel_report(self, analysis_data: Dict) -> bytes:
        """Gera relatório em Excel."""
        output = io.BytesIO()
        wb = Workbook()
        
        header_fill = PatternFill(start_color="1F77B4", end_color="1F77B4", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        title_font = Font(bold=True, size=14)
        
        # Resumo
        ws = wb.active
        ws.title = "Resumo"
        
        ws['A1'] = "RESUMO EXECUTIVO"
        ws['A1'].font = title_font
        
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
            ("Total de Taxas", analysis_data['total_taxes']),
            ("Custo Produção", analysis_data['total_prod_cost']),
            ("Lucro Líquido", analysis_data['net_profit']),
        ]
        
        for label, value in metrics:
            ws[f'A{row}'] = label
            ws[f'B{row}'] = value
            row += 1
        
        # Métodos de Pagamento
        ws = wb.create_sheet("Métodos de Pagamento")
        
        ws['A1'] = "MÉTODO"
        ws['B1'] = "TOTAL (R$)"
        ws['C1'] = "QUANTIDADE"
        ws['D1'] = "TICKET MÉD (R$)"
        ws['E1'] = "TAXAS (R$)"
        ws['F1'] = "TAX RATE (%)"
        
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
        
        row = 2
        for method, breakdown in analysis_data['payment_taxes_breakdown'].items():
            method_data = analysis_data['payment_analysis']['by_method'].get(method, {})
            ws[f'A{row}'] = method
            ws[f'B{row}'] = breakdown['total']
            ws[f'C{row}'] = method_data.get('count', 0)
            ws[f'D{row}'] = method_data.get('avg_ticket', 0)
            ws[f'E{row}'] = breakdown['total_tax']
            ws[f'F{row}'] = breakdown['effective_tax_rate']
            row += 1
        
        # Categorias
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
        
        wb.save(output)
        output.seek(0)
        return output.getvalue()

    @staticmethod
    def generate_pdf(analysis_data: Dict) -> bytes:
        """Gera PDF com a análise."""
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=letter)
        elements = []
        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=16,
            textColor=colors.HexColor('#1f77b4'),
            spaceAfter=30,
            alignment=1
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
        
        date_text = f"Data: {analysis_data['analysis_date'].strftime('%d/%m/%Y %H:%M')}"
        elements.append(Paragraph(date_text, styles['Normal']))
        
        if analysis_data.get('sales_period') and analysis_data['sales_period'].get('start_date'):
            period_text = f"Período: {analysis_data['sales_period']['start_date']} até {analysis_data['sales_period']['end_date']}"
            elements.append(Paragraph(period_text, styles['Normal']))
        
        elements.append(Spacer(1, 0.3*inch))
        
        # Métodos de Pagamento
        elements.append(Paragraph("ANÁLISE DE MÉTODOS DE PAGAMENTO", heading_style))
        data = [['Método', 'Total (R$)', 'Pedidos', 'Taxas (R$)', 'Taxa Efetiva (%)']]
        
        for method, breakdown in analysis_data['payment_taxes_breakdown'].items():
            method_data = analysis_data['payment_analysis']['by_method'].get(method, {})
            data.append([
                method,
                f"R$ {breakdown['total']:,.2f}",
                str(method_data.get('count', 0)),
                f"R$ {breakdown['total_tax']:,.2f}",
                f"{breakdown['effective_tax_rate']:.2f}%"
            ])
        
        table = Table(data, colWidths=[1.5*inch, 1.2*inch, 1*inch, 1.2*inch, 1.1*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f77b4')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.lightblue),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        elements.append(table)
        elements.append(Spacer(1, 0.3*inch))
        
        # Resumo Financeiro
        elements.append(Paragraph("RESUMO FINANCEIRO", heading_style))
        data = [
            ['Total Recebido', f"R$ {analysis_data['total_received']:,.2f}"],
            ['(-) Taxas', f"R$ {analysis_data['total_taxes']:,.2f}"],
            ['(-) Custo Produção', f"R$ {analysis_data['total_prod_cost']:,.2f}"],
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
    
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                history = json.load(f)
        except:
            history = []
    
    analysis_entry = {
        'timestamp': datetime.now().isoformat(),
        'store': analysis_data['store_name'],
        'paid_orders': analysis_data['paid_count'],
        'revenue': float(analysis_data['total_received']),
        'profit': float(analysis_data['net_profit']),
        'shipping': float(analysis_data['total_shipping'])
    }
    
    history.append(analysis_entry)
    
    if len(history) > 100:
        history = history[-100:]
    
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


def main():
    """Função principal da aplicação web."""
    st.title("📊 Analisador de Vendas Multi-Loja v2.0")
    st.markdown("*Com detecção automática de método de pagamento e taxas personalizadas*")
    
    if 'custom_categories' not in st.session_state:
        st.session_state.custom_categories = DEFAULT_CATEGORIAS_CONFIG.copy()
    
    if 'payment_rates' not in st.session_state:
        st.session_state.payment_rates = DEFAULT_PAYMENT_RATES.copy()
    
    analyzer = VendasAnalyzerWeb(st.session_state.custom_categories)
    
    tab1, tab2, tab3, tab4 = st.tabs(["📈 Análise Básica", "💳 Métodos de Pagamento", "⚙️ Configurações", "ℹ️ Ajuda"])
    
    with tab1:
        st.markdown("### 1️⃣ Informações Básicas")
        col1, col2 = st.columns([2, 1])
        
        with col1:
            store_name = st.text_input(
                "Nome da Loja",
                value=DEFAULT_STORE_NAME,
                help="Nome da loja para o relatório"
            )
        
        st.markdown("### 2️⃣ Seleção de Arquivo")
        uploaded_file = st.file_uploader(
            "Selecione seu arquivo CSV de vendas",
            type=['csv'],
            help="Arquivo exportado da Shopify"
        )
        
        if uploaded_file is not None:
            try:
                for encoding in ['utf-8', 'latin-1', 'iso-8859-1']:
                    try:
                        df = pd.read_csv(uploaded_file, encoding=encoding)
                        st.success(f"✅ Arquivo carregado com sucesso ({len(df)} linhas)")
                        break
                    except UnicodeDecodeError:
                        continue
                
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
                
                if st.button("🚀 GERAR ANÁLISE COMPLETA", use_container_width=True):
                    try:
                        analysis_data = analyzer.process_data(
                            df=df,
                            store_name=store_name,
                            costs_map=costs,
                            default_cost=default_cost,
                            payment_rates=st.session_state.payment_rates
                        )
                        
                        st.session_state.analysis_data = analysis_data
                        st.session_state.df_analysis = df
                        st.session_state.show_results = True
                        
                        save_analysis_history(analysis_data)
                        
                        st.success("✅ Análise realizada com sucesso!")
                        
                    except Exception as e:
                        st.error(f"❌ Erro na análise: {str(e)}")
                
                if st.session_state.get('show_results'):
                    analysis = st.session_state.analysis_data
                    
                    st.markdown("---")
                    st.markdown("## 📊 Resultado da Análise")
                    
                    if analysis.get('sales_period') and analysis['sales_period'].get('start_date'):
                        period_text = f"📅 **Período:** {analysis['sales_period']['start_date']} até {analysis['sales_period']['end_date']}"
                        st.info(period_text)
                    
                    # Métricas principais
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.metric("Pedidos Pagos", analysis['paid_count'])
                    
                    with col2:
                        st.metric("Cancelados", analysis['cancelled_count'])
                    
                    with col3:
                        st.metric("Pendentes", analysis['pending_count'])
                    
                    with col4:
                        st.metric(
                            "Lucro",
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
                    
                    # Resumo Financeiro
                    st.markdown("### Resumo Financeiro")
                    
                    col1, col2 = st.columns([1, 1])
                    
                    with col1:
                        st.write("**Despesas:**")
                        st.write(f"- Taxas: R$ {analysis['total_taxes']:,.2f}")
                        st.write(f"- Custo Produção: R$ {analysis['total_prod_cost']:,.2f}")
                    
                    with col2:
                        st.write("**Resultado:**")
                        color = '🟢' if analysis['net_profit'] > 0 else '🔴'
                        st.write(f"{color} **Lucro: R$ {analysis['net_profit']:,.2f}**")
                    
                    # Downloads
                    st.markdown("---")
                    st.markdown("### 📥 Downloads")
                    
                    pdf_data = analyzer.generate_pdf(analysis)
                    excel_data = analyzer.generate_excel_report(analysis)
                    filename_pdf = f"Relatorio_{store_name}_{analysis['analysis_date'].strftime('%Y%m%d_%H%M')}.pdf"
                    filename_excel = f"Relatorio_{store_name}_{analysis['analysis_date'].strftime('%Y%m%d_%H%M')}.xlsx"
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.download_button(
                            label="📄 Baixar PDF",
                            data=pdf_data,
                            file_name=filename_pdf,
                            mime="application/pdf",
                            use_container_width=True
                        )
                    
                    with col2:
                        st.download_button(
                            label="📊 Baixar Excel",
                            data=excel_data,
                            file_name=filename_excel,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
            
            except Exception as e:
                st.error(f"❌ Erro ao processar arquivo: {str(e)}")
    
    with tab2:
        st.markdown("## 💳 Configuração de Métodos de Pagamento")
        
        st.markdown("""
        Configure as taxas para cada método de pagamento. O sistema detectará automaticamente 
        o método no CSV e aplicará as taxas correspondentes.
        """)
        
        for method in DEFAULT_PAYMENT_RATES.keys():
            st.markdown(f"### {method}")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                gateway = st.number_input(
                    f"{method} - Taxa Gateway (%)",
                    value=float(st.session_state.payment_rates[method]["gateway_tax"]) * 100,
                    min_value=0.0,
                    max_value=100.0,
                    step=0.01,
                    key=f"gateway_{method}"
                )
                st.session_state.payment_rates[method]["gateway_tax"] = gateway / 100
            
            with col2:
                platform = st.number_input(
                    f"{method} - Taxa Plataforma (%)",
                    value=float(st.session_state.payment_rates[method]["platform_tax"]) * 100,
                    min_value=0.0,
                    maxvalue=100.0,
                    step=0.01,
                    key=f"platform_{method}"
                )
                st.session_state.payment_rates[method]["platform_tax"] = platform / 100
            
            with col3:
                installment = st.number_input(
                    f"{method} - Taxa Parcelamento (%)",
                    value=float(st.session_state.payment_rates[method]["installment_fee"]) * 100,
                    min_value=0.0,
                    max_value=100.0,
                    step=0.01,
                    key=f"installment_{method}"
                )
                st.session_state.payment_rates[method]["installment_fee"] = installment / 100
            
            st.write(f"**Taxa Efetiva Total: {(gateway + platform + installment):.2f}%**")
            st.markdown("---")
    
    with tab3:
        st.markdown("## ⚙️ Gerenciamento de Categorias")
        
        st.markdown("Configure as categorias de produtos para análise mais precisa.")
        
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
                        st.error("❌ Erro ao adicionar categoria.")
                else:
                    st.error("❌ Preencha o nome e palavras-chave.")
    
    with tab4:
        st.markdown("""
        ## 📖 Como Usar
        
        ### Detecção Automática de Método de Pagamento
        
        O sistema detecta automaticamente o método de pagamento a partir do CSV:
        - **Pix**: Procura por "pix" no Payment Method ou Payment Reference
        - **Boleto**: Procura por "boleto"
        - **Cartão de Crédito**: Padrão para outros casos
        
        ### Cálculo de Taxas e Parcelamento
        
        ✅ **Taxa Gateway**: Cobrada sobre cada transação
        ✅ **Taxa Plataforma**: Taxa adicional da plataforma
        ✅ **Taxa Parcelamento**: Cobrada por parcela (exceto a primeira)
        
        ### Exportação do CSV
        
        Seu arquivo precisa ter essas colunas:
        - `Name` - ID do pedido
        - `Financial Status` - paid, cancelled, pending
        - `Payment Method` - Método de pagamento (Pix, Boleto, Cartão, etc)
        - `Payment Reference` - Referência com info de parcelamento
        - `Lineitem name` - Nome do produto
        - `Lineitem quantity` - Quantidade
        - `Lineitem price` - Preço
        - `Total` - Total do pedido
        """)


if __name__ == "__main__":
    if 'show_results' not in st.session_state:
        st.session_state.show_results = False
    if 'analysis_data' not in st.session_state:
        st.session_state.analysis_data = None
    
    main()