# analyzer.py
import re
import pandas as pd
from datetime import datetime
from typing import Dict

DEFAULT_CATEGORIAS_CONFIG = {
    "Oversized": ["oversized"],
    "Short 2 em 1": ["short", "2 em 1", "2em1"],
    "Dryfit": ["dryfit", "dry fit"],
    "Moletom": ["moletom", "hoodie"],
    "Calça": ["calça", "calca", "pants"],
    "Combo": ["combo", "kit"]
}

class VendasAnalyzerWeb:
    def __init__(self, categorias_config: Dict[str, list] = None):
        self.categorias_config = categorias_config or DEFAULT_CATEGORIAS_CONFIG.copy()

    def get_categories_list(self) -> list:
        return [cat for cat in self.categorias_config.keys() if cat != "Outros"]

    def _identify_category(self, product_name: str) -> str:
        product_name_lower = str(product_name).lower().strip()
        for category, keywords in self.categorias_config.items():
            for keyword in keywords:
                if keyword.lower() in product_name_lower: return category
        return "Outros"

    def _identify_payment_method(self, row: pd.Series) -> str:
        payment_str = str(row.get('Payment Method', '')) + " " + str(row.get('Tags', '')) + " " + str(row.get('Note Attributes', ''))
        payment_lower = payment_str.lower()
        if 'pix' in payment_lower: return 'pix'
        elif 'boleto' in payment_lower: return 'boleto'
        elif 'cartao' in payment_lower or 'cartão' in payment_lower or 'credit' in payment_lower: return 'cartao'
        else: return 'outro'

    def _identify_installments(self, row: pd.Series) -> int:
        payment_str = str(row.get('Payment Method', '')) + " " + str(row.get('Tags', '')) + " " + str(row.get('Note Attributes', '')) + " " + str(row.get('Payment Terms Name', ''))
        match = re.search(r'\b(\d{1,2})\s*[xX]\b', payment_str, re.IGNORECASE)
        if match: return int(match.group(1))
        match = re.search(r'(?:installments|parcelas)\s*:\s*(\d{1,2})', payment_str, re.IGNORECASE)
        if match: return int(match.group(1))
        return 1

    def _calculate_abc_curve(self, df: pd.DataFrame) -> Dict:
        if 'Lineitem name' not in df.columns: return {}
        paid = df[df['Financial Status'].str.lower() == 'paid'].copy()
        paid['Line_Revenue'] = pd.to_numeric(paid['Lineitem price'], errors='coerce').fillna(0) * pd.to_numeric(paid['Lineitem quantity'], errors='coerce').fillna(0)
        prod_rev = paid.groupby('Lineitem name')['Line_Revenue'].sum().reset_index().sort_values('Line_Revenue', ascending=False)
        total_rev = prod_rev['Line_Revenue'].sum()
        if total_rev == 0: return {}
        prod_rev['Cum_Rev'] = prod_rev['Line_Revenue'].cumsum()
        prod_rev['Cum_Pct'] = prod_rev['Cum_Rev'] / total_rev
        abc = {'A': [], 'B': [], 'C': []}
        for _, row in prod_rev.iterrows():
            item = {'product': row['Lineitem name'], 'revenue': row['Line_Revenue']}
            if row['Cum_Pct'] <= 0.80: abc['A'].append(item)
            elif row['Cum_Pct'] <= 0.95: abc['B'].append(item)
            else: abc['C'].append(item)
        return abc

    def _analyze_heatmap(self, df: pd.DataFrame) -> list:
        if 'Created at' not in df.columns: return []
        paid = df[df['Financial Status'].str.lower() == 'paid'].copy()
        paid['Created at'] = pd.to_datetime(paid['Created at'], errors='coerce')
        paid = paid.dropna(subset=['Created at'])
        days_map = {0: 'Segunda', 1: 'Terça', 2: 'Quarta', 3: 'Quinta', 4: 'Sexta', 5: 'Sábado', 6: 'Domingo'}
        paid['Weekday'] = paid['Created at'].dt.dayofweek.map(days_map)
        paid['Hour'] = paid['Created at'].dt.hour
        return paid.groupby(['Weekday', 'Hour']).size().reset_index(name='Count').to_dict('records')

    def _analyze_geographic(self, df: pd.DataFrame) -> Dict:
        unique_paid = df[df['Financial Status'].str.lower() == 'paid'].drop_duplicates(subset=['Name']).copy()
        geo_data = {"top_states": [], "top_cities": []}
        unique_paid['Total'] = pd.to_numeric(unique_paid['Total'], errors='coerce').fillna(0)
        
        province_col = 'Shipping Province' if 'Shipping Province' in df.columns else 'Billing Province'
        city_col = 'Shipping City' if 'Shipping City' in df.columns else 'Billing City'
        
        if province_col in unique_paid.columns:
            state_sales = unique_paid.groupby(province_col)['Total'].agg(['sum', 'count']).reset_index().sort_values('sum', ascending=False)
            geo_data['top_states'] = state_sales.head(10).values.tolist()
        
        if city_col in unique_paid.columns:
            city_sales = unique_paid.groupby(city_col)['Total'].agg(['sum', 'count']).reset_index().sort_values('sum', ascending=False)
            geo_data['top_cities'] = city_sales.head(10).values.tolist()
            
        return geo_data

    def process_data(self, df: pd.DataFrame, store_name: str, costs_map: dict, default_cost: float, card_taxes: dict, pix_tax: float, boleto_tax_brl: float, gateway_antifraude_brl: float, platform_tax: float, ads_cost: float, traffic_manager_cost: float) -> dict:
        df = df.copy()
        for col in ['Subtotal', 'Shipping', 'Total', 'Lineitem quantity', 'Lineitem price']:
            if col in df.columns: df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        paid = df[df['Financial Status'].str.lower() == 'paid']
        unique_paid = paid.drop_duplicates(subset=['Name'])
        total_received = unique_paid['Total'].sum()
        total_shipping_global = unique_paid['Shipping'].sum()

        unique_customers = len(unique_paid['Email'].unique()) if 'Email' in unique_paid.columns else len(unique_paid)
        cac = (ads_cost + traffic_manager_cost) / unique_customers if unique_customers > 0 else 0
        avg_ticket = total_received / len(unique_paid) if len(unique_paid) > 0 else 0
        orders_per_customer = len(unique_paid) / unique_customers if unique_customers > 0 else 0
        ltv = avg_ticket * orders_per_customer

        # LÓGICA DE FRETES (Gratuito, PAC, SEDEX, Transportadoras)
        shipping_stats = {
            'gratis': {'count': 0, 'value': 0.0},
            'pac': {'count': 0, 'value': 0.0},
            'sedex': {'count': 0, 'value': 0.0},
            'transportadora': {'count': 0, 'value': 0.0},
            'outros': {'count': 0, 'value': 0.0}
        }
        
        courier_col = None
        for col in ['Shipping Method', 'Shipping Name', 'Fulfillment Method', 'Carrier', 'Transportadora']:
            if col in df.columns:
                courier_col = col
                break

        for _, order in unique_paid.iterrows():
            cost = float(order['Shipping'])
            method = str(order[courier_col]).strip().lower() if courier_col and pd.notna(order[courier_col]) else ""
            
            if cost == 0:
                shipping_stats['gratis']['count'] += 1
            elif 'pac' in method:
                shipping_stats['pac']['count'] += 1
                shipping_stats['pac']['value'] += cost
            elif 'sedex' in method:
                shipping_stats['sedex']['count'] += 1
                shipping_stats['sedex']['value'] += cost
            elif method and method != 'nan':
                shipping_stats['transportadora']['count'] += 1
                shipping_stats['transportadora']['value'] += cost
            else:
                shipping_stats['outros']['count'] += 1
                shipping_stats['outros']['value'] += cost

        # Cálculo de Produção
        stats = {}
        for _, row in paid.iterrows():
            cat = self._identify_category(row['Lineitem name'])
            qty = row['Lineitem quantity']
            if cat not in stats: stats[cat] = {"qty": 0, "value": 0, "cost": 0}
            stats[cat]["qty"] += qty
            stats[cat]["value"] += (row['Lineitem price'] * qty)
            stats[cat]["cost"] += (costs_map.get(cat, default_cost) * qty)

        total_prod_cost = sum(c['cost'] for c in stats.values())
        
        # Pagamentos
        payment_stats = {
            'cartao': {'count': 0, 'total': 0, 'tax_amount': 0.0},
            'pix': {'count': 0, 'total': 0, 'tax_amount': 0.0},
            'boleto': {'count': 0, 'total': 0, 'tax_amount': 0.0},
            'outro': {'count': 0, 'total': 0, 'tax_amount': 0.0}
        }

        for _, row in unique_paid.iterrows():
            method = self._identify_payment_method(row)
            order_total = row.get('Total', 0)
            payment_stats[method]['count'] += 1
            payment_stats[method]['total'] += order_total
            if method == 'cartao': payment_stats['cartao']['tax_amount'] += order_total * (card_taxes.get(self._identify_installments(row), card_taxes.get(1, 4.99)) / 100)
            elif method == 'pix': payment_stats['pix']['tax_amount'] += order_total * (pix_tax / 100)
            elif method == 'boleto': payment_stats['boleto']['tax_amount'] += boleto_tax_brl

        total_taxes = sum(s['tax_amount'] for s in payment_stats.values()) + (total_received * platform_tax) + (len(unique_paid) * gateway_antifraude_brl)
        
        return {
            "store_name": store_name,
            "paid_count": len(unique_paid),
            "total_received": total_received,
            "total_shipping": total_shipping_global, # Total cobrado de frete
            "shipping_stats": shipping_stats, # Nova chave de fretes
            "stats": stats,
            "payment_stats": payment_stats,
            "total_prod_cost": total_prod_cost,
            "ads_cost": ads_cost,
            "traffic_manager_cost": traffic_manager_cost,
            "net_profit": total_received - total_taxes - total_prod_cost - (ads_cost + traffic_manager_cost),
            "cac": cac, "ltv": ltv, "avg_ticket": avg_ticket,
            "abc_curve": self._calculate_abc_curve(df),
            "heatmap": self._analyze_heatmap(df),
            "geographic": self._analyze_geographic(df),
            "analysis_date": datetime.now()
        }