import re
import pandas as pd
from datetime import datetime
from typing import Dict

# Configuração Padrão de Categorias Atualizada
DEFAULT_CATEGORIAS_CONFIG = {
    "Kit / Combo": ["kit", "combo", "conjunto"],
    "Moletom / Frio": ["moletom", "hoodie", "casaco", "jaqueta"],
    "Short 2 em 1": ["2 em 1", "2em1"],
    "Rashguard / Compressão": ["rashguard", "compressão", "compressao", "lycra"],
    "Bermuda / Short": ["bermuda", "short", "shorts"],
    "Oversized": ["oversized"],
    "Dryfit / Esportiva": ["dryfit", "dry fit", "esportiva"],
    "Camiseta Padrão": ["camiseta", "t-shirt", "camisa"],
    "Calça": ["calça", "calca", "pants", "jogger"],
    "Acessórios / Outros": ["kimono", "faixa", "luva", "bandagem", "boné", "garrafa"]
}

class VendasAnalyzerWeb:
    """Classe responsável pelo processamento e análise de dados de vendas."""

    def __init__(self, categorias_config: Dict[str, list] = None):
        self.categorias_config = categorias_config or DEFAULT_CATEGORIAS_CONFIG.copy()

    def get_categories_list(self) -> list:
        return [cat for cat in self.categorias_config.keys() if cat != "Outros"]

    def _identify_category(self, product_name: str) -> str:
        product_name_lower = str(product_name).lower().strip()
        for category, keywords in self.categorias_config.items():
            for keyword in keywords:
                if keyword.lower() in product_name_lower:
                    return category
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

    def _analyze_shipping(self, df: pd.DataFrame) -> Dict:
        """Método inteligente e vetorizado para calcular fretes."""
        unique_paid = df[df['Financial Status'].str.lower() == 'paid'].drop_duplicates(subset=['Name']).copy()

        if 'Shipping' not in unique_paid.columns:
            unique_paid['Shipping'] = 0.0
        unique_paid['Shipping'] = pd.to_numeric(unique_paid['Shipping'], errors='coerce').fillna(0.0)

        courier_col = next((col for col in ['Shipping Method', 'Shipping Name', 'Fulfillment Method', 'Carrier', 'Transportadora'] if col in unique_paid.columns), None)
        
        if courier_col:
            unique_paid['Method_Lower'] = unique_paid[courier_col].astype(str).str.lower().fillna('')
        else:
            unique_paid['Method_Lower'] = ''

        mask_gratis = unique_paid['Shipping'] <= 0.01
        mask_pac = unique_paid['Method_Lower'].str.contains('pac', na=False) & ~mask_gratis
        mask_sedex = unique_paid['Method_Lower'].str.contains('sedex', na=False) & ~mask_gratis
        mask_transp = (~mask_gratis) & (~mask_pac) & (~mask_sedex) & (unique_paid['Method_Lower'] != '') & (unique_paid['Method_Lower'] != 'nan')

        return {
            'gratis': {'count': int(mask_gratis.sum()), 'value': 0.0},
            'pac': {'count': int(mask_pac.sum()), 'value': float(unique_paid.loc[mask_pac, 'Shipping'].sum())},
            'sedex': {'count': int(mask_sedex.sum()), 'value': float(unique_paid.loc[mask_sedex, 'Shipping'].sum())},
            'transportadora': {'count': int(mask_transp.sum()), 'value': float(unique_paid.loc[mask_transp, 'Shipping'].sum())}
        }

    def _analyze_repeat_customers(self, df: pd.DataFrame) -> Dict:
        unique_paid = df[df['Financial Status'].str.lower() == 'paid'].drop_duplicates(subset=['Name'])
        if 'Email' not in df.columns:
            return {"total_unique_customers": len(unique_paid), "repeat_customers": 0, "repeat_percentage": 0.0, "top_customers": []}
        
        email_counts = unique_paid['Email'].value_counts()
        repeat_customers = email_counts[email_counts > 1]
        
        top_customers = []
        for email, count in email_counts.head(10).items():
            customer_orders = unique_paid[unique_paid['Email'] == email]
            total_spent = customer_orders['Total'].sum()
            top_customers.append({
                'email': email, 'orders': count, 'total_spent': total_spent, 'avg_order': total_spent / count
            })
            
        return {
            "total_unique_customers": len(email_counts),
            "repeat_customers": len(repeat_customers),
            "repeat_percentage": (len(repeat_customers) / len(email_counts) * 100) if len(email_counts) > 0 else 0.0,
            "top_customers": top_customers
        }

    def _analyze_fulfillment(self, df: pd.DataFrame) -> Dict:
        if 'Fulfillment Status' not in df.columns:
            return {"total_orders": 0, "fulfilled": 0, "unfulfilled": 0, "cancelled": 0}
            
        unique_orders = df.drop_duplicates(subset=['Name'])
        fulfillment_counts = unique_orders['Fulfillment Status'].value_counts()
        
        return {
            "total_orders": len(unique_orders),
            "fulfilled": int(fulfillment_counts.get('fulfilled', 0)),
            "unfulfilled": int(fulfillment_counts.get('unfulfilled', 0)),
            "cancelled": int(fulfillment_counts.get('cancelled', 0))
        }

    def _analyze_discount_codes(self, df: pd.DataFrame) -> Dict:
        if 'Discount Code' not in df.columns:
            return {"total_discounts": 0, "total_discount_value": 0.0, "top_codes": []}
            
        unique_paid = df[df['Financial Status'].str.lower() == 'paid'].drop_duplicates(subset=['Name'])
        with_discount = unique_paid[unique_paid['Discount Code'].notna() & (unique_paid['Discount Code'] != '')]
        
        discount_stats = with_discount.groupby('Discount Code').agg({'Discount Amount': ['sum', 'count'], 'Total': ['mean']}).reset_index()
        discount_stats.columns = ['code', 'total_discount', 'usage_count', 'avg_order_value']
        discount_stats = discount_stats.sort_values('total_discount', ascending=False)
        
        return {
            "total_discounts": int(discount_stats['usage_count'].sum()),
            "total_discount_value": float(discount_stats['total_discount'].sum()),
            "top_codes": discount_stats.head(5).to_dict('records')
        }

    def process_data(self, df: pd.DataFrame, store_name: str, costs_map: dict, default_cost: float, card_taxes: dict, pix_tax: float, boleto_tax_brl: float, gateway_antifraude_brl: float, platform_tax: float, ads_cost: float, traffic_manager_cost: float) -> dict:
        df = df.copy()
        
        # Conversão de colunas financeiras
        for col in ['Subtotal', 'Shipping', 'Total', 'Lineitem quantity', 'Lineitem price']:
            if col in df.columns: df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        paid = df[df['Financial Status'].str.lower() == 'paid']
        unique_paid = paid.drop_duplicates(subset=['Name'])
        
        total_received = unique_paid['Total'].sum()
        total_shipping_global = unique_paid['Shipping'].sum()
        total_items = unique_paid['Subtotal'].sum()

        # Métricas de Clientes (LTV / CAC)
        unique_customers = len(unique_paid['Email'].unique()) if 'Email' in unique_paid.columns else len(unique_paid)
        cac = (ads_cost + traffic_manager_cost) / unique_customers if unique_customers > 0 else 0
        avg_ticket = total_received / len(unique_paid) if len(unique_paid) > 0 else 0
        orders_per_customer = len(unique_paid) / unique_customers if unique_customers > 0 else 0
        ltv = avg_ticket * orders_per_customer

        # Estatísticas de Produção por Categoria
        stats = {}
        for _, row in paid.iterrows():
            cat = self._identify_category(row['Lineitem name'])
            qty = row['Lineitem quantity']
            if cat not in stats: stats[cat] = {"qty": 0, "value": 0, "cost": 0}
            stats[cat]["qty"] += qty
            stats[cat]["value"] += (row['Lineitem price'] * qty)
            stats[cat]["cost"] += (costs_map.get(cat, default_cost) * qty)

        total_prod_cost = sum(c['cost'] for c in stats.values())
        
        # Estatísticas de Pagamentos e Taxas
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
            
            if method == 'cartao':
                payment_stats['cartao']['tax_amount'] += order_total * (card_taxes.get(self._identify_installments(row), card_taxes.get(1, 4.99)) / 100)
            elif method == 'pix':
                payment_stats['pix']['tax_amount'] += order_total * (pix_tax / 100)
            elif method == 'boleto':
                payment_stats['boleto']['tax_amount'] += boleto_tax_brl

        total_taxes = sum(s['tax_amount'] for s in payment_stats.values()) + (total_received * platform_tax) + (len(unique_paid) * gateway_antifraude_brl)
        
        # Retorno final com todas as análises para o Streamlit
        return {
            "store_name": store_name,
            "paid_count": len(unique_paid),
            "total_items": total_items,
            "total_received": total_received,
            "total_shipping": total_shipping_global,
            "shipping_stats": self._analyze_shipping(df),
            "stats": stats,
            "payment_stats": payment_stats,
            "total_prod_cost": total_prod_cost,
            "ads_cost": ads_cost,
            "traffic_manager_cost": traffic_manager_cost,
            "net_profit": total_received - total_taxes - total_prod_cost - (ads_cost + traffic_manager_cost),
            "cac": cac, 
            "ltv": ltv, 
            "avg_ticket": avg_ticket,
            "abc_curve": self._calculate_abc_curve(df),
            "heatmap": self._analyze_heatmap(df),
            "geographic": self._analyze_geographic(df),
            "repeat_customers": self._analyze_repeat_customers(df),
            "fulfillment": self._analyze_fulfillment(df),
            "discounts": self._analyze_discount_codes(df),
            "analysis_date": datetime.now()
        }
