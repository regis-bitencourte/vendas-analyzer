# -*- coding: utf-8 -*-
"""
Analisador de Vendas Web
Aplicação web para análise de dados de vendas com geração de relatórios em PDF.
"""

import io
import logging
from datetime import datetime
from typing import Dict, Optional, Tuple

import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch

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
DEFAULT_TAX_RATE = 4.99

# Categorias padrão (fallback)
DEFAULT_CATEGORIAS_CONFIG = {
    "Oversized": ["oversized"],
    "Short 2 em 1": ["short", "2 em 1", "2em1"],
    "Dryfit": ["dryfit", "dry fit"],
    "Moletom": ["moletom", "hoodie"],
    "Calça": ["calça", "calca", "pants"],
    "Combo": ["combo", "kit"]
}


class VendasAnalyzerWeb:
    """Analisador de vendas para interface web."""

    def __init__(self, categorias_config: Dict[str, list] = None):
        """Inicializa o analisador."""
        self.categorias_config = categorias_config or DEFAULT_CATEGORIAS_CONFIG.copy()

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

    def process_data(
        self,
        df: pd.DataFrame,
        store_name: str,
        costs_map: Dict[str, float],
        default_cost: float,
        card_tax: float,
        platform_tax: float,
        ads_cost: float
    ) -> Dict:
        """
        Processa os dados e retorna análise.
        
        Args:
            df: DataFrame com dados de vendas
            store_name: Nome da loja
            costs_map: Mapa de custos por categoria
            default_cost: Custo padrão
            card_tax: Taxa do cartão em decimal
            platform_tax: Taxa da plataforma em decimal
            ads_cost: Custo com ADS
            
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
        total_taxes = total_received * (card_tax + platform_tax)
        net_profit = total_received - total_taxes - total_prod_cost - ads_cost

        return {
            "store_name": store_name,
            "paid_count": len(unique_paid),
            "cancelled_count": len(cancelled.drop_duplicates(subset=['Name'])),
            "pending_count": len(pending.drop_duplicates(subset=['Name'])),
            "total_items": total_items,
            "total_shipping": total_shipping,
            "total_received": total_received,
            "stats": stats,
            "tax_rate": card_tax + platform_tax,
            "total_taxes": total_taxes,
            "total_prod_cost": total_prod_cost,
            "ads_cost": ads_cost,
            "net_profit": net_profit,
            "analysis_date": datetime.now()
        }

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
            [f"(-) Taxas ({analysis_data['tax_rate']*100:.2f}%)", 
             f"R$ {analysis_data['total_taxes']:,.2f}"],
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


def main():
    """Função principal da aplicação web."""
    st.title("📊 Analisador de Vendas Multi-Loja")
    
    # Inicializar categorias na session_state se não existir
    if 'custom_categories' not in st.session_state:
        st.session_state.custom_categories = DEFAULT_CATEGORIAS_CONFIG.copy()
    
    # Criar instância do analyzer com categorias customizadas
    analyzer = VendasAnalyzerWeb(st.session_state.custom_categories)
    
    # Layout com abas
    tab1, tab2, tab3 = st.tabs(["📈 Análise", "⚙️ Categorias", "ℹ️ Ajuda"])
    
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
                
                # Seção 4: Taxas e Marketing
                st.markdown("### 4️⃣ Taxas e Marketing")
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    card_tax = st.number_input(
                        "% Taxa Cartão/Gateway",
                        value=DEFAULT_TAX_RATE,
                        min_value=0.0,
                        max_value=100.0,
                        step=0.01
                    )
                
                with col2:
                    platform_tax = st.number_input(
                        "% Taxa Plataforma",
                        value=0.0,
                        min_value=0.0,
                        max_value=100.0,
                        step=0.01
                    )
                
                with col3:
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
                            card_tax=card_tax/100,
                            platform_tax=platform_tax/100,
                            ads_cost=ads_cost
                        )
                        
                        # Salva na sessão
                        st.session_state.analysis_data = analysis_data
                        st.session_state.show_results = True
                        st.success("✅ Análise realizada com sucesso!")
                        
                    except Exception as e:
                        st.error(f"❌ Erro na análise: {str(e)}")
                
                # Exibir resultados se existirem
                if st.session_state.get('show_results'):
                    analysis = st.session_state.analysis_data
                    
                    st.markdown("---")
                    st.markdown("## 📊 Resultado da Análise")
                    
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
                    
                    # Resumo Financeiro
                    st.markdown("### Resumo Financeiro")
                    
                    financial_data = {
                        'Taxas': analysis['total_taxes'],
                        'Custo Produção': analysis['total_prod_cost'],
                        'Gasto ADS': analysis['ads_cost'],
                        'Lucro': analysis['net_profit']
                    }
                    
                    col1, col2 = st.columns([1, 1])
                    
                    with col1:
                        st.write("**Despesas:**")
                        for item, value in list(financial_data.items())[:-1]:
                            st.write(f"- {item}: R$ {value:,.2f}")
                    
                    with col2:
                        st.write("**Resultado:**")
                        color = '🟢' if analysis['net_profit'] > 0 else '🔴'
                        st.write(f"{color} **Lucro Líquido: R$ {analysis['net_profit']:,.2f}**")
                    
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
                    
                    # Botão de Download PDF
                    st.markdown("---")
                    
                    pdf_data = analyzer.generate_pdf(analysis)
                    filename = f"Relatorio_{store_name}_{analysis['analysis_date'].strftime('%Y%m%d_%H%M')}.pdf"
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.download_button(
                            label="📥 Baixar Relatório (PDF)",
                            data=pdf_data,
                            file_name=filename,
                            mime="application/pdf",
                            use_container_width=True
                        )
                    
                    with col2:
                        # Botão para fazer nova análise
                        if st.button("🔄 Nova Análise", use_container_width=True):
                            st.session_state.show_results = False
                            st.rerun()
            
            except Exception as e:
                st.error(f"❌ Erro ao processar arquivo: {str(e)}")
    
    with tab2:
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
    
    with tab3:
        st.markdown("""
        ## 📖 Como Usar
        
        ### Passo 1: Prepare seu arquivo CSV
        - Exporte seus dados de vendas da Shopify ou plataforma de vendas
        - O arquivo deve conter as colunas: Name, Financial Status, Lineitem name, Lineitem quantity, Lineitem price, Subtotal, Shipping, Total
        
        ### Passo 2: Carregue o arquivo
        - Clique em "Selecione seu arquivo CSV de vendas"
        - Escolha o arquivo da sua loja
        
        ### Passo 3: Configure os custos
        - Defina o custo de produção para cada categoria de produto
        - Configure o custo padrão para produtos não categorizados
        
        ### Passo 4: Configure as taxas
        - Taxa de Cartão/Gateway: taxa cobrada pelo gateway de pagamento (padrão: 4.99%)
        - Taxa de Plataforma: taxa da plataforma de vendas (exemplo: Shopify)
        - Gasto com ADS: total gasto em publicidade
        
        ### Passo 5: Gere a análise
        - Clique em "GERAR ANÁLISE COMPLETA"
        - Visualize os resultados na tela
        - Baixe o relatório em PDF se desejado
        
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
        
        ### Teste de Categorização:
        Use a ferramenta de teste na aba de categorias para verificar se seus produtos estão sendo categorizados corretamente.
        
        ## 💡 Dicas
        
        ✅ **Personalize as categorias** conforme seu catálogo de produtos
        ✅ Use nomes de produtos consistentes para melhor categorização
        ✅ Adicione variações de palavras (ex: "infantil, criança, baby, kids")
        ✅ Teste sempre novas categorias com a ferramenta de teste
        ✅ Atualize regularmente os custos de produção por categoria
        ✅ Revise as taxas de gateway conforme suas negociações
        ✅ Acompanhe o gasto com ADS para medir ROI
        """)


if __name__ == "__main__":
    # Inicializa estado da sessão
    if 'show_results' not in st.session_state:
        st.session_state.show_results = False
    if 'analysis_data' not in st.session_state:
        st.session_state.analysis_data = None
    
    main()
