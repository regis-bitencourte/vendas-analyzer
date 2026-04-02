# app.py
import streamlit as st
import pandas as pd
import os
import json
from datetime import datetime

# Importando os módulos do seu projeto
from analyzer import VendasAnalyzerWeb, DEFAULT_CATEGORIAS_CONFIG
from charts import create_abc_chart, create_heatmap_chart
from reports import generate_excel_report, generate_pdf_report

st.set_page_config(page_title="Dashboard Dropshipping", page_icon="📊", layout="wide")

HISTORY_FILE = "analise_history.json"

# --- FUNÇÕES DE HISTÓRICO ---
def save_analysis_history(analysis_data: dict) -> None:
    history = []
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                history = json.load(f)
        except: pass
    
    analysis_entry = {
        'timestamp': datetime.now().isoformat(),
        'store': analysis_data.get('store_name', 'Loja'),
        'paid_orders': analysis_data.get('paid_count', 0),
        'revenue': float(analysis_data.get('total_received', 0)),
        'profit': float(analysis_data.get('net_profit', 0)),
        'shipping': float(analysis_data.get('total_shipping', 0))
    }
    history.append(analysis_entry)
    if len(history) > 100: history = history[-100:]
    with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
        json.dump(history, f, indent=2, ensure_ascii=False)

def load_analysis_history() -> list:
    if not os.path.exists(HISTORY_FILE): return []
    try:
        with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except: return []

# --- CACHE DO CSV ---
@st.cache_data
def load_csv(file):
    for enc in ['utf-8', 'latin-1', 'iso-8859-1']:
        try:
            return pd.read_csv(file, encoding=enc)
        except UnicodeDecodeError: continue
    return None

def main():
    st.title("📊 Dashboard de Vendas Multi-Loja")
    
    # Inicialização do State
    if 'custom_categories' not in st.session_state:
        st.session_state.custom_categories = DEFAULT_CATEGORIAS_CONFIG.copy()
    if 'show_results' not in st.session_state:
        st.session_state.show_results = False
    if 'analysis_data' not in st.session_state:
        st.session_state.analysis_data = None
        
    analyzer = VendasAnalyzerWeb(st.session_state.custom_categories)
    
    # As 4 Abas Originais
    tab1, tab2, tab3, tab4 = st.tabs(["📈 Análise Básica", "🔍 Análises Avançadas", "⚙️ Categorias", "ℹ️ Ajuda"])
    
    # ==========================================
    # ABA 1: ANÁLISE BÁSICA
    # ==========================================
    with tab1:
        st.markdown("### 1️⃣ Informações e Arquivo")
        colA, colB = st.columns([1, 2])
        with colA: store_name = st.text_input("Nome da Loja", value="Minha Loja")
        with colB: uploaded_file = st.file_uploader("Selecione seu arquivo CSV da Shopify", type=['csv'])
        
        if uploaded_file is not None:
            df = load_csv(uploaded_file)
            st.success(f"✅ Arquivo carregado com sucesso ({len(df)} linhas)")
            
            with st.sidebar:
                st.header("⚙️ Configurações de Custos")
                
                st.markdown("#### Custos de Marketing e Taxas")
                ads_cost = st.number_input("Gasto em ADS (R$)", value=0.0, step=10.0)
                traffic_mgr = st.number_input("Gestor de Tráfego (R$)", value=0.0, step=10.0)
                pix_tax = st.number_input("Taxa PIX (%)", value=0.0)
                boleto_tax = st.number_input("Boleto (R$ Fixo)", value=1.50)
                gateway_tax = st.number_input("Antifraude/Gateway (R$ Fixo)", value=1.50)
                plat_tax = st.number_input("Plataforma (%)", value=2.0)
                
                st.markdown("#### Taxas de Cartão")
                with st.expander("💳 Configurar Parcelas"):
                    card_taxes = {i: st.number_input(f"{i}x (%)", value=4.99 + (i*0.5)) for i in range(1, 13)}
                
                st.markdown("#### Custos de Produção")
                costs_map = {}
                for cat in analyzer.get_categories_list():
                    costs_map[cat] = st.number_input(f"Custo: {cat}", value=30.0 if "Oversized" in cat else 0.0)
                default_cost = st.number_input("Custo Padrão (Outros)", value=0.0)
                
            if st.button("🚀 GERAR ANÁLISE COMPLETA", use_container_width=True):
                try:
                    analysis = analyzer.process_data(
                        df, store_name, costs_map, default_cost, card_taxes, 
                        pix_tax, boleto_tax, gateway_tax, plat_tax/100, ads_cost, traffic_mgr
                    )
                    st.session_state.analysis_data = analysis
                    st.session_state.df_analysis = df
                    st.session_state.show_results = True
                    save_analysis_history(analysis)
                except Exception as e:
                    st.error(f"Erro ao processar: {e}")
            
            # --- MOSTRAR RESULTADOS ---
            if st.session_state.show_results:
                analysis = st.session_state.analysis_data
                
                st.markdown("---")
                st.markdown("## 📊 Resultado Financeiro")
                
                margem = (analysis['net_profit'] / analysis['total_received']) * 100 if analysis['total_received'] > 0 else 0
                if margem < 10: st.warning(f"🚨 Atenção: Margem líquida de {margem:.1f}%.")
                else: st.success(f"✅ Margem saudável de {margem:.1f}%.")
                
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Pedidos Pagos", analysis.get('paid_count', 0))
                col2.metric("Receita Total", f"R$ {analysis.get('total_received', 0):,.2f}")
                col3.metric("Lucro Líquido", f"R$ {analysis.get('net_profit', 0):,.2f}", f"{margem:.1f}%")
                col4.metric("LTV (Valor do Cliente)", f"R$ {analysis.get('ltv', 0):,.2f}")
                
                st.markdown("### 💳 Taxas por Pagamento")
                if 'payment_stats' in analysis:
                    df_pag = pd.DataFrame([
                        {"Método": k.upper(), "Pedidos": v['count'], "Faturado": f"R$ {v['total']:.2f}", "Descontado": f"R$ {v['tax_amount']:.2f}"}
                        for k, v in analysis['payment_stats'].items() if v['count'] > 0
                    ])
                    st.dataframe(df_pag, hide_index=True, use_container_width=True)
                
                st.markdown("### 🚚 Resumo de Fretes Inteligente")
                if 'shipping_stats' in analysis:
                    f_stats = analysis['shipping_stats']
                    c1, c2 = st.columns(2)
                    c1.metric("🎁 Fretes Gratuitos", f"{f_stats.get('gratis', {}).get('count', 0)} Pedidos")
                    c2.metric("Total Cobrado dos Clientes", f"R$ {analysis.get('total_shipping', 0):,.2f}")
                    
                    df_fretes = pd.DataFrame([
                        {"Método": "PAC", "Qtd": f_stats.get('pac', {}).get('count', 0), "Valor": f"R$ {f_stats.get('pac', {}).get('value', 0):,.2f}"},
                        {"Método": "SEDEX", "Qtd": f_stats.get('sedex', {}).get('count', 0), "Valor": f"R$ {f_stats.get('sedex', {}).get('value', 0):,.2f}"},
                        {"Método": "Transportadora", "Qtd": f_stats.get('transportadora', {}).get('count', 0), "Valor": f"R$ {f_stats.get('transportadora', {}).get('value', 0):,.2f}"},
                    ])
                    st.dataframe(df_fretes, hide_index=True, use_container_width=True)

                st.markdown("---")
                st.markdown("### 📥 Downloads")
                colX, colY, colZ = st.columns(3)
                with colX:
                    excel_data = generate_excel_report(analysis)
                    st.download_button("📊 Excel Completo", data=excel_data, file_name="Relatorio.xlsx", use_container_width=True)
                with colY:
                    pdf_data = generate_pdf_report(analysis)
                    st.download_button("📄 PDF Executivo", data=pdf_data, file_name="Relatorio.pdf", use_container_width=True)
                with colZ:
                    if st.button("🔄 Nova Análise", use_container_width=True):
                        st.session_state.show_results = False
                        st.rerun()

    # ==========================================
    # ABA 2: ANÁLISES AVANÇADAS
    # ==========================================
    with tab2:
        st.markdown("## 🔍 Análises Avançadas")
        if not st.session_state.get('show_results'):
            st.info("⏳ Execute uma análise na aba '📈 Análise Básica' primeiro.")
        else:
            analysis = st.session_state.analysis_data
            sub1, sub2, sub3, sub4 = st.tabs(["🔥 Mapa e Curva ABC", "👥 Clientes & LTV", "📦 Fulfillment", "📚 Histórico"])
            
            with sub1:
                colA, colB = st.columns(2)
                with colA:
                    st.markdown("### 🔥 Mapa de Calor")
                    if 'heatmap' in analysis and analysis['heatmap']:
                        st.plotly_chart(create_heatmap_chart(analysis['heatmap']), use_container_width=True)
                with colB:
                    st.markdown("### 📦 Curva ABC")
                    if 'abc_curve' in analysis and analysis['abc_curve']:
                        st.plotly_chart(create_abc_chart(analysis['abc_curve']), use_container_width=True)
            
            with sub2:
                st.markdown("### 👥 Métricas de Clientes")
                col1, col2 = st.columns(2)
                col1.metric("Custo de Aquisição (CAC)", f"R$ {analysis.get('cac', 0):,.2f}")
                col2.metric("Ticket Médio", f"R$ {analysis.get('avg_ticket', 0):,.2f}")
                
                # Se você mantiver o _analyze_repeat_customers no seu analyzer, ele mostrará aqui:
                if 'repeat_customers' in analysis:
                    st.markdown("#### Top Clientes Recorrentes")
                    top = analysis['repeat_customers'].get('top_customers', [])
                    if top: st.dataframe(pd.DataFrame(top), hide_index=True, use_container_width=True)
            
            with sub3:
                st.markdown("### 📦 Status de Entrega")
                if 'fulfillment' in analysis:
                    f = analysis['fulfillment']
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Entregues", f.get('fulfilled', 0))
                    col2.metric("Não Entregues", f.get('unfulfilled', 0))
                    col3.metric("Cancelados", f.get('cancelled', 0))
                else:
                    st.info("A função de Fulfillment não foi ativada nesta análise.")
            
            with sub4:
                st.markdown("### 📋 Histórico de Análises")
                hist = load_analysis_history()
                if hist:
                    df_hist = pd.DataFrame(hist).sort_values('timestamp', ascending=False)
                    st.dataframe(df_hist, hide_index=True, use_container_width=True)
                    if st.button("🗑️ Limpar Histórico"):
                        if os.path.exists(HISTORY_FILE): os.remove(HISTORY_FILE)
                        st.rerun()

    # ==========================================
    # ABA 3: CATEGORIAS (Gerenciamento)
    # ==========================================
    with tab3:
        st.markdown("## ⚙️ Gerenciamento de Categorias")
        st.markdown("Personalize as palavras-chave para que o sistema identifique seus produtos automaticamente.")
        
        st.markdown("### ➕ Adicionar Categoria")
        c1, c2, c3 = st.columns([2, 3, 1])
        with c1: new_cat = st.text_input("Nome (Ex: Camiseta)")
        with c2: new_kw = st.text_input("Palavras-chave separadas por vírgula")
        with c3:
            if st.button("Adicionar", use_container_width=True):
                if new_cat and new_kw:
                    st.session_state.custom_categories[new_cat] = [k.strip().lower() for k in new_kw.split(',')]
                    st.success("Categoria Adicionada!")
                    st.rerun()
        
        st.markdown("---")
        st.markdown("### 📋 Categorias Atuais")
        for cat_name, kws in list(st.session_state.custom_categories.items()):
            with st.expander(f"📦 {cat_name}"):
                col1, col2, col3 = st.columns([2, 3, 1])
                with col1: edit_cat = st.text_input("Nome", value=cat_name, key=f"nm_{cat_name}")
                with col2: edit_kw = st.text_input("Palavras-chave", value=", ".join(kws), key=f"kw_{cat_name}")
                with col3:
                    if st.button("🗑️ Deletar", key=f"del_{cat_name}"):
                        del st.session_state.custom_categories[cat_name]
                        st.rerun()
                    if st.button("💾 Salvar", key=f"sav_{cat_name}"):
                        del st.session_state.custom_categories[cat_name]
                        st.session_state.custom_categories[edit_cat] = [k.strip().lower() for k in edit_kw.split(',')]
                        st.rerun()
                        
        if st.button("🔄 Resetar para Padrão de Fábrica"):
            st.session_state.custom_categories = DEFAULT_CATEGORIAS_CONFIG.copy()
            st.rerun()

    # ==========================================
    # ABA 4: AJUDA
    # ==========================================
    with tab4:
        st.markdown("""
        ## 📖 Guia de Uso do Dashboard
        1. **Exporte o CSV da Shopify** com todas as colunas de métricas.
        2. Carregue na **Aba Básica**.
        3. Ajuste as **Taxas de Cartão**, **Gateway**, **Pix** e **Boletos**. Boletos e Gateway agora descontam um **valor fixo em Reais**.
        4. O sistema usa as palavras-chave configuradas na aba de **Categorias** para encontrar o custo de cada item vendido.
        5. Faça o download dos Relatórios em PDF/Excel.
        """)

if __name__ == "__main__":
    main()
