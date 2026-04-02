# app.py
import streamlit as st
import pandas as pd
from analyzer import VendasAnalyzerWeb
from charts import create_abc_chart, create_heatmap_chart
from reports import generate_excel_report, generate_pdf_report

st.set_page_config(page_title="Dashboard Dropshipping", layout="wide")

@st.cache_data
def load_csv(file):
    for enc in ['utf-8', 'latin-1']:
        try:
            return pd.read_csv(file, encoding=enc)
        except Exception: pass
    return None

def main():
    st.title("📊 Dashboard Avançado de Vendas")
    
    analyzer = VendasAnalyzerWeb()
    
    uploaded_file = st.file_uploader("Suba seu CSV da Shopify", type=['csv'])
    
    if uploaded_file is not None:
        df = load_csv(uploaded_file)
        
        with st.sidebar:
            st.header("⚙️ Custos e Taxas")
            ads_cost = st.number_input("Gasto em ADS (R$)", value=0.0, step=10.0)
            traffic_mgr = st.number_input("Gestor de Tráfego (R$)", value=0.0, step=10.0)
            pix_tax = st.number_input("Taxa PIX (%)", value=0.0)
            boleto_tax = st.number_input("Boleto (R$ Fixo)", value=1.50)
            gateway_tax = st.number_input("Antifraude/Gateway (R$ Fixo)", value=1.50)
            plat_tax = st.number_input("Plataforma (%)", value=2.0)
            
            st.markdown("### Custos de Produto")
            costs_map = {}
            for cat in analyzer.get_categories_list():
                costs_map[cat] = st.number_input(f"Custo: {cat}", value=30.0 if "Oversized" in cat else 0.0)
            default_cost = st.number_input("Custo Padrão", value=0.0)
            
            st.markdown("### Taxas Cartão")
            card_taxes = {i: st.number_input(f"{i}x (%)", value=4.99 + (i*0.5)) for i in range(1, 4)} 

        if st.button("🚀 Gerar Relatório DRE"):
            analysis = analyzer.process_data(
                df, "Minha Loja", costs_map, default_cost, card_taxes, pix_tax, boleto_tax, gateway_tax, plat_tax/100, ads_cost, traffic_mgr
            )
            
            margem = (analysis['net_profit'] / analysis['total_received']) * 100 if analysis['total_received'] > 0 else 0
            if margem < 10: st.error(f"🚨 Atenção! Sua margem líquida é de {margem:.1f}%. Avalie cortar custos de tráfego ou aumentar o ticket.")
            else: st.success(f"✅ Operação Saudável! Margem de {margem:.1f}%.")
            
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Receita Total", f"R$ {analysis['total_received']:,.2f}")
            col2.metric("Lucro Líquido", f"R$ {analysis['net_profit']:,.2f}", f"{margem:.1f}%")
            col3.metric("CAC (Custo de Aquisição)", f"R$ {analysis['cac']:,.2f}")
            col4.metric("LTV (Valor do Cliente)", f"R$ {analysis['ltv']:,.2f}")
            
            st.markdown("---")
            
            tab1, tab2, tab3, tab4 = st.tabs(["🔥 Mapa de Calor", "📦 Curva ABC", "💳 Taxas", "🚚 Análise de Fretes"])
            
            with tab1:
                st.plotly_chart(create_heatmap_chart(analysis['heatmap']), use_container_width=True)
                
            with tab2:
                colA, colB = st.columns([2, 1])
                with colA: st.plotly_chart(create_abc_chart(analysis['abc_curve']), use_container_width=True)
                with colB:
                    st.write("**Curva A (80% da Receita):**")
                    for item in analysis['abc_curve'].get('A', [])[:5]:
                        st.write(f"- {item['product']} (R$ {item['revenue']:.2f})")
            
            with tab3:
                df_pagamentos = pd.DataFrame([
                    {"Método": k.upper(), "Pedidos": v['count'], "Total Faturado": f"R$ {v['total']:.2f}", "Taxas Descontadas": f"R$ {v['tax_amount']:.2f}"}
                    for k, v in analysis['payment_stats'].items()
                ])
                st.dataframe(df_pagamentos, hide_index=True, use_container_width=True)
                
            with tab4:
                st.markdown("### Resumo de Envios")
                f_stats = analysis['shipping_stats']
                
                col_f1, col_f2 = st.columns(2)
                col_f1.metric("Fretes Gratuitos (Pedidos)", f_stats['gratis']['count'], help="Pedidos em que o cliente não pagou o frete.")
                col_f2.metric("Receita Total C/ Frete", f"R$ {analysis['total_shipping']:,.2f}")
                
                st.markdown("#### Detalhamento de Custos Pagos (O que os clientes pagaram)")
                df_fretes = pd.DataFrame([
                    {"Método": "PAC", "Qtd Pedidos": f_stats['pac']['count'], "Valor Gasto (R$)": f"R$ {f_stats['pac']['value']:,.2f}"},
                    {"Método": "SEDEX", "Qtd Pedidos": f_stats['sedex']['count'], "Valor Gasto (R$)": f"R$ {f_stats['sedex']['value']:,.2f}"},
                    {"Método": "Transportadora", "Qtd Pedidos": f_stats['transportadora']['count'], "Valor Gasto (R$)": f"R$ {f_stats['transportadora']['value']:,.2f}"},
                ])
                st.dataframe(df_fretes, hide_index=True, use_container_width=True)

            st.markdown("---")
            st.markdown("### 📥 Download dos Relatórios Completos")
            colX, colY = st.columns(2)
            
            with colX:
                excel_data = generate_excel_report(analysis)
                st.download_button("📊 Baixar Planilha Excel Detalhada", data=excel_data, file_name="DRE_Completo.xlsx", use_container_width=True)
                
            with colY:
                pdf_data = generate_pdf_report(analysis)
                st.download_button("📄 Baixar PDF Executivo", data=pdf_data, file_name="Relatorio_Executivo.pdf", use_container_width=True)

if __name__ == "__main__":
    main()
