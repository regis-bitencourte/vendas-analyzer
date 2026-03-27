# 📊 Analisador de Vendas Web - Guia Completo

Uma aplicação web profissional para análise de dados de vendas com geração automática de relatórios em PDF.

## 🎯 Características

✅ **Interface Web Intuitiva** - Sem necessidade de instalar Python
✅ **Upload de CSV** - Compatível com Shopify e outras plataformas
✅ **Análise Automática** - Categoriza produtos e calcula métricas
✅ **Relatórios PDF** - Exportação profissional de resultados
✅ **Responsivo** - Funciona em PC, tablet e mobile
✅ **Deploy Fácil** - Hospedagem gratuita disponível

---

## 🚀 Início Rápido (Para Desenvolvedores)

### Opção 1: Executar Localmente

#### Pré-requisitos
- Python 3.8 ou superior instalado

#### 1. Clone ou baixe os arquivos

```bash
cd seu-diretorio
```

#### 2. Crie um ambiente virtual (opcional, mas recomendado)

**Windows:**
```bash
python -m venv venv
venv\Scripts\activate
```

**Mac/Linux:**
```bash
python -m venv venv
source venv/bin/activate
```

#### 3. Instale as dependências

```bash
pip install -r requirements.txt
```

#### 4. Execute a aplicação

```bash
streamlit run app_web.py
```

A aplicação abrirá automaticamente em http://localhost:8501

---

## 🌐 Deploy na Nuvem (Para Usar sem Python Instalado)

### Opção 1: Streamlit Cloud (RECOMENDADO - GRATUITO)

**Melhor opção para começar - Totalmente gratuito!**

#### Passo 1: Prepare o repositório GitHub

1. Crie uma conta em https://github.com (se não tiver)
2. Crie um novo repositório chamado `vendas-analyzer`
3. Upload dos arquivos:
   - `app_web.py`
   - `requirements.txt`
   - `README.md` (este arquivo)

#### Passo 2: Deploy no Streamlit Cloud

1. Acesse https://share.streamlit.io
2. Clique em "New app"
3. Conecte sua conta GitHub
4. Selecione:
   - Repository: `seu-usuario/vendas-analyzer`
   - Branch: `main`
   - File path: `app_web.py`

5. Clique em "Deploy"

**Pronto!** Sua app estará online em 2-3 minutos.

**Vantagens:**
- ✅ Totalmente gratuito
- ✅ Domínio próprio (seu-app.streamlit.app)
- ✅ Deploy automático ao fazer push no GitHub
- ✅ Sem limite de apps

---

### Opção 2: Heroku (Pago, mas com plano free limitado)

#### Passo 1: Criar conta Heroku

1. Acesse https://www.heroku.com
2. Crie uma conta gratuita

#### Passo 2: Instale Heroku CLI

- **Windows**: https://devcenter.heroku.com/articles/heroku-cli
- **Mac/Linux**: `curl https://cli-assets.heroku.com/install.sh | sh`

#### Passo 3: Prepare arquivos adicionais

Crie `Procfile` (sem extensão) na raiz do projeto:
```
web: streamlit run app_web.py --logger.level=error
```

Crie `.gitignore`:
```
venv/
__pycache__/
*.pyc
.DS_Store
```

#### Passo 4: Deploy

```bash
heroku login
heroku create seu-app-name
git push heroku main
```

---

### Opção 3: Railway.app (Novo, muito fácil)

1. Acesse https://railway.app
2. Clique em "Deploy Now"
3. Conecte seu repositório GitHub
4. Pronto! Deploy automático

---

### Opção 4: Google Cloud Run (Com crédito gratuito)

1. Acesse https://cloud.google.com
2. Faz setup do Cloud Run
3. Deploy da imagem Docker

**Mais técnico, mas com crédito free $300**

---

## 📋 Preparando seu CSV

Seu arquivo deve ter essas colunas (nomes exatos):

| Coluna | Exemplo | Descrição |
|--------|---------|-----------|
| `Name` | Order #1001 | ID/Nome do pedido |
| `Financial Status` | paid | paid, cancelled, pending |
| `Lineitem name` | Camiseta Oversized | Nome do produto |
| `Lineitem quantity` | 2 | Quantidade |
| `Lineitem price` | 89.90 | Preço unitário |
| `Subtotal` | 179.80 | Valor dos itens |
| `Shipping` | 19.90 | Frete |
| `Total` | 199.70 | Total do pedido |

### Exportando da Shopify

1. Vá para **Pedidos**
2. Clique em **Exportar** (canto superior direito)
3. Download do CSV
4. Upload na aplicação

---

## 🎨 Customização

### Adicionar Novas Categorias

Edite em `app_web.py`:

```python
CATEGORIAS_CONFIG = {
    "Oversized": ["oversized"],
    "Short 2 em 1": ["short", "2 em 1", "2em1"],
    "Dryfit": ["dryfit", "dry fit"],
    "Moletom": ["moletom", "hoodie"],
    "Calça": ["calça", "calca", "pants"],
    "Combo": ["combo", "kit"],
    "Sua Categoria": ["palavra1", "palavra2"]  # ← Adicione aqui
}
```

### Alterar Cores e Estilo

Os estilos CSS estão no início da function `main()`:

```python
st.markdown("""
<style>
    .main {
        padding: 0rem 1rem;
    }
    h1 {
        color: #1f77b4;  # ← Mude a cor aqui
    }
</style>
""", unsafe_allow_html=True)
```

---

## 🛠️ Troubleshooting

### Erro: "ModuleNotFoundError"

```bash
pip install -r requirements.txt
```

### Erro: "File not found"

Certifique-se que `app_web.py` está no mesmo diretório onde você roda `streamlit run`

### CSV não carrega

- Verifique o encoding (UTF-8, Latin-1 ou ISO-8859-1)
- Confirme que as colunas têm nomes exatos
- Veja se há caracteres especiais problemáticos

### Aplicação lenta

- Envie arquivos CSV menores (< 50MB)
- Reduza a frequência de atualização
- Limpe cache do navegador

---

## 📚 Estrutura de Arquivos

```
seu-projeto/
├── app_web.py           # Aplicação principal
├── teste.py             # Versão desktop (opcional)
├── requirements.txt     # Dependências Python
├── Procfile            # Para Heroku (se usar)
├── .gitignore          # Arquivos ignorados
└── README.md           # Este arquivo
```

---

## 🔐 Segurança

### Dados Privados

- ✅ Arquivos são processados localmente (Streamlit Cloud)
- ✅ Nenhum dado é armazenado no servidor
- ✅ Nenhuma conta necessária
- ✅ Seguro para dados sensíveis de negócio

### Para Mais Privacidade

Execute localmente:
```bash
streamlit run app_web.py
```

---

## 💰 Custos

### Streamlit Cloud (RECOMENDADO)
- ✅ **GRATUITO** para até 3 apps
- Múltiplos viewers simultâneos ilimitados
- 1GB de memória, 1 vCPU
- Perfeito para pequenas/médias lojas

### Heroku (Descontinuado free tier)
- 💰 Começa em $6/mês
- Mais caro que outros opções

### Railway.app
- 💰 Desde $5/mês
- ótima performance

### Google Cloud Run
- 💰 Crédito free $300
- Então $0.00001667 por vCPU-segundo (praticamente grátis!)

---

## 📞 Suporte

### Documentação Oficial
- Streamlit: https://docs.streamlit.io
- Pandas: https://pandas.pydata.org/docs
- ReportLab: https://www.reportlab.com/docs/reportlab-userguide.pdf

### Comunidade
- Streamlit Forum: https://discuss.streamlit.io
- Stack Overflow: tag `streamlit`

---

## 📈 Próximos Passos (Evolutivo)

- [ ] Adicionar gráficos (Plotly, Matplotlib)
- [ ] Histórico de análises (banco de dados)
- [ ] Comparação de períodos
- [ ] Integração direta com Shopify API
- [ ] Notificações por email
- [ ] App mobile nativa

---

## 📄 Licença

Uso livre para seus negócios. Sinta-se livre para customizar!

---

**Última atualização:** Março 2026
**Versão:** 1.0.0
