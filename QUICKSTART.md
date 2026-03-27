# ⚡ Guia Rápido de Início

## Teste Local em 5 Minutos

### 1️⃣ Instale Python (se não tiver)
Baixe em: https://www.python.org/downloads/
- ✅ Marque "Add Python to PATH"
- ✅ Clique "Install Now"

### 2️⃣ Abra PowerShell na Pasta do Projeto

```powershell
# Navegue até a pasta
cd "D:\Users\Downloads\Meus sites\onfight\Pedidos\teste"

# Instale dependências
pip install -r requirements.txt
```

### 3️⃣ rodalize a Aplicação Web

```powershell
streamlit run app_web.py
```

**Pronto!** Abre automaticamente em http://localhost:8501

---

## 🌐 Deploy Grátis em 5 Minutos (Streamlit Cloud)

### Passo 1: GitHub
1. Acesse https://github.com/new
2. Nome do repo: `vendas-analyzer`
3. Clique **Create repository**

### Passo 2: Upload dos Arquivos
1. Clique em "uploading an existing file"
2. Arraste:
   - `app_web.py`
   - `requirements.txt`
3. Clique "Commit changes"

### Passo 3: Deploy
1. Acesse https://share.streamlit.io
2. **New app**
3. Selecione seu repositório
4. Clique **Deploy**

**Pronto! Sua app está online!** 🎉

---

## 📁 Verificar Arquivos

Você deve ter esses arquivos na pasta:

```
✅ app_web.py          (Nova aplicação web)
✅ teste.py            (Versão desktop original)
✅ requirements.txt    (Dependências)
✅ README.md          (Documentação)
✅ Procfile           (Para deploy)
✅ .gitignore         (Para GitHub)
```

---

## ⚠️ Problemas Comuns

### "pip: comando não encontrado"
**Solução:** Python pode não estar em PATH
```powershell
# Use py em vez de python
py -m pip install -r requirements.txt
py -m streamlit run app_web.py
```

### "ModuleNotFoundError: No module named 'streamlit'"
```powershell
# Reinstale dependências
pip install --upgrade -r requirements.txt
```

### Porta 8501 já em uso
```powershell
streamlit run app_web.py --server.port 8502
```

---

## 📚 Próximas Ações

✅ Ao lado, você encontra:
- **README.md** - Documentação completa
- **app_web.py** - Aplicação web
- **teste.py** - Versão desktop

🚀 Escolha:
1. **Teste local** → Execute `streamlit run app_web.py`
2. **Deploy gratuito** → Siga o passo 3 acima
3. **Compartilhe** → Envie link para clientes usarem!

---

**Dúvidas?** Consulte o README.md para documentação completa!
