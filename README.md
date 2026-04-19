# ⚖ Santos & Associados — Sistema de Controle Financeiro

Sistema web em Python/Flask para controle financeiro do escritório de advocacia.

---

## 📁 Estrutura do Projeto

```
santos_financeiro/
├── app.py                  ← Backend Flask (rotas, banco de dados)
├── seed_dados.py           ← Importa os dados da planilha 2025
├── requirements.txt        ← Dependências Python
├── financeiro.db           ← Banco SQLite (criado automaticamente)
└── templates/
    ├── base.html           ← Layout base com sidebar
    ├── dashboard.html      ← Página principal com gráfico e resumo
    ├── movimentacoes.html  ← Listagem com filtros
    └── nova.html           ← Formulário de cadastro
```

---

## 🚀 Como Rodar (Passo a Passo)

### 1. Instale o Python
Certifique-se de ter o **Python 3.8+** instalado.
Verifique com: `python --version`

### 2. Crie e ative um ambiente virtual (recomendado)
```bash
# Windows
python -m venv venv
venv\Scripts\activate

# Mac / Linux
python -m venv venv
source venv/bin/activate
```

### 3. Instale as dependências
```bash
pip install -r requirements.txt
```

### 4. Popule o banco com os dados da planilha 2025
```bash
python seed_dados.py
```
> Execute **apenas uma vez**. Insere os 144 registros da planilha Santos & Associados.

### 5. Inicie o servidor
```bash
python app.py
```

### 6. Acesse no navegador
```
http://127.0.0.1:5000
```

---

## 🖥 Telas do Sistema

| Tela | URL | Descrição |
|------|-----|-----------|
| Dashboard | `/` | Resumo do mês + gráfico anual |
| Movimentações | `/movimentacoes` | Listagem com filtros |
| Nova movimentação | `/nova` | Cadastro de entrada ou saída |

---

## 💡 Funcionalidades

- ✅ Cadastro de entradas e saídas com categoria
- ✅ Dashboard com total recebido, total gasto e lucro líquido
- ✅ Gráfico de barras Entradas × Saídas (12 meses)
- ✅ Top categorias de receita e despesa
- ✅ Filtro por mês e ano
- ✅ Busca por descrição ou categoria
- ✅ Exclusão de movimentações
- ✅ Margem líquida em %
- ✅ Banco de dados local SQLite (sem servidor externo)

---

## 🔄 Para Reiniciar os Dados

```bash
# Delete o banco e re-importe
del financeiro.db        # Windows
rm financeiro.db         # Mac/Linux
python seed_dados.py
```

---

## 📌 Observações Técnicas

- O banco `financeiro.db` é criado automaticamente na primeira execução
- Todos os dados ficam armazenados localmente no arquivo SQLite
- Para produção futura, basta trocar SQLite por PostgreSQL/MySQL
- O sistema usa Bootstrap 5 via CDN (requer internet no primeiro acesso)
