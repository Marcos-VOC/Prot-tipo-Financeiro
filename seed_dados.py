# ============================================================
#  seed_dados.py
#  Popula o banco com os dados da planilha Santos & Associados
#  Execute UMA VEZ: python seed_dados.py
# ============================================================

import sqlite3, os, sys

DB_PATH = os.path.join(os.path.dirname(__file__), "financeiro.db")

# Garantir que o banco e tabela existam
conn = sqlite3.connect(DB_PATH)
conn.execute("""
    CREATE TABLE IF NOT EXISTS movimentacoes (
        id        INTEGER PRIMARY KEY AUTOINCREMENT,
        tipo      TEXT NOT NULL,
        categoria TEXT NOT NULL,
        descricao TEXT NOT NULL,
        valor     REAL NOT NULL,
        data      TEXT NOT NULL
    )
""")

# Verifica se já tem dados
qtd = conn.execute("SELECT COUNT(*) FROM movimentacoes").fetchone()[0]
if qtd > 0:
    print(f"⚠  Banco já contém {qtd} registros. Seed ignorado.")
    print("   Para recriar: delete o arquivo financeiro.db e rode novamente.")
    conn.close()
    sys.exit(0)

# ── Dados da planilha (Jan–Dez 2025) ──────────────────────────

# Receitas mensais (dia 5 de cada mês para simular recebimento)
receitas = [
    # (categoria, [jan,fev,mar,abr,mai,jun,jul,ago,set,out,nov,dez])
    ("Honorários – Consultoria",  [18500,21000,17800,23500,19200,22000,20500,24000,18900,21500,23000,25000]),
    ("Honorários – Contencioso",  [32000,28500,35000,31000,29800,34000,33000,36000,30000,32500,35000,38000]),
    ("Honorários – Trabalhista",  [14200,15000,13500,16000,14800,15500,14500,16500,13800,15200,16000,17000]),
    ("Acordos e Êxitos",          [ 8000,12000, 5000,15000,10000, 9500,11000,13000, 7500,10500,12000,14000]),
    ("Pareceres e Assessoria",    [ 6500, 7200, 6800, 7500, 6900, 7100, 6800, 7500, 6600, 7000, 7300, 7800]),
]

# Despesas mensais (dia 10 de cada mês)
despesas = [
    ("Salários e Encargos",            [22000,22000,22000,22000,22000,22000,22000,22000,22000,22000,22000,22000]),
    ("Aluguel e Condomínio",           [ 8500, 8500, 8500, 8500, 8500, 8500, 8500, 8500, 8500, 8500, 8500, 8500]),
    ("Softwares Jurídicos",            [ 1800, 1800, 1800, 1800, 1800, 1800, 1800, 1800, 1800, 1800, 1800, 1800]),
    ("Marketing e Publicidade",        [ 2500, 3200, 1800, 4000, 2800, 3100, 2700, 3500, 2200, 3000, 3300, 4000]),
    ("Custas e Despesas Processuais",  [ 3200, 4100, 2900, 5200, 3800, 4500, 4000, 5000, 3500, 4200, 4600, 5500]),
    ("Treinamento e Capacitação",      [  800,  600, 1200,  500, 1000,  900,  700,  800, 1100,  600,  900, 1200]),
    ("Despesas Administrativas",       [ 1500, 1700, 1400, 1600, 1550, 1650, 1500, 1700, 1450, 1600, 1650, 1800]),
]

meses = ["01","02","03","04","05","06","07","08","09","10","11","12"]
registros = []

for cat, valores in receitas:
    for i, v in enumerate(valores):
        if v > 0:
            registros.append((
                "Entrada", cat,
                f"{cat} – {['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'][i]}/2025",
                v,
                f"2025-{meses[i]}-05"
            ))

for cat, valores in despesas:
    for i, v in enumerate(valores):
        if v > 0:
            registros.append((
                "Saida", cat,
                f"{cat} – {['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'][i]}/2025",
                v,
                f"2025-{meses[i]}-10"
            ))

conn.executemany(
    "INSERT INTO movimentacoes (tipo,categoria,descricao,valor,data) VALUES (?,?,?,?,?)",
    registros
)
conn.commit()
conn.close()

print(f"✅  {len(registros)} registros inseridos com sucesso!")
print("   Agora rode: python app.py")
