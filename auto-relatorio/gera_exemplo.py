"""
gera_exemplo.py — Cria uma planilha de vendas ficticia para demonstracao
Uso: python gera_exemplo.py
"""

import pandas as pd
import random
from datetime import datetime, timedelta
from pathlib import Path

random.seed(42)

produtos = ["Notebook", "Monitor", "Teclado", "Mouse", "Headset", "Webcam", "SSD", "Memoria RAM"]
vendedores = ["Ana Lima", "Bruno Souza", "Carla Dias", "Diego Melo", "Erica Costa"]
regioes = ["Norte", "Sul", "Leste", "Oeste", "Centro"]
status = ["Concluido", "Concluido", "Concluido", "Cancelado", "Pendente"]

linhas = []
data_inicio = datetime(2024, 1, 1)

for i in range(120):
    data = data_inicio + timedelta(days=random.randint(0, 364))
    qtd = random.randint(1, 10)
    preco = round(random.uniform(50, 3000), 2)
    linhas.append({
        "ID": f"VND-{i+1:04d}",
        "Data": data.strftime("%d/%m/%Y"),
        "Produto": random.choice(produtos),
        "Vendedor": random.choice(vendedores),
        "Regiao": random.choice(regioes),
        "Quantidade": qtd,
        "Preco_Unitario": preco,
        "Total": round(qtd * preco, 2),
        "Status": random.choice(status),
        "Avaliacao": random.randint(1, 5) if random.random() > 0.2 else None,
    })

df = pd.DataFrame(linhas)
caminho = Path("dados/vendas_exemplo.xlsx")
caminho.parent.mkdir(exist_ok=True)
df.to_excel(caminho, index=False)
print(f"[OK] Planilha criada: {caminho} ({len(df)} registros)")
