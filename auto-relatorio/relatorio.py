"""
auto-relatorio — Automacao de Leitura de Planilhas e Geracao de Relatorios
Autor: Victory Lesson | github.com/victorylesson
"""

import os
import sys
import json
import argparse
from datetime import datetime
from pathlib import Path

try:
    import pandas as pd
    import openpyxl
    from jinja2 import Template
except ImportError:
    print("[ERRO] Dependencias nao instaladas. Execute: pip install -r requirements.txt")
    sys.exit(1)


# ─────────────────────────────────────────
# LEITURA DA PLANILHA
# ─────────────────────────────────────────

def ler_planilha(caminho: str) -> pd.DataFrame:
    """Le .xlsx ou .csv e retorna um DataFrame."""
    ext = Path(caminho).suffix.lower()
    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(caminho)
    elif ext == ".csv":
        df = pd.read_csv(caminho, sep=None, engine="python")
    else:
        raise ValueError(f"Formato nao suportado: {ext}. Use .xlsx ou .csv")

    print(f"[OK] Planilha carregada: {len(df)} linhas, {len(df.columns)} colunas")
    return df


# ─────────────────────────────────────────
# ANALISE DOS DADOS
# ─────────────────────────────────────────

def analisar(df: pd.DataFrame) -> dict:
    """Gera estatisticas automaticas por tipo de coluna."""
    analise = {
        "gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "total_registros": len(df),
        "total_colunas": len(df.columns),
        "colunas": list(df.columns),
        "numericas": [],
        "texto": [],
        "alertas": [],
    }

    # Detecta nulos
    nulos = df.isnull().sum()
    for col, n in nulos.items():
        if n > 0:
            pct = round(n / len(df) * 100, 1)
            analise["alertas"].append(f"'{col}': {n} valores nulos ({pct}%)")

    # Estatisticas por coluna numerica
    for col in df.select_dtypes(include="number").columns:
        serie = df[col].dropna()
        if len(serie) == 0:
            continue
        analise["numericas"].append({
            "coluna": col,
            "soma": round(float(serie.sum()), 2),
            "media": round(float(serie.mean()), 2),
            "mediana": round(float(serie.median()), 2),
            "minimo": round(float(serie.min()), 2),
            "maximo": round(float(serie.max()), 2),
            "desvio": round(float(serie.std()), 2) if len(serie) > 1 else 0,
        })

    # Estatisticas por coluna de texto
    for col in df.select_dtypes(include="object").columns:
        serie = df[col].dropna()
        contagem = serie.value_counts()
        analise["texto"].append({
            "coluna": col,
            "unicos": int(serie.nunique()),
            "top5": contagem.head(5).to_dict(),
        })

    return analise


# ─────────────────────────────────────────
# GERACAO DO RELATORIO HTML
# ─────────────────────────────────────────

TEMPLATE_HTML = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Relatorio Automatico</title>
<style>
  :root {
    --bg: #0f0f0f;
    --surface: #1a1a1a;
    --border: #2a2a2a;
    --accent: #e8e8e8;
    --muted: #888;
    --green: #4ade80;
    --yellow: #facc15;
    --red: #f87171;
    --font: 'Courier New', monospace;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: var(--bg); color: var(--accent); font-family: var(--font); font-size: 13px; padding: 32px; }
  header { border-bottom: 1px solid var(--border); padding-bottom: 20px; margin-bottom: 28px; }
  h1 { font-size: 22px; letter-spacing: 4px; text-transform: uppercase; }
  .meta { color: var(--muted); font-size: 11px; margin-top: 6px; }
  .badge { display: inline-block; background: var(--surface); border: 1px solid var(--border);
           padding: 3px 10px; margin-right: 8px; font-size: 11px; letter-spacing: 1px; }
  .section { margin-bottom: 32px; }
  .section-title { font-size: 11px; letter-spacing: 3px; color: var(--muted); text-transform: uppercase;
                   border-bottom: 1px solid var(--border); padding-bottom: 6px; margin-bottom: 14px; }
  .cards { display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 12px; }
  .card { background: var(--surface); border: 1px solid var(--border); padding: 16px; }
  .card-label { font-size: 10px; color: var(--muted); letter-spacing: 2px; text-transform: uppercase; }
  .card-value { font-size: 24px; margin-top: 4px; }
  table { width: 100%; border-collapse: collapse; }
  th { background: var(--surface); color: var(--muted); font-size: 10px; letter-spacing: 2px;
       text-align: left; padding: 8px 12px; border-bottom: 1px solid var(--border); }
  td { padding: 8px 12px; border-bottom: 1px solid var(--border); font-size: 12px; }
  tr:hover td { background: var(--surface); }
  .alerta { background: var(--surface); border-left: 3px solid var(--yellow);
            padding: 8px 12px; margin-bottom: 6px; font-size: 12px; color: var(--yellow); }
  .pill { display: inline-block; background: var(--border); padding: 2px 8px;
          font-size: 10px; margin: 2px; }
  footer { margin-top: 40px; border-top: 1px solid var(--border); padding-top: 12px;
           font-size: 10px; color: var(--muted); letter-spacing: 2px; }
</style>
</head>
<body>
<header>
  <h1>Relatorio Automatico</h1>
  <div class="meta">
    <span class="badge">GERADO EM {{ dados.gerado_em }}</span>
    <span class="badge">{{ dados.total_registros }} REGISTROS</span>
    <span class="badge">{{ dados.total_colunas }} COLUNAS</span>
  </div>
</header>

{% if dados.alertas %}
<div class="section">
  <div class="section-title">Alertas de Qualidade</div>
  {% for a in dados.alertas %}
  <div class="alerta">⚠ {{ a }}</div>
  {% endfor %}
</div>
{% endif %}

{% if dados.numericas %}
<div class="section">
  <div class="section-title">Colunas Numericas</div>
  {% for col in dados.numericas %}
  <div style="margin-bottom:18px">
    <div style="font-size:11px;letter-spacing:2px;color:#aaa;margin-bottom:8px">{{ col.coluna | upper }}</div>
    <div class="cards">
      <div class="card"><div class="card-label">Soma</div><div class="card-value" style="color:var(--green)">{{ col.soma }}</div></div>
      <div class="card"><div class="card-label">Media</div><div class="card-value">{{ col.media }}</div></div>
      <div class="card"><div class="card-label">Mediana</div><div class="card-value">{{ col.mediana }}</div></div>
      <div class="card"><div class="card-label">Minimo</div><div class="card-value" style="color:var(--muted)">{{ col.minimo }}</div></div>
      <div class="card"><div class="card-label">Maximo</div><div class="card-value" style="color:var(--yellow)">{{ col.maximo }}</div></div>
      <div class="card"><div class="card-label">Desvio Padrao</div><div class="card-value" style="color:var(--muted)">{{ col.desvio }}</div></div>
    </div>
  </div>
  {% endfor %}
</div>
{% endif %}

{% if dados.texto %}
<div class="section">
  <div class="section-title">Colunas de Texto</div>
  <table>
    <tr>
      <th>Coluna</th>
      <th>Valores Unicos</th>
      <th>Top 5 Ocorrencias</th>
    </tr>
    {% for col in dados.texto %}
    <tr>
      <td>{{ col.coluna }}</td>
      <td>{{ col.unicos }}</td>
      <td>{% for k, v in col.top5.items() %}<span class="pill">{{ k }}: {{ v }}</span>{% endfor %}</td>
    </tr>
    {% endfor %}
  </table>
</div>
{% endif %}

<footer>
  AUTO-RELATORIO &nbsp;|&nbsp; github.com/victorylesson &nbsp;|&nbsp; GERADO AUTOMATICAMENTE EM {{ dados.gerado_em }}
</footer>
</body>
</html>
"""

def gerar_html(analise: dict, saida: str):
    tpl = Template(TEMPLATE_HTML)
    html = tpl.render(dados=analise)
    with open(saida, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[OK] Relatorio HTML gerado: {saida}")


# ─────────────────────────────────────────
# GERACAO DO RELATORIO JSON
# ─────────────────────────────────────────

def gerar_json(analise: dict, saida: str):
    with open(saida, "w", encoding="utf-8") as f:
        json.dump(analise, f, ensure_ascii=False, indent=2)
    print(f"[OK] Relatorio JSON gerado: {saida}")


# ─────────────────────────────────────────
# CLI
# ─────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="auto-relatorio: le planilha e gera relatorio automatico"
    )
    parser.add_argument("planilha", help="Caminho para o arquivo .xlsx ou .csv")
    parser.add_argument(
        "-o", "--output",
        default="relatorio",
        help="Nome base do arquivo de saida (sem extensao). Padrao: relatorio"
    )
    parser.add_argument(
        "-f", "--formato",
        choices=["html", "json", "ambos"],
        default="ambos",
        help="Formato de saida: html, json ou ambos. Padrao: ambos"
    )
    args = parser.parse_args()

    # Leitura
    df = ler_planilha(args.planilha)

    # Analise
    analise = analisar(df)

    # Saida
    pasta = Path(args.output).parent
    pasta.mkdir(parents=True, exist_ok=True)

    if args.formato in ("html", "ambos"):
        gerar_html(analise, f"{args.output}.html")

    if args.formato in ("json", "ambos"):
        gerar_json(analise, f"{args.output}.json")

    print(f"\n[PRONTO] {args.total_registros if hasattr(args,'total_registros') else analise['total_registros']} registros processados.")


if __name__ == "__main__":
    main()
