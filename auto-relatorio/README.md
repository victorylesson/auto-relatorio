# auto-relatorio

Automacao de leitura de planilhas (.xlsx / .csv) com geracao de relatorios HTML e JSON.

Desenvolvido como parte do portfolio de automacao de processos — sem interface, sem clique manual, zero configuracao extra.

---

## O que faz

- Le qualquer planilha `.xlsx` ou `.csv`
- Detecta colunas numericas e de texto automaticamente
- Calcula estatisticas: soma, media, mediana, minimo, maximo, desvio padrao
- Identifica valores nulos e exibe alertas de qualidade
- Gera relatorio HTML estilizado (pronto para abrir no navegador)
- Gera relatorio JSON estruturado (pronto para integrar com outros sistemas)

---

## Instalacao

```bash
git clone https://github.com/victorylesson/auto-relatorio
cd auto-relatorio
pip install -r requirements.txt
```

---

## Uso

### Geracao basica (HTML + JSON)
```bash
python relatorio.py dados/vendas.xlsx
```

### Escolher formato de saida
```bash
python relatorio.py dados/vendas.xlsx -f html
python relatorio.py dados/vendas.xlsx -f json
python relatorio.py dados/vendas.xlsx -f ambos
```

### Definir nome do arquivo de saida
```bash
python relatorio.py dados/vendas.xlsx -o saidas/relatorio_janeiro
```

---

## Demonstracao rapida

```bash
# Gera uma planilha de vendas ficticia com 120 registros
python gera_exemplo.py

# Gera o relatorio a partir dela
python relatorio.py dados/vendas_exemplo.xlsx -o saidas/demo
```

Abre `saidas/demo.html` no navegador para ver o resultado.

---

## Estrutura do projeto

```
auto-relatorio/
├── relatorio.py        # Script principal
├── gera_exemplo.py     # Gerador de planilha demo
├── requirements.txt    # Dependencias
├── dados/              # Coloque suas planilhas aqui
└── saidas/             # Relatorios gerados
```

---

## Exemplo de saida JSON

```json
{
  "gerado_em": "15/05/2025 14:32",
  "total_registros": 120,
  "total_colunas": 10,
  "alertas": [
    "'Avaliacao': 23 valores nulos (19.2%)"
  ],
  "numericas": [
    {
      "coluna": "Total",
      "soma": 184732.50,
      "media": 1539.44,
      "mediana": 1287.00,
      "minimo": 52.00,
      "maximo": 29850.00,
      "desvio": 1102.33
    }
  ]
}
```

---

## Stack

- **Python 3.10+**
- **pandas** — leitura e analise de dados
- **openpyxl** — suporte a .xlsx
- **jinja2** — template do relatorio HTML

---

## Autor

**Victory Lesson** — [github.com/victorylesson](https://github.com/victorylesson) | [linkedin.com/in/victory-lesson](https://linkedin.com/in/victory-lesson)
