# 📊 auto-relatorio

> Automação de leitura de planilhas e geração de relatórios com Python

Lê arquivos `.xlsx` ou `.csv`, analisa automaticamente as colunas e gera relatórios em **HTML estilizado** e **JSON estruturado** — sem configuração manual.

---

## ✨ Funcionalidades

- 📥 Suporte a `.xlsx` e `.csv` via linha de comando
- 🔢 **Colunas numéricas**: soma, média, mediana, mínimo, máximo e desvio padrão
- 🔤 **Colunas de texto**: top 5 valores mais frequentes
- ⚠️ **Alertas de qualidade**: detecta e reporta valores nulos com percentual
- 📄 Exporta relatório em **HTML** (estilizado, pronto para apresentar) ou **JSON** (integrável com outras ferramentas)

---

## 🚀 Como usar

### Instalação

```bash
git clone https://github.com/victorylesson/auto-relatorio.git
cd auto-relatorio
pip install -r requirements.txt
```

### Execução

```bash
# Uso básico — gera HTML na pasta atual
python relatorio.py sua_planilha.xlsx

# Especificando formato e pasta de saída
python relatorio.py vendas.csv -f html -o relatorios/janeiro

# Gerar apenas JSON
python relatorio.py dados.xlsx -f json -o resultados/
```

### Gerar planilha de exemplo

```bash
python gera_exemplo.py
# Cria demo.xlsx com 120 registros fictícios para testar
```

---

## 📁 Estrutura do projeto

```
auto-relatorio/
├── relatorio.py        # Script principal
├── gera_exemplo.py     # Gerador de planilha demo (120 registros)
├── requirements.txt    # Dependências
├── README.md
└── .gitignore
```

---

## 📦 Dependências

| Pacote | Uso |
|---|---|
| `pandas` | Leitura e análise dos dados |
| `openpyxl` | Suporte a arquivos `.xlsx` |
| `jinja2` | Template do relatório HTML |

Instale tudo com:

```bash
pip install -r requirements.txt
```

---

## 📊 Exemplo de saída

Dado um arquivo `vendas.csv` com colunas como `valor`, `produto`, `regiao`:

```
=== RELATÓRIO: vendas.csv ===

[NUMÉRICO] valor
  Soma:            R$ 48.320,00
  Média:           R$ 402,67
  Mediana:         R$ 385,00
  Mín / Máx:       R$ 50,00 / R$ 990,00
  Desvio padrão:   R$ 218,45

[TEXTO] produto — top 5 mais frequentes
  1. Notebook         18x
  2. Monitor          15x
  3. Teclado          12x
  ...

⚠️  regiao: 4 valores nulos (3.3%)
```

---

## 🛠️ Argumentos disponíveis

| Argumento | Descrição | Padrão |
|---|---|---|
| `arquivo` | Caminho do `.xlsx` ou `.csv` | obrigatório |
| `-f`, `--formato` | Formato de saída: `html` ou `json` | `html` |
| `-o`, `--output` | Pasta de destino do relatório | pasta atual |

---

## 💡 Casos de uso

- Análise rápida de exportações de ERP/CRM
- Validação de qualidade de dados antes de importar
- Geração de relatórios periódicos automatizados via `cron` ou agendador
- Integração em pipelines de dados simples


## 🌐 Demo

Veja um exemplo do relatório gerado em:  
🔗 [https://auto-relatorio-three.vercel.app/](https://auto-relatorio-three.vercel.app/)

---

## 🤝 Contribuindo

Pull requests são bem-vindos. Para mudanças maiores, abra uma *issue* primeiro para discutir o que você gostaria de alterar.

---

## 📄 Licença

[MIT](LICENSE)
