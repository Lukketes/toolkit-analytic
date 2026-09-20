# 🐼 Toolkit Pandas

Análise exploratória de dados **feita para ler no terminal**: um comando, um relatório organizado em seções, com tabelas alinhadas, números no formato brasileiro (`1.234,56`), barras de texto e uma leitura em linguagem natural dos números. Nada de arquivos de saída.

---

## Rodar

```bash
python toolkit_pandas.py                          # usa o CONFIG do topo do arquivo
python toolkit_pandas.py vendas.xlsx              # outro arquivo
python toolkit_pandas.py -v Valor -c Tipo -d Data # colunas escolhidas na hora
python toolkit_pandas.py base.csv --sep ";" --encoding latin-1
python toolkit_pandas.py --ajuda
```

| Opção | Significado |
|---|---|
| `arquivo` | `.csv`, `.xlsx`, `.xls`, `.json` ou `.parquet` (padrão: `ARQUIVO` do CONFIG) |
| `-v`, `--valor` | coluna numérica principal (padrão: `COL_VALOR`; se `None`, ele escolhe) |
| `-c`, `--categoria` | coluna para agrupar |
| `-d`, `--data` | coluna de data (ativa a análise temporal) |
| `-n`, `--top` | quantas linhas nas listas de maiores/menores (padrão 5) |
| `--sep`, `--encoding` | só se a autodetecção do CSV errar |

Não sabe quais colunas usar? Rode sem `-c` e `-d`: a seção **Visão geral** sugere as candidatas.

### Explorar depois do relatório

```bash
python -i toolkit_pandas.py
```

O relatório roda e o terminal continua aberto com o `df` carregado e todas as funções disponíveis:

```python
>>> df[df["Valor"] > 1000]
>>> mostrar(agrupar(df, "Tipo", "Valor"))
>>> mostrar(outliers_iqr(df, "Valor"))
```

---

## O que o relatório mostra

| # | Seção | Conteúdo |
|---|---|---|
| 1 | Visão geral | formato, dicionário de dados (tipo, nulos, distintos, exemplo), amostra e sugestões de colunas |
| 2 | Qualidade | nulos por coluna e linhas duplicadas |
| 3 | Ficha da variável | centro, espalhamento, posição e a **leitura em uma frase** |
| 4 | Outliers | faixa normal (IQR), quantidade e os maiores desvios |
| 5 | Maiores e menores | as N linhas extremas |
| 6 | Categorias | total, %, média, quantidade e barra por categoria; sem `-c`, a contagem das candidatas |
| 7 | Tempo | total por mês/semana/dia com variação %, se houver `-d` |
| 8 | Correlação | matriz entre colunas numéricas (IDs ficam de fora) |

A seção 3 segue a "ficha" da planilha de estatística: quatro perguntas (que dado é esse, qual o valor típico, quanto varia, onde cada valor está), e o resultado final é uma frase, não uma tabela.

---

## Robustez embutida

- **Caminhos** ancorados na pasta do script: funciona de qualquer terminal.
- **CSV brasileiro:** detecta separador (`; , TAB |`) e encoding (utf-8 ou latin-1) sozinho.
- **Números em texto:** `R$ 1.234,56` vira número. IDs como `00123` não são tocados.
- **Datas:** aceita `2026-01-31` e `31/01/2026`.
- **Nome de coluna:** `"valor "`, `"VALOR"` e `"Vàlor"` são reconhecidos. Coluna inexistente mostra as disponíveis.
- **Erros** aparecem como uma linha `ERRO: ...`, sem traceback.

---

## Usar no seu projeto

```python
from toolkit_pandas import carregar, agrupar, outliers_iqr, mostrar

df = carregar("transacoes.csv")
mostrar(agrupar(df, "Tipo", "Valor"))
```

### Funções

| Grupo | Função | O que faz |
|---|---|---|
| Carga | `carregar(caminho, verbose=True, **kwargs)` | lê pela extensão; CSV com autodetecção |
| Apresentação | `mostrar(df, casas=2, indice=True)` | imprime DataFrame alinhado, formato BR |
| | `tabela(cabecalho, linhas, alinh)` | tabela alinhada a partir de listas |
| | `br(x, casas=2)` | formata número: `1234.5` → `1.234,50` |
| | `barra(valor, maximo)` | barra de texto proporcional |
| Inspeção | `dicionario_dados(df)` | DataFrame: tipo, nulos, distintos, exemplo por coluna |
| | `visao_geral(df)` | dicionário + amostra impressos |
| Limpeza | `diagnostico_nulos(df)` | nulos por coluna (qtd e %) |
| | `tratar_nulos(df, metodo, colunas, min_nao_nulos)` | cópia tratada: `remover_linhas`, `remover_colunas`, `zero`, `media`, `mediana`, `moda` |
| | `duplicatas(df, subset=None)` | todas as linhas duplicadas |
| Estatística | `resumo_numerico(serie)` | média, mediana, moda, desvio, variância, CV, quartis, amplitude |
| | `leitura_em_frase(serie, n_outliers)` | os números traduzidos em texto |
| | `limites_iqr(serie, k=1.5)` / `outliers_iqr(df, coluna, k)` | faixa normal e linhas fora dela |
| | `top_n(df, coluna, n)` | N maiores (`nsmallest` para as menores) |
| | `correlacao(df, colunas=None)` | matriz de correlação |
| Agrupamento | `agrupar(df, chave, valor, funcs)` | `groupby` com várias métricas |
| | `cruzar(df, linhas, colunas, valor, func)` | tabela dinâmica com totais |
| | `serie_temporal(df, col_data, valor, freq)` | agrega por período (`D`, `W`, `MS`, `QS`) |
| | `frequencia(df, coluna, limite)` | contagem de cada valor, com % e barra |
| Colunas | `resolver_coluna(df, nome)` | acha a coluna ignorando caixa/acento/espaço |
| | `escolher_coluna_valor(df)` | adivinha a coluna de valor |
| | `colunas_data(df)` / `colunas_categoricas(df)` | candidatas a data e a categoria |
| | `parece_id(df, coluna)` | identifica colunas de ID |

---

## Cola rápida de filtros

```python
df.loc['B', 'X']                                  # por NOME (linha, coluna)
df.iloc[1, 1]                                     # por POSIÇÃO
df[df['Total'] > 100]                             # filtro simples
df[(df['A'] > 0) & (df['B'] > 0)]                 # E  (parênteses obrigatórios)
df[(df['A'] > 0) | (df['B'] > 0)]                 # OU
df[~(df['A'] > 0)]                                # NÃO
df[df['Nome'].isin(['João', 'Maria'])]            # está na lista
df[df['Total'].between(100, 500)]                 # intervalo
df[df['Nome'].str.contains('jo', case=False, na=False)]  # texto
df.query("Total > 100 and Loja == 'SP'")          # sintaxe legível
```

---

## Princípios

- **Nulo não é lixo, é pista.** Antes de apagar ou preencher, pergunte por que está vazio.
- **`fillna(0)` não é neutro:** o pandas ignora `NaN` na média, mas o `0` entra como valor real. Para dados numéricos, a mediana costuma ser mais segura (é o padrão em `tratar_nulos`).
- **`thresh` é o mínimo de preenchidos:** `dropna(axis=1, thresh=N)` mantém colunas com pelo menos N valores preenchidos.
- **Trabalhe com cópias.** As funções de limpeza não alteram o original.
- **O produto é a frase, não a tabela.** Traduza os números em conclusão.

---

## Roadmap

- [ ] Módulo antifraude: Lei de Benford, z-score modificado (MAD), fracionamento de valores, transações fora do padrão de horário
- [ ] Validação de regras de negócio (datas futuras, IDs únicos, categorias inesperadas)
- [ ] Padronização de texto (`"SP"`, `"sp "`, `"São Paulo"`) e `checar_merge()` antes de cruzar tabelas
- [ ] Detecção e mascaramento de dados pessoais (CPF, e-mail, telefone)
- [ ] Análise ABC/Pareto e coorte
- [ ] Testes com `pytest`
