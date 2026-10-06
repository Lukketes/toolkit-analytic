# 🐼 Toolkit Pandas

Análise exploratória **feita para ler no terminal**. Um comando lê CSV, Excel, JSON ou Parquet, arruma os números brasileiros (`R$ 1.234,56`) e as datas (`31/01/2026`) e gera um relatório em seções. Tem qualidade dos dados, estatística, categorias, tempo, correlação e antifraude, e termina numa lista de **achados em português**: a frase, não a tabela.

Tudo num arquivo só (`toolkit_pandas.py`): dá para copiar para qualquer pasta, rodar no terminal, explorar com `python -i` ou importar num notebook.

```bash
pip install -r requirements.txt
python toolkit_pandas.py --demo          # veja tudo funcionando com dados fictícios
```

---

## Rodar

```bash
python toolkit_pandas.py vendas.xlsx                                # tudo automático
python toolkit_pandas.py vendas.xlsx -v Total -c Loja -d Data       # colunas escolhidas
python toolkit_pandas.py extrato.csv -f -e Cliente --limite 10000   # + antifraude
python toolkit_pandas.py base.csv -q "Status == 'Aprovada'"         # filtra antes
python toolkit_pandas.py base.csv -s visao,qualidade                # só olhar a base nova
python toolkit_pandas.py base.csv -f --exportar achados.xlsx        # manda pro gestor
python toolkit_pandas.py -h
```

| Opção | Para quê |
|---|---|
| `arquivo` | `.csv` `.txt` `.tsv` `.xlsx` `.xlsm` `.xls` `.json` `.jsonl` `.parquet` |
| `-v`, `--valor` | coluna numérica principal (padrão: escolhe pelo nome: valor, total, vlr, preço…) |
| `-c`, `--categoria` | coluna para agrupar (sem ela, resume as candidatas) |
| `-d`, `--data` | coluna de data (padrão: a única que houver) |
| `-e`, `--entidade` | quem transaciona: cliente, fornecedor, conta, CPF… (ABC e fracionamento) |
| `-f`, `--fraude` | inclui a seção antifraude |
| `--limite` | limite de alçada para o teste de fracionamento (`10000`, `10.000`, `10k`) |
| `-q`, `--filtro` | filtra antes de analisar, na sintaxe do `df.query` (nome com espaço entre crases) |
| `-s`, `--secoes` | só algumas seções: `visao,qualidade,ficha,outliers,extremos,categorias,entidades,tempo,correlacao,fraude,achados` |
| `-n`, `--top` | linhas nas listas de maiores/menores (padrão 5) |
| `--aba`, `--cabecalho` | aba do Excel (nome ou número) / linha do cabeçalho; os dois são detectados sozinhos |
| `--sep`, `--encoding` | só se a autodetecção do CSV errar |
| `--exportar ARQ.xlsx` | achados e todas as tabelas em abas do Excel (filtro, cabeçalho fixo, formato BR) |
| `--salvar ARQ.txt` | o relatório em texto, sem as cores |
| `--sem-cor`, `--largura N` | saída sem cor / largura fixa (padrão: a do terminal) |
| `--demo` | dados fictícios com problemas plantados |
| `--debug` | mostra o traceback quando algo dá errado |

Não sabe quais colunas usar? Rode sem nada. A tabela **Colunas usadas na análise** mostra o que foi escolhido e lista as candidatas.

Guarde as bases em `dados/` e os relatórios em `saidas/`: as duas pastas estão no `.gitignore`, então dado do estágio não vai parar no GitHub por engano.

Usa sempre a mesma base? Preencha o `CONFIG` no topo do arquivo (`ARQUIVO`, `COL_VALOR`, `FRAUDE = True`…). A linha de comando sobrescreve o `CONFIG`, e uma coluna do `CONFIG` que não exista na base é ignorada com aviso, sem dar erro.

---

## O que o relatório mostra

| Seção | Conteúdo |
|---|---|
| Cabeçalho | formato, separador, encoding, aba, linha do cabeçalho e **o que foi convertido** (e o que falhou) |
| Visão geral | dicionário de dados com o **papel** de cada coluna (identificador, categórica, data, pessoal…), amostra e colunas escolhidas |
| Qualidade | vazios, linhas duplicadas e checagens automáticas de regras de negócio (lista abaixo) |
| Ficha da variável | as 4 perguntas da planilha: que dado é, valor típico, quanto varia, onde cada valor está. Inclui histograma e a **leitura em uma frase** |
| Outliers | faixa normal pelo IQR, quantos ficam acima/abaixo, quantos são extremos e os maiores desvios |
| Maiores e menores | as N linhas extremas |
| Categorias | total, %, % acumulado, **classe ABC**, média e barra. Avisa se a mesma categoria está escrita de jeitos diferentes |
| Entidades (`-e`) | concentração (o maior responde por X%, os 10% maiores por Y%) e curva ABC |
| Tempo | total por dia/semana/mês/trimestre com variação, **períodos incompletos marcados**, tendência robusta, meses sem registro, dia da semana e faixa de horário |
| Correlação | matriz (até 6 colunas) e os pares mais relacionados, escritos por extenso |
| Antifraude (`-f`) | veja abaixo |
| Principais achados | tudo o que importa em uma lista, alertas primeiro, pronta para colar num e-mail |

Nas tabelas de registros, a coluna **`linha` é a linha no arquivo original**. Dá para abrir o Excel e ir direto nela.

### Checagens automáticas de qualidade

Colunas vazias ou constantes · ID repetido numa coluna quase única · datas no futuro (exceto vencimento/validade) ou antes de 1900 · **grafias diferentes da mesma categoria** (`PIX`/`pix`/`Pix `, `SP`/`S.P.`/`São Paulo`) · espaços sobrando · número guardado como texto · **CPF/CNPJ com dígito verificador errado** (inclui o CNPJ alfanumérico de 2026) · valores negativos/zerados · erros de fórmula do Excel (`#N/D`, `#VALOR!`) · dados pessoais (LGPD).

### Antifraude

| Teste | O que procura | Cuidados |
|---|---|---|
| Lei de Benford (1º e 2 primeiros dígitos) | dígitos iniciais fora da proporção natural; conformidade pelo **MAD de Nigrini** + qui² | precisa de ≥ 300 valores espalhados em ≥ 2 ordens de grandeza (o relatório avisa quando não se aplica) |
| Z-score robusto (MAD) | valores atípicos sem ser enganado pelos próprios outliers | em valores de dinheiro, calcula na escala log (senão a cauda inteira vira "atípica") |
| Valores repetidos e redondos | o mesmo valor exato muitas vezes; excesso de múltiplos de 1.000 | preços tabelados são legítimos |
| Possível duplicidade | mesmo valor + mesma entidade + mesmo dia (sem entidade: mesmo minuto) | mostra quanto teria sido pago a mais |
| Fracionamento | **degraus logo abaixo de limites redondos** (detectados sozinhos) e, com `-e`, a mesma entidade quebrando um valor grande em vários pequenos no mesmo dia | no varejo, preço psicológico (9,90) também gera degrau |
| Calendário | fim de semana, **feriados nacionais** (inclui Carnaval, Sexta-feira Santa e Corpus Christi, calculados pela Páscoa), madrugada | lê a hora da própria data ou de uma coluna "Hora" |
| Registros para conferir | pontuação de risco por linha somando os alertas, do mais suspeito para o menos | lista completa em `rel.alertas` ou na aba **Alertas** do `--exportar` |

Sinais, não provas: cada alerta é um convite para conferir.

---

## Explorar depois do relatório

```bash
python -i toolkit_pandas.py base.csv -f
```

O relatório roda e o terminal continua aberto com `df` (a base já arrumada) e `rel` (o resultado):

```python
>>> ajuda()                                   # todas as funções, por grupo
>>> ajuda(curva_abc)                          # detalhes de uma
>>> cola()                                    # filtros e comandos de pandas que você sempre esquece
>>> mostrar(rel.alertas, max_linhas=50)       # registros sinalizados
>>> rel.tabelas.keys()                        # todas as tabelas do relatório
>>> df["UF"] = padronizar_uf(df["UF"])
>>> mostrar(agrupar(df, "UF", "Valor"))
>>> exportar_excel(rel, "achados.xlsx")
```

`mostrar()` imprime qualquer DataFrame no formato BR e **cabe no terminal**: colunas que não cabem são listadas em vez de quebrar a linha.

---

## Usar no seu projeto ou notebook

```python
from toolkit_pandas import carregar, checar_merge, curva_abc, mostrar

vendas = carregar("vendas.xlsx")                   # já com números e datas convertidos
clientes = carregar("clientes.csv")
checar_merge(vendas, clientes, "CPF")              # SEMPRE antes do merge
mostrar(curva_abc(vendas, "Cliente", "Valor").head(20))
```

Veio de SQL, de API ou da área de transferência? `preparar(df)` aplica a mesma limpeza. Todas as funções devolvem DataFrames (para imprimir bonito, use `mostrar`) e nenhuma altera o original.

| Grupo | Funções |
|---|---|
| Carga | `carregar` · `preparar` · `converter_numero` · `converter_data` |
| Ver | `mostrar` · `tabela` · `visao_geral` · `dicionario_dados` · `br` · `pct` · `barra` |
| Qualidade | `checar_qualidade` · `diagnostico_nulos` · `duplicatas` · `tratar_nulos` · `grafias_inconsistentes` · `padronizar_texto` · `padronizar_uf` |
| LGPD | `detectar_pii` · `mascarar_pii` · `validar_cpf` · `validar_cnpj` |
| Estatística | `resumo_numerico` · `leitura_em_frase` · `histograma` · `limites_iqr` · `outliers_iqr` · `zscore_robusto` · `outliers_mad` · `correlacao` · `pares_correlacao` |
| Agrupar | `agrupar` · `cruzar` · `curva_abc` · `frequencia` · `serie_temporal` · `resumo_temporal` · `coorte` |
| Antifraude | `benford` · `valores_repetidos` · `valores_redondos` · `possiveis_duplicatas` · `acumulo_abaixo_de_limites` · `fracionamento` · `marcar_calendario` · `feriados_br` · `sinalizar` |
| Juntar tabelas | `checar_merge` · `padronizar_chave` |
| Relatório | `analisar` · `exportar_excel` · `dados_exemplo` · `ajuda` · `cola` |
| Colunas | `resolver_coluna` · `escolher_coluna_valor` · `colunas_data` · `colunas_categoricas` · `colunas_entidade` · `parece_id` |

---

## Robustez embutida

- **CSV brasileiro de verdade:** detecta separador (`; , TAB |`) mesmo com vírgula dentro de aspas, encoding (utf-8, cp1252, latin-1) e **linhas de título antes do cabeçalho**. Linhas mal formadas são puladas, com aviso.
- **Excel de empresa:** acha o cabeçalho mesmo com título e data em cima, avisa quando há outras abas e remove linhas e colunas vazias de formatação.
- **Números:** `R$ 1.234,56`, `(500,00)`, `1.234,56-`, `12,5%` e `1,234.56` (americano) viram número. **Nada é corrompido em silêncio:** um `1.5` perdido numa coluna BR fica vazio e aparece no aviso, em vez de virar 15. Códigos com zero à esquerda (`00123`), CPF, CEP, telefone e cartão continuam texto.
- **Datas:** `31/01/2026`, `2026-01-31`, `31.01.2026`, com ou sem hora. Datas ambíguas (`01/02`) são lidas como dia/mês. As que não convertem são contadas e mostradas.
- **Colunas:** `"valor "`, `"VALOR"` e `"Vàlor"` funcionam. Errou o nome? Ele sugere o certo (`Você quis dizer: 'Valor'?`).
- **Erros** aparecem em uma linha (`ERRO: ...`): arquivo aberto no Excel, biblioteca faltando, aba inexistente, filtro inválido… Se uma seção falhar, as outras continuam.
- **Desempenho:** 1 milhão de linhas com relatório completo + antifraude em ~30 s. As conversões rodam sobre os valores distintos, não linha a linha.
- **Testado com pandas 2.3 e 3.0**, sem nenhum aviso. Cores no terminal (desligam sozinhas ao salvar em arquivo; respeita `NO_COLOR`).

---

## Princípios

- **Nulo não é lixo, é pista.** Antes de apagar ou preencher, pergunte por que está vazio.
- **`fillna(0)` não é neutro:** o pandas ignora `NaN` na média, mas o `0` entra como valor real. Para números, a mediana costuma ser mais segura (é o padrão de `tratar_nulos`).
- **Média mentirosa:** com poucos valores muito altos, a média engana. O relatório avisa e fala do típico pela mediana.
- **Mês incompleto não se compara.** Períodos parciais ficam marcados com `*` e saem da variação e da tendência.
- **Confira o merge antes de fazer.** Chave repetida dos dois lados multiplica linhas e infla somas (`checar_merge`).
- **Trabalhe com cópias.** Nenhuma função altera o DataFrame original.
- **O produto é a frase, não a tabela.** Traduza os números em conclusão.

---

## Testes

```bash
pip install pytest
python -m pytest -q
```

66 testes cobrem formatação, leitura (CSV BR/US, Excel com título, JSON aninhado), conversões, detecção de colunas, estatística, antifraude, merge, relatório e linha de comando.

---

## Roadmap

- [x] Módulo antifraude: Benford, z-score robusto (MAD), fracionamento, horário/feriados, duplicidade
- [x] Validação de regras de negócio (datas futuras, IDs repetidos, grafias diferentes, CPF/CNPJ)
- [x] Padronização de texto e UF, `checar_merge()` antes de cruzar tabelas
- [x] Detecção e mascaramento de dados pessoais (CPF, e-mail, telefone, nome)
- [x] Curva ABC/Pareto e coorte
- [x] Testes com `pytest`
- [ ] Comparar duas versões da mesma base (o que entrou, saiu e mudou)
- [ ] Feriados estaduais e municipais
- [ ] Relatório em HTML com gráficos

### O que mudou da versão 1

Bugs corrigidos: `Idade` era tratada como ID e sumia da correlação; `"1.5"` numa coluna BR virava `15` sem aviso; o último mês incompleto aparecia como queda de −40%; bases largas geravam linhas de 1.700 caracteres; arquivos sem coluna `Valor` davam erro; erro de encoding aparecia como `ERRO: utf-8`; arquivo aberto no Excel gerava traceback.

Novidades: antifraude, checagens de qualidade, papel de cada coluna, curva ABC, entidades, dia da semana e horário, achados finais, exportação para Excel, `--filtro`, `--secoes`, `--demo`, cores, `ajuda()`/`cola()`, LGPD, `checar_merge` e suporte a Excel com título, JSON aninhado e CSV com título.
