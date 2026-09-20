"""
TOOLKIT PANDAS - Análise exploratória feita para ler no terminal
================================================================

RODAR (relatório completo no terminal):
    python toolkit_pandas.py                       # usa o CONFIG abaixo
    python toolkit_pandas.py outra_base.csv        # outro arquivo
    python toolkit_pandas.py -v Valor -c Tipo -d Data
    python toolkit_pandas.py --ajuda               # todas as opções

EXPLORAR DEPOIS (o DataFrame `df` fica na memória):
    python -i toolkit_pandas.py
    >>> df[df["Valor"] > 1000]
    >>> agrupar(df, "Tipo", "Valor")

USAR NO SEU PROJETO:
    from toolkit_pandas import carregar, agrupar, outliers_iqr, mostrar
    df = carregar("transacoes.csv")
    mostrar(agrupar(df, "Tipo", "Valor"))

Fluxo:  carregar -> visão geral -> qualidade -> resumo -> outliers
        -> maiores/menores -> categorias -> tempo -> correlação
"""
import argparse
import codecs
import re
import sys
import textwrap
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# CONFIG - valores padrão (a linha de comando sobrescreve qualquer um deles)
# ---------------------------------------------------------------------------
BASE = Path(__file__).resolve().parent   # pasta onde este script está

ARQUIVO = BASE / "transacoes.csv"        # .csv, .xlsx, .xls, .json, .parquet
COL_VALOR = "Valor"        # coluna numérica principal (None = escolhe sozinho)
COL_CATEGORIA = None       # ex: "Tipo" (None = só sugere candidatas)
COL_DATA = None            # ex: "Data" (None = só sugere candidatas)
TOP = 5                    # quantas linhas nas listas de maiores/menores

# Vazio = autodetecção. Só preencha se a detecção errar, ex:
# LEITURA_KWARGS = {"sep": ";", "encoding": "latin-1"}
LEITURA_KWARGS = {}

LARGURA = 78               # largura das linhas do relatório


# ===========================================================================
# APRESENTAÇÃO - tudo que deixa a saída legível
# ===========================================================================
def br(x, casas=2):
    """Número no formato brasileiro: 1234567.891 -> '1.234.567,89'."""
    if x is None or pd.isna(x):
        return "-"
    if isinstance(x, (int, np.integer)):
        casas = 0
    s = f"{x:,.{casas}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _cortar(texto, n):
    texto = str(texto)
    return texto if len(texto) <= n else texto[: n - 1] + "…"


def _celula(v, casas=2):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return "-"
    if isinstance(v, pd.Timestamp):
        if v.hour == 0 and v.minute == 0 and v.second == 0:
            return v.strftime("%d/%m/%Y")
        return v.strftime("%d/%m/%Y %H:%M")
    if isinstance(v, (bool, np.bool_)):
        return str(v)
    if isinstance(v, (int, np.integer)):
        return str(v)
    if isinstance(v, (float, np.floating)):
        return br(v, casas)
    return _cortar(v, 30)


def titulo(texto):
    print("\n" + "=" * LARGURA)
    print(f"  {texto}")
    print("=" * LARGURA)


def subtitulo(texto):
    print(f"\n  ── {texto} " + "─" * max(0, LARGURA - len(texto) - 7))


def aviso(texto):
    print(f"  ! {texto}")


def nota(texto):
    print(f"  {texto}")


def tabela(cabecalho, linhas, alinh=None, recuo=2):
    """Imprime tabela alinhada. alinh: lista de 'l' (esquerda) ou 'r' (direita)."""
    alinh = alinh or ["l"] * len(cabecalho)
    larg = [max([len(str(c))] + [len(str(l[i])) for l in linhas])
            for i, c in enumerate(cabecalho)]

    def fmt(celulas):
        partes = []
        for i, c in enumerate(celulas):
            c = str(c)
            partes.append(c.rjust(larg[i]) if alinh[i] == "r" else c.ljust(larg[i]))
        return " " * recuo + "  ".join(partes).rstrip()

    print(fmt(cabecalho))
    print(" " * recuo + "  ".join("─" * w for w in larg))
    for linha in linhas:
        print(fmt(linha))


def mostrar(df, casas=2, indice=True, recuo=2):
    """Imprime um DataFrame alinhado, com números no formato brasileiro."""
    if isinstance(df, pd.Series):
        df = df.to_frame()
    cab = ([df.index.name or ""] if indice else []) + [str(c) for c in df.columns]
    alinh = ["l"] if indice else []
    for c in df.columns:
        num = pd.api.types.is_numeric_dtype(df[c]) and not pd.api.types.is_bool_dtype(df[c])
        alinh.append("r" if num else "l")
    linhas = []
    for idx, linha in zip(df.index, df.itertuples(index=False)):
        celulas = [_celula(v, casas) for v in linha]
        linhas.append(([_celula(idx, casas)] if indice else []) + celulas)
    tabela(cab, linhas, alinh, recuo)


def barra(valor, maximo, largura=24):
    if maximo is None or maximo == 0 or pd.isna(valor):
        return ""
    return "█" * max(1 if valor else 0, round(largura * abs(valor) / maximo))


def par(rotulo, valor, larg=22):
    print(f"    {rotulo:<{larg}}{valor}")


# ===========================================================================
# UTILITÁRIOS DE COLUNAS
# ===========================================================================
def _normalizar(texto):
    """'  Valor Total ' -> 'valortotal' (sem acento, espaço ou maiúscula)."""
    texto = unicodedata.normalize("NFKD", str(texto))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]", "", texto.lower())


def _eh_texto(serie):
    return pd.api.types.is_object_dtype(serie) or pd.api.types.is_string_dtype(serie)


def resolver_coluna(df, nome):
    """Acha a coluna mesmo com diferença de maiúscula, acento ou espaço.
    Se não achar, levanta erro listando as colunas disponíveis."""
    if nome in df.columns:
        return nome
    alvo = _normalizar(nome)
    for col in df.columns:
        if _normalizar(col) == alvo:
            return col
    raise KeyError(f"Coluna '{nome}' não existe.\nColunas disponíveis: {list(df.columns)}")


_PISTAS_VALOR = ("valor", "total", "preco", "amount", "value", "montante",
                 "quantia", "faturamento", "receita", "saldo", "price")


def parece_id(df, col):
    """Coluna numérica que é só identificador (nome 'id' ou inteiros únicos)."""
    nome = _normalizar(col)
    s = df[col].dropna()
    unico_inteiro = len(s) > 0 and s.is_unique and (s % 1 == 0).all()
    return nome == "id" or nome.startswith("id") or nome.endswith("id") or unico_inteiro


def escolher_coluna_valor(df):
    """Coluna numérica mais provável de ser 'o valor': primeiro por nome
    (valor, total, preço...), senão a primeira numérica que não seja ID."""
    numericas = list(df.select_dtypes("number").columns)
    for col in numericas:
        if any(p in _normalizar(col) for p in _PISTAS_VALOR):
            return col
    candidatas = [c for c in numericas if not parece_id(df, c)]
    return candidatas[0] if candidatas else None


def converter_data(serie):
    """Converte para datetime aceitando ISO (2026-01-31) e BR (31/01/2026)."""
    try:
        return pd.to_datetime(serie, format="ISO8601")
    except (ValueError, TypeError):
        return pd.to_datetime(serie, dayfirst=True)


def colunas_data(df):
    """Colunas que parecem datas (já datetime ou texto no formato de data)."""
    out = []
    for col in df.columns:
        s = df[col]
        if pd.api.types.is_datetime64_any_dtype(s):
            out.append(col)
        elif _eh_texto(s):
            amostra = s.dropna().astype(str).head(50)
            if amostra.empty:
                continue
            if amostra.str.match(r"^\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}").mean() < 0.9:
                continue
            try:
                converter_data(amostra)
                out.append(col)
            except Exception:
                pass
    return out


def colunas_categoricas(df, max_unicos=30):
    """Colunas de texto/booleanas com poucos valores distintos (bons grupos)."""
    datas = set(colunas_data(df))
    out = []
    for col in df.columns:
        s = df[col]
        if col in datas:
            continue
        if pd.api.types.is_bool_dtype(s) or _eh_texto(s):
            n = s.nunique(dropna=True)
            if 2 <= n <= max_unicos and n < len(s) * 0.5:
                out.append(col)
    return out


# ===========================================================================
# 1. CARGA
# ===========================================================================
def _detectar_csv(caminho):
    """Descobre encoding e separador olhando o começo do arquivo."""
    with open(caminho, "rb") as f:
        amostra = f.read(1_000_000)
    encoding = "utf-8-sig"
    try:
        # incremental: tolera um caractere cortado no fim da amostra
        texto = codecs.getincrementaldecoder("utf-8-sig")().decode(amostra, final=False)
    except UnicodeDecodeError:
        encoding = "latin-1"
        texto = amostra.decode("latin-1")
    linhas = [l for l in texto.splitlines() if l.strip()]
    cabecalho = linhas[0] if linhas else ""
    contagem = {s: cabecalho.count(s) for s in [";", ",", "\t", "|"]}
    sep = max(contagem, key=contagem.get)
    return (sep if contagem[sep] else ","), encoding


_PADRAO_BR = re.compile(
    r"^-?\s*(R\$)?\s*-?\d{1,3}(\.\d{3})*(,\d+)?$|^-?\s*(R\$)?\s*-?\d+,\d+$"
)


def _converter_numericas_br(df):
    """Texto tipo 'R$ 1.234,56' vira número (1234.56). Só mexe em coluna onde
    >=90% dos valores têm cara de número brasileiro e há vírgula ou 'R$' -
    IDs como '00123' ficam intactos."""
    df = df.copy()
    for col in df.columns:
        serie = df[col]
        if not _eh_texto(serie) or pd.api.types.is_numeric_dtype(serie):
            continue
        texto = serie.dropna().astype(str).str.strip()
        if texto.empty:
            continue
        bate = texto.map(lambda v: bool(_PADRAO_BR.match(v)))
        if bate.mean() >= 0.90 and texto.str.contains(",|R\\$", regex=True).any():
            limpo = (
                serie.astype("string").str.strip()
                .str.replace("R$", "", regex=False)
                .str.replace(r"\s", "", regex=True)
                .str.replace(".", "", regex=False)
                .str.replace(",", ".", regex=False)
            )
            df[col] = pd.to_numeric(limpo, errors="coerce")
    return df


def carregar(caminho, verbose=True, **kwargs):
    """Lê o arquivo escolhendo o leitor pela extensão.
    CSV: detecta separador/encoding (sobrescreva via kwargs) e converte
    números no formato brasileiro."""
    caminho = Path(caminho)
    leitores = {
        ".xlsx": pd.read_excel, ".xls": pd.read_excel, ".csv": pd.read_csv,
        ".json": pd.read_json, ".parquet": pd.read_parquet,
    }
    if not caminho.exists():
        raise FileNotFoundError(
            f"Arquivo não encontrado: {caminho}\nPasta de execução atual: {Path.cwd()}"
        )
    if caminho.is_dir():
        raise ValueError(f"{caminho} é uma pasta. Aponte para o arquivo.")
    ext = caminho.suffix.lower()
    if ext not in leitores:
        raise ValueError(
            f"Extensão não suportada: '{ext}' em {caminho.name}\n"
            f"Suportadas: {', '.join(leitores)}"
        )
    if ext == ".csv":
        sep, enc = _detectar_csv(caminho)
        kwargs.setdefault("sep", sep)
        kwargs.setdefault("encoding", enc)
        if verbose:
            nota(f"Leitura do CSV: separador={kwargs['sep']!r}  encoding={kwargs['encoding']!r}")
    return _converter_numericas_br(leitores[ext](caminho, **kwargs))


# ===========================================================================
# 2. VISÃO GERAL E QUALIDADE
# ===========================================================================
def _nome_tipo(serie):
    if pd.api.types.is_bool_dtype(serie):
        return "booleano"
    if pd.api.types.is_integer_dtype(serie):
        return "inteiro"
    if pd.api.types.is_float_dtype(serie):
        return "decimal"
    if pd.api.types.is_datetime64_any_dtype(serie):
        return "data"
    return "texto"


def dicionario_dados(df):
    """Uma linha por coluna: tipo, nulos, valores distintos e um exemplo."""
    linhas = []
    for col in df.columns:
        s = df[col]
        nulos = int(s.isnull().sum())
        ex = s.dropna().iloc[0] if nulos < len(s) else None
        linhas.append({
            "coluna": col,
            "tipo": _nome_tipo(s),
            "nulos": nulos,
            "% nulos": round(nulos / len(s) * 100, 1) if len(s) else 0.0,
            "distintos": int(s.nunique(dropna=True)),
            "exemplo": _celula(ex),
        })
    return pd.DataFrame(linhas)


def visao_geral(df, linhas_amostra=5):
    """Formato, dicionário de dados e amostra."""
    nota(f"{br(len(df))} linhas  x  {len(df.columns)} colunas")
    subtitulo("Colunas")
    d = dicionario_dados(df)
    tabela(
        ["coluna", "tipo", "nulos", "% nulos", "distintos", "exemplo"],
        [[r["coluna"], r["tipo"], br(r["nulos"]), br(r["% nulos"], 1),
          br(r["distintos"]), r["exemplo"]] for _, r in d.iterrows()],
        ["l", "l", "r", "r", "r", "l"],
    )
    subtitulo(f"Primeiras {linhas_amostra} linhas")
    mostrar(df.head(linhas_amostra), indice=False)


def diagnostico_nulos(df):
    """Tabela de nulos por coluna (quantidade e %). Só mostra quem tem nulo.

    Lembrete: antes de apagar/preencher, pergunte POR QUE está vazio.
    O vazio pode ser informação (ex: pedido ainda não entregue).
    """
    qtd = df.isnull().sum()
    out = pd.DataFrame({"nulos": qtd, "pct": (qtd / len(df) * 100).round(2)})
    return out[out["nulos"] > 0].sort_values("pct", ascending=False)


def duplicatas(df, subset=None):
    """Mostra TODAS as linhas duplicadas (inclusive a primeira ocorrência).
    subset = lista de colunas que definem 'duplicado' (None = linha inteira).
    Para remover: df.drop_duplicates(subset=subset)
    """
    out = df[df.duplicated(subset=subset, keep=False)]
    return out.sort_values(subset) if subset else out


def tratar_nulos(df, metodo="mediana", colunas=None, min_nao_nulos=None):
    """Devolve uma CÓPIA tratada (não altera o original).

    metodo:
      'remover_linhas'   -> dropna(): remove linhas com qualquer nulo
      'remover_colunas'  -> dropna(axis=1, thresh=min_nao_nulos):
                            mantém só colunas com pelo menos N valores preenchidos
      'zero' | 'media' | 'mediana' | 'moda' -> preenche as colunas indicadas
    """
    df = df.copy()
    if metodo == "remover_linhas":
        return df.dropna()
    if metodo == "remover_colunas":
        return df.dropna(axis=1, thresh=min_nao_nulos)
    if colunas is None:
        colunas = list(df.select_dtypes("number").columns)
    for col in colunas:
        if metodo == "zero":
            df[col] = df[col].fillna(0)
        elif metodo == "media":
            df[col] = df[col].fillna(df[col].mean())
        elif metodo == "mediana":
            df[col] = df[col].fillna(df[col].median())
        elif metodo == "moda":
            moda = df[col].mode()
            if not moda.empty:
                df[col] = df[col].fillna(moda.iloc[0])
        else:
            raise ValueError(f"Método desconhecido: {metodo}")
    return df


# ===========================================================================
# 3. ESTATÍSTICA
# ===========================================================================
def resumo_numerico(serie):
    """Todos os indicadores de uma coluna numérica em uma tacada."""
    moda = serie.mode()
    media, desvio = serie.mean(), serie.std()
    return pd.Series({
        "contagem": serie.count(),
        "media": media,
        "mediana": serie.median(),
        "moda": moda.iloc[0] if not moda.empty else np.nan,
        "desvio_padrao": desvio,
        "variancia": serie.var(),
        "coef_variacao_pct": desvio / abs(media) * 100 if media else np.nan,
        "min": serie.min(),
        "q25": serie.quantile(0.25),
        "q75": serie.quantile(0.75),
        "max": serie.max(),
        "amplitude": serie.max() - serie.min(),
    })


def limites_iqr(serie, k=1.5):
    """(limite inferior, limite superior) = Q1 - k*IQR, Q3 + k*IQR."""
    q1, q3 = serie.quantile([0.25, 0.75])
    iqr = q3 - q1
    return q1 - k * iqr, q3 + k * iqr


def outliers_iqr(df, coluna, k=1.5):
    """Linhas com valor fora dos limites do IQR. Candidatos a erro ou fraude."""
    inf, sup = limites_iqr(df[coluna], k)
    return df[(df[coluna] < inf) | (df[coluna] > sup)]


def top_n(df, coluna, n=3):
    """As N maiores linhas por coluna (nlargest). Use nsmallest para as menores."""
    return df.nlargest(n, coluna)


def correlacao(df, colunas=None):
    """Matriz de correlação entre colunas numéricas (covariância: .cov())."""
    base = df[colunas] if colunas else df.select_dtypes("number")
    return base.corr()


def leitura_em_frase(serie, n_outliers=0):
    """Traduz os números em texto - a frase é o produto, não a tabela."""
    r = resumo_numerico(serie)
    media, med, cv = r["media"], r["mediana"], r["coef_variacao_pct"]
    frases = [f"O valor típico (mediana) é {br(med)}."]
    if med:
        dif = (media - med) / abs(med)
        if dif > 0.10:
            frases.append(f"A média ({br(media)}) fica acima da mediana: poucos valores "
                          "altos puxam ela para cima.")
        elif dif < -0.10:
            frases.append(f"A média ({br(media)}) fica abaixo da mediana: poucos valores "
                          "baixos puxam ela para baixo.")
        else:
            frases.append("Média e mediana são próximas: distribuição equilibrada.")
    if not pd.isna(cv):
        nivel = "baixa" if cv < 15 else "moderada" if cv < 30 else "alta"
        frases.append(f"A variação é {nivel}: o desvio padrão ({br(r['desvio_padrao'])}) "
                      f"equivale a {br(cv, 1)}% da média.")
    frases.append(f"Metade dos valores fica entre {br(r['q25'])} e {br(r['q75'])}.")
    if n_outliers:
        frases.append(f"{br(n_outliers)} valores fogem do padrão (outliers) e merecem conferência.")
    return " ".join(frases)


# ===========================================================================
# 4. AGRUPAMENTO E CRUZAMENTO
# ===========================================================================
def agrupar(df, chave, valor, funcs=("sum", "mean", "count")):
    """groupby com várias métricas, ordenado pela primeira.
    chave pode ser uma coluna ou lista: agrupar(df, ['Produto','Vendedor'], 'Total')
    """
    funcs = list(funcs)
    return df.groupby(chave)[valor].agg(funcs).sort_values(funcs[0], ascending=False)


def cruzar(df, linhas, colunas, valor, func="sum"):
    """Tabela dinâmica (estilo Excel) com totais nas margens."""
    return df.pivot_table(index=linhas, columns=colunas, values=valor,
                          aggfunc=func, margins=True, fill_value=0)


def serie_temporal(df, col_data, valor, freq="MS", func="sum"):
    """Agrega por período. freq: 'D' dia, 'W' semana, 'MS' mês, 'QS' trimestre."""
    tmp = df.copy()
    tmp[col_data] = converter_data(tmp[col_data])
    tmp = tmp.dropna(subset=[col_data])
    return tmp.set_index(col_data)[valor].resample(freq).agg(func)


def frequencia(df, coluna, limite=8):
    """Contagem de cada valor de uma coluna, com % e barra."""
    vc = df[coluna].value_counts(dropna=False)
    total, maximo = vc.sum(), vc.max()
    linhas = []
    for k, v in vc.head(limite).items():
        nome = "(vazio)" if pd.isna(k) else _cortar(k, 26)
        linhas.append([nome, br(v), f"{v / total * 100:.1f}%".replace(".", ","),
                       barra(v, maximo)])
    tabela([coluna, "qtd", "%", ""], linhas, ["l", "r", "r", "l"], recuo=4)
    if len(vc) > limite:
        nota(f"    (+{len(vc) - limite} outros valores)")


def _tabela_categoria(df, cat, val, limite=15):
    g = agrupar(df, cat, val, ("sum", "mean", "count"))
    total, maximo = g["sum"].sum(), g["sum"].abs().max()
    linhas = []
    for nome, r in g.head(limite).iterrows():
        pct = r["sum"] / total * 100 if total else np.nan
        linhas.append([_cortar(nome, 24), br(r["sum"]),
                       "-" if pd.isna(pct) else f"{pct:.1f}%".replace(".", ","),
                       br(r["mean"]), br(int(r["count"])), barra(r["sum"], maximo)])
    tabela([cat, "total", "%", "média", "qtd", ""], linhas,
           ["l", "r", "r", "r", "r", "l"], recuo=4)
    if len(g) > limite:
        nota(f"    (+{len(g) - limite} categorias menores não exibidas)")


def _tabela_tempo(df, col_data, val, limite=24):
    datas = converter_data(df[col_data]).dropna()
    dias = (datas.max() - datas.min()).days
    freq, fmt, nome = (("MS", "%Y-%m", "mês") if dias > 60 else
                       ("W", "%d/%m/%Y", "semana") if dias > 14 else
                       ("D", "%d/%m/%Y", "dia"))
    s = serie_temporal(df, col_data, val, freq)
    delta = s.pct_change() * 100
    maximo = s.abs().max()
    linhas = []
    for ts, v in s.tail(limite).items():
        d = delta.loc[ts]
        linhas.append([ts.strftime(fmt), br(v),
                       "-" if pd.isna(d) or np.isinf(d) else f"{d:+.1f}%".replace(".", ","),
                       barra(v, maximo)])
    tabela([nome, "total", "var.", ""], linhas, ["l", "r", "r", "l"], recuo=4)
    if len(s) > limite:
        nota(f"    (mostrando os últimos {limite} de {len(s)} períodos)")


# ===========================================================================
# RELATÓRIO COMPLETO
# ===========================================================================
def analisar(df, col_valor=None, col_categoria=None, col_data=None, top=5):
    """Relatório completo no terminal. Devolve o próprio df."""
    # --- resolve colunas contra a base real ---
    if col_valor is None:
        col_valor = escolher_coluna_valor(df)
        if col_valor is None:
            raise ValueError("Não achei coluna numérica que sirva como valor.\n"
                             f"Use -v NOME. Colunas: {list(df.columns)}")
        aviso(f"Coluna de valor não definida; usando '{col_valor}'.")
    else:
        col_valor = resolver_coluna(df, col_valor)
        if not pd.api.types.is_numeric_dtype(df[col_valor]):
            raise ValueError(
                f"A coluna '{col_valor}' não é numérica (tipo: {df[col_valor].dtype}).\n"
                f"Exemplos: {df[col_valor].dropna().head(5).tolist()}")
    col_categoria = resolver_coluna(df, col_categoria) if col_categoria else None
    col_data = resolver_coluna(df, col_data) if col_data else None
    serie = df[col_valor]
    contador = iter(range(1, 20))
    prox = lambda: next(contador)

    # --- 1. visão geral ---
    titulo(f"{prox()}. VISÃO GERAL")
    visao_geral(df)
    cand_cat = [c for c in colunas_categoricas(df) if c != col_valor]
    cand_data = colunas_data(df)
    if (not col_categoria and cand_cat) or (not col_data and cand_data):
        subtitulo("Sugestões para aprofundar")
        if not col_categoria and cand_cat:
            nota(f"Candidatas a categoria: {', '.join(cand_cat)}   (use -c NOME)")
        if not col_data and cand_data:
            nota(f"Candidatas a data:      {', '.join(cand_data)}   (use -d NOME)")

    # --- 2. qualidade ---
    titulo(f"{prox()}. QUALIDADE DOS DADOS")
    nulos = diagnostico_nulos(df)
    subtitulo("Nulos")
    if nulos.empty:
        nota("Nenhum valor nulo.")
    else:
        tabela(["coluna", "nulos", "%"],
               [[c, br(int(r["nulos"])), br(r["pct"], 1)] for c, r in nulos.iterrows()],
               ["l", "r", "r"], recuo=4)
        nota("Antes de apagar ou preencher, descubra POR QUE está vazio.")
    subtitulo("Duplicatas")
    dup = int(df.duplicated().sum())
    if dup == 0:
        nota("Nenhuma linha duplicada.")
    else:
        aviso(f"{br(dup)} linhas duplicadas (linha inteira idêntica).")
        mostrar(duplicatas(df).head(top), indice=False, recuo=4)

    # --- 3. ficha da variável ---
    titulo(f"{prox()}. FICHA DA VARIÁVEL: {col_valor}")
    r = resumo_numerico(serie)
    repete = serie.duplicated().any()
    subtitulo("Centro")
    par("média", br(r["media"]))
    par("mediana", br(r["mediana"]))
    par("moda", br(r["moda"]) if repete else "sem repetição")
    subtitulo("Espalhamento")
    par("amplitude", br(r["amplitude"]))
    par("desvio padrão", br(r["desvio_padrao"]))
    par("coef. de variação", f"{br(r['coef_variacao_pct'], 1)}%")
    subtitulo("Posição")
    par("mínimo", br(r["min"]))
    par("Q1 (25%)", br(r["q25"]))
    par("Q3 (75%)", br(r["q75"]))
    par("máximo", br(r["max"]))
    outs = outliers_iqr(df, col_valor)
    subtitulo("Leitura em uma frase")
    for linha in textwrap.wrap(leitura_em_frase(serie, len(outs)), LARGURA - 6):
        nota("  " + linha)

    # --- 4. outliers ---
    titulo(f"{prox()}. OUTLIERS (regra do IQR)")
    inf, sup = limites_iqr(serie)
    nota(f"Faixa normal: {br(inf)}  a  {br(sup)}")
    if outs.empty:
        nota("Nenhum outlier encontrado.")
    else:
        pct = len(outs) / len(df) * 100
        aviso(f"{br(len(outs))} valores fora da faixa ({br(pct, 1)}% da base).")
        mostrar(outs.reindex(outs[col_valor].abs().sort_values(ascending=False).index).head(10),
                indice=False, recuo=4)
        if len(outs) > 10:
            nota(f"    (+{len(outs) - 10} outliers não exibidos)")

    # --- 5. maiores / menores ---
    titulo(f"{prox()}. MAIORES E MENORES ({col_valor})")
    subtitulo(f"{top} maiores")
    mostrar(df.nlargest(top, col_valor), indice=False, recuo=4)
    subtitulo(f"{top} menores")
    mostrar(df.nsmallest(top, col_valor), indice=False, recuo=4)

    # --- 6. categorias ---
    if col_categoria:
        titulo(f"{prox()}. {col_valor.upper()} POR {col_categoria.upper()}")
        _tabela_categoria(df, col_categoria, col_valor)
    elif cand_cat:
        titulo(f"{prox()}. DISTRIBUIÇÃO DAS CATEGORIAS")
        nota("Nenhuma categoria escolhida; contagem das candidatas:")
        for c in cand_cat[:4]:
            print()
            frequencia(df, c)

    # --- 7. tempo ---
    if col_data:
        titulo(f"{prox()}. {col_valor.upper()} AO LONGO DO TEMPO ({col_data})")
        _tabela_tempo(df, col_data, col_valor)

    # --- 8. correlação ---
    numericas = [c for c in df.select_dtypes("number").columns if not parece_id(df, c)]
    if len(numericas) >= 2:
        titulo(f"{prox()}. CORRELAÇÃO ENTRE COLUNAS NUMÉRICAS")
        nota("Perto de +1 ou -1 = relação forte; perto de 0 = pouca relação.")
        print()
        mostrar(correlacao(df, numericas[:12]), casas=2, recuo=2)

    print("\n" + "=" * LARGURA)
    print("  Fim. Para explorar: python -i toolkit_pandas.py  (o df fica na memória)")
    print("=" * LARGURA)
    return df


# ===========================================================================
# COLA RÁPIDA - filtros que você sempre esquece
# ===========================================================================
# df.loc['B', 'X']                     -> por NOME (linha, coluna)
# df.iloc[1, 1]                        -> por POSIÇÃO
# df[df['Total'] > 100]                -> filtro simples
# df[(df['A'] > 0) & (df['B'] > 0)]    -> E  (parênteses obrigatórios)
# df[(df['A'] > 0) | (df['B'] > 0)]    -> OU
# df[~(df['A'] > 0)]                   -> NÃO
# df[df['Nome'].isin(['João','Maria'])]-> está na lista
# df[df['Total'].between(100, 500)]    -> intervalo
# df[df['Nome'].str.contains('jo', case=False, na=False)] -> texto
# df.query("Total > 100 and Loja == 'SP'")                -> sintaxe legível
# df['Nova'] = df['Total'] * 0.05      -> coluna calculada
# df.drop('Nova', axis=1)              -> axis=1 coluna / axis=0 linha


# ===========================================================================
# LINHA DE COMANDO
# ===========================================================================
def main(argv=None):
    p = argparse.ArgumentParser(
        description="Análise exploratória no terminal.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        add_help=False,
        epilog=textwrap.dedent("""\
            exemplos:
              python toolkit_pandas.py
              python toolkit_pandas.py vendas.xlsx -v Total -c Loja -d Data
              python toolkit_pandas.py base.csv --sep ";" --encoding latin-1
              python -i toolkit_pandas.py      (explorar o df depois do relatório)
        """),
    )
    p.add_argument("arquivo", nargs="?", default=None, help="arquivo (padrão: CONFIG)")
    p.add_argument("-v", "--valor", help="coluna numérica principal")
    p.add_argument("-c", "--categoria", help="coluna para agrupar")
    p.add_argument("-d", "--data", help="coluna de data")
    p.add_argument("-n", "--top", type=int, default=TOP, help="linhas em maiores/menores")
    p.add_argument("--sep", help="separador do CSV (padrão: autodetecta)")
    p.add_argument("--encoding", help="encoding do CSV (padrão: autodetecta)")
    p.add_argument("--ajuda", action="help", help="mostra esta ajuda")
    args = p.parse_args(argv)

    arquivo = Path(args.arquivo) if args.arquivo else Path(ARQUIVO)
    kwargs = dict(LEITURA_KWARGS)
    if args.sep:
        kwargs["sep"] = args.sep
    if args.encoding:
        kwargs["encoding"] = args.encoding

    try:
        titulo(f"ARQUIVO: {arquivo.name}")
        df = carregar(arquivo, **kwargs)
        return analisar(
            df,
            col_valor=args.valor or COL_VALOR,
            col_categoria=args.categoria or COL_CATEGORIA,
            col_data=args.data or COL_DATA,
            top=args.top,
        )
    except (FileNotFoundError, KeyError, ValueError) as erro:
        print(f"\nERRO: {erro.args[0] if erro.args else erro}", file=sys.stderr)
        return None


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):          # acentos e barras no Windows
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    df = main()
    if df is None and not sys.flags.interactive:
        sys.exit(1)