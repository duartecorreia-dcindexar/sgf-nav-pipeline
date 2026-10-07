import io
import os
import re
import time
import unicodedata
from datetime import datetime, timedelta, timezone
from urllib.parse import urljoin
from zoneinfo import ZoneInfo

import pandas as pd
import requests
from google.cloud import bigquery
from openpyxl import load_workbook

# O site da SGF foi refeito em Setembro de 2026. O Excel nao tem um URL fixo,
# por isso o URL e descoberto em cada execucao (pagina publica + API de media
# do WordPress). Como pode haver mais do que um ficheiro de cotacoes publicado
# (e um deles desactualizado), descarregam-se TODOS os candidatos e usa-se o
# que tiver a data mais recente.
PAGE_URL = "https://goldensgf.pt/informacao-dos-fundos/"
MEDIA_API = "https://goldensgf.pt/wp-json/wp/v2/media"
FALLBACK_URL = "https://goldensgf.pt/wp-content/uploads/2026/09/Historico-de-Cotacoes_0916.xlsx"

# Grafia exacta como aparece na coluna "Nome do Fundo" do Excel. A comparacao
# ignora acentos e maiusculas, mas e este o valor que fica gravado no BigQuery.
FUNDS = [
    "SGF DR FINANÇAS",
    "Golden SGF Poupança Dinamica",
    "Golden SGF ETF Start",
    "Golden SGF ETF Plus",
    "PPR SGF Stoik",
    "Golden SGF TOP GESTORES",
]

# A tabela original continua a receber so o DR Financas, para nao partir nada
# que ja dependa dela. Os seis fundos vao para a tabela nova.
LEGACY_FUND = "SGF DR FINANÇAS"
TABLE_LEGACY = "sgf_dr_financas_nav"
TABLE_ALL = "sgf_navs"

MIN_ROWS_POR_FUNDO = 100

# Validacoes antes de reescrever as tabelas.
MAX_ATRASO_DIAS = 4        # ultima cotacao comum aos 6 fundos nao pode ter mais de 4 dias
JANELA_REPETIDOS_DIAS = 60 # dias recentes verificados contra todo o historico

TZ_LISBOA = ZoneInfo("Europe/Lisbon")

PROJECT_ID = os.environ["GCP_PROJECT_ID"]
DATASET = "PPR_SGF_DF"

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Cache-Control": "no-cache, no-store, max-age=0",
    "Pragma": "no-cache",
}

SCHEMA = [
    bigquery.SchemaField("data", "DATE"),
    bigquery.SchemaField("nav", "FLOAT64"),
    bigquery.SchemaField("fundo", "STRING"),
    bigquery.SchemaField("data_extracao", "TIMESTAMP"),
]


def table_ref(table):
    return f"{PROJECT_ID}.{DATASET}.{table}"


def norm(value):
    """Maiusculas, sem acentos e sem espacos a mais, para comparar nomes."""
    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(c for c in text if not unicodedata.combining(c))
    return " ".join(text.split()).upper()


def sem_cache(url):
    """Acrescenta um parametro que muda a cada execucao, para furar caches/CDN."""
    sep = "&" if "?" in url else "?"
    return f"{url}{sep}nocache={int(time.time())}"


def links_xlsx_da_pagina(html):
    """Apanha links .xlsx absolutos, relativos e escapados em JSON/JS."""
    html = html.replace("\\/", "/")
    links = []
    # Atributos href/src/data-* (absolutos ou relativos)
    links += re.findall(
        r"""(?:href|src|data-[\w-]+)\s*=\s*["']([^"']+?\.xlsx(?:\?[^"']*)?)["']""",
        html,
        flags=re.IGNORECASE,
    )
    # Qualquer URL .xlsx solto no HTML (scripts, JSON embebido, etc.)
    links += re.findall(
        r"""(?:https?:)?//[^"'\s<>()]+?\.xlsx""", html, flags=re.IGNORECASE
    )
    links += re.findall(
        r"""(?<![\w/.:-])/wp-content/uploads/[^"'\s<>()]+?\.xlsx""",
        html,
        flags=re.IGNORECASE,
    )
    return [urljoin(PAGE_URL, l.strip()) for l in links]


def resolve_excel_urls():
    """Junta todos os candidatos: pagina publica + API de media + fallback."""
    candidates = []

    try:
        print(f"A procurar links do Excel em: {PAGE_URL}")
        resp = requests.get(sem_cache(PAGE_URL), headers=HEADERS, timeout=60)
        resp.raise_for_status()
        encontrados = links_xlsx_da_pagina(resp.text)
        print(f"  Na pagina: {encontrados}")
        candidates += encontrados
    except Exception as e:
        print(f"Nao foi possivel ler a pagina: {e}")

    try:
        print("A consultar a API de media do WordPress...")
        resp = requests.get(
            MEDIA_API,
            params={"search": "Cotac", "per_page": 50,
                    "orderby": "modified", "order": "desc"},
            headers=HEADERS,
            timeout=60,
        )
        resp.raise_for_status()
        encontrados = [
            item.get("source_url", "")
            for item in resp.json()
            if str(item.get("source_url", "")).lower().endswith(".xlsx")
        ]
        print(f"  Na API de media: {encontrados}")
        candidates += encontrados
    except Exception as e:
        print(f"API de media falhou: {e}")

    candidates.append(FALLBACK_URL)

    # Ficar so com os ficheiros de cotacoes, sem repetidos, mantendo a ordem.
    cotacoes = [u for u in candidates if "COTAC" in norm(u)]
    seen, urls = set(), []
    for u in cotacoes or candidates:
        chave = u.split("?")[0].replace("http://", "https://")
        if chave not in seen:
            seen.add(chave)
            urls.append(u)

    print(f"Candidatos: {urls}")
    return urls


def download_excel(url, retries=3):
    last_error = None
    for attempt in range(retries):
        try:
            print(f"A descarregar: {url} (tentativa {attempt + 1}/{retries})")
            response = requests.get(sem_cache(url), headers=HEADERS, timeout=120)
            response.raise_for_status()
            if not response.content.startswith(b"PK"):
                raise ValueError("A resposta nao e um ficheiro xlsx.")
            print(
                f"  Descarregado: {len(response.content)} bytes | "
                f"Last-Modified: {response.headers.get('Last-Modified')} | "
                f"ETag: {response.headers.get('ETag')} | "
                f"Cache: {response.headers.get('X-Cache') or response.headers.get('CF-Cache-Status') or response.headers.get('Age')}"
            )
            return response.content
        except Exception as e:
            last_error = e
            print(f"  Erro: {e}")
            if attempt < retries - 1:
                time.sleep(10)
    raise RuntimeError(f"Falhou o download de {url}: {last_error}")


def read_sheet(content):
    """Le a folha em modo streaming.

    O ficheiro declara ~1M de linhas, quase todas vazias. Em modo read_only
    a memoria mantem-se estavel e paramos assim que acabam os dados.
    """
    wb = load_workbook(io.BytesIO(content), read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = ws.iter_rows(values_only=True)

    def tres(row):
        """Garante exactamente 3 celulas: o read_only corta as vazias do fim."""
        cells = list(row)[:3]
        return cells + [None] * (3 - len(cells))

    header = tres(next(rows))
    print(f"  Cabecalho: {header}")

    records, empty_streak = [], 0
    for row in rows:
        cells = tres(row)
        if all(c is None for c in cells):
            empty_streak += 1
            if empty_streak >= 500:
                break
            continue
        empty_streak = 0
        records.append(cells)
    wb.close()

    df = pd.DataFrame(records, columns=[str(h) for h in header])
    print(f"  Linhas com dados: {len(df)}")
    return df


def pick_columns(df):
    """Identifica as colunas pelo nome; se falhar, usa a posicao."""
    by_name = {norm(c): c for c in df.columns}

    def find(*keys):
        for key in keys:
            for name, original in by_name.items():
                if key in name:
                    return original
        return None

    col_fundo = find("NOME DO FUNDO", "FUNDO") or df.columns[0]
    col_nav = find("COTACAO", "NAV", "VALOR") or df.columns[1]
    col_data = find("DATA") or df.columns[2]
    print(f"  Colunas: fundo={col_fundo}, nav={col_nav}, data={col_data}")
    return col_fundo, col_nav, col_data


def transform(df):
    df.columns = [str(c).strip() for c in df.columns]
    col_fundo, col_nav, col_data = pick_columns(df)

    # norm(nome do ficheiro) -> grafia canonica que fica no BigQuery
    canonico = {norm(f): f for f in FUNDS}

    fundos_norm = df[col_fundo].map(norm)
    df_filtered = df[fundos_norm.isin(set(canonico))].copy()
    print(f"  Registos dos {len(FUNDS)} fundos pedidos: {len(df_filtered)}")

    if df_filtered.empty:
        disponiveis = sorted(set(df[col_fundo].dropna().astype(str)))
        print(f"  Fundos disponiveis no ficheiro: {disponiveis}")
        raise ValueError("Nenhum dos fundos pedidos foi encontrado.")

    serie_data = df_filtered[col_data]
    if pd.api.types.is_datetime64_any_dtype(serie_data):
        datas = serie_data
    else:
        datas = pd.to_datetime(serie_data, dayfirst=True, errors="coerce")

    df_out = pd.DataFrame()
    df_out["data"] = datas.dt.date
    df_out["nav"] = pd.to_numeric(df_filtered[col_nav], errors="coerce")
    df_out["fundo"] = fundos_norm.loc[df_filtered.index].map(canonico)
    df_out["data_extracao"] = datetime.now(timezone.utc)

    before = len(df_out)
    df_out = df_out.dropna(subset=["nav", "data"])
    print(f"  Removidas {before - len(df_out)} linhas nulas. Total: {len(df_out)}")

    before = len(df_out)
    df_out = df_out.drop_duplicates(subset=["fundo", "data"], keep="last")
    if before != len(df_out):
        print(f"  Removidas {before - len(df_out)} duplicadas (mesmo fundo e data).")

    df_out = df_out.sort_values(["fundo", "data"], ascending=[True, False])
    df_out = df_out.reset_index(drop=True)

    verificar_fundos(df_out)
    return df_out


def verificar_fundos(df):
    """Um fundo que desaparece do ficheiro tem de dar erro, nao passar em silencio."""
    resumo = df.groupby("fundo").agg(
        linhas=("data", "size"), inicio=("data", "min"), fim=("data", "max")
    )
    print("\n  Por fundo:")
    print(resumo.to_string())
    print()

    em_falta = [f for f in FUNDS if f not in resumo.index]
    if em_falta:
        raise ValueError(f"Fundos nao encontrados no ficheiro: {em_falta}")

    magros = resumo[resumo["linhas"] < MIN_ROWS_POR_FUNDO]
    if not magros.empty:
        raise ValueError(
            f"Fundos com menos de {MIN_ROWS_POR_FUNDO} registos: "
            f"{list(magros.index)}. Carga abortada."
        )


def ultima_data_comum(df):
    """Ultima data em que TODOS os fundos tem cotacao."""
    return df.groupby("fundo")["data"].max().min()


def escolher_melhor_ficheiro(urls):
    """Descarrega e processa todos os candidatos; fica com o mais recente."""
    validos = []
    for url in urls:
        try:
            content = download_excel(url)
            df = transform(read_sheet(content))
            fim = ultima_data_comum(df)
            print(f"  -> {url}: ultima data comum {fim}, {len(df)} registos\n")
            validos.append((fim, len(df), url, df))
        except Exception as e:
            print(f"  -> {url} descartado: {e}\n")

    if not validos:
        raise RuntimeError("Nenhum candidato produziu dados validos.")

    validos.sort(key=lambda v: (v[0], v[1]), reverse=True)
    fim, n, url, df = validos[0]
    print(f"Fonte escolhida: {url} (ultima data {fim}, {n} registos)")
    return url, df


def verificar_atraso(df):
    """Nao reescrever a tabela com um ficheiro parado no tempo."""
    hoje = datetime.now(TZ_LISBOA).date()
    fim = ultima_data_comum(df)
    atraso = (hoje - fim).days
    print(f"Ultima cotacao comum: {fim} ({atraso} dias de atraso).")
    if atraso > MAX_ATRASO_DIAS:
        raise ValueError(
            f"O ficheiro mais recente acaba em {fim} ({atraso} dias de atraso, "
            f"maximo {MAX_ATRASO_DIAS}). Possivel ficheiro desactualizado ou em "
            f"cache. Carga abortada; as tabelas ficam como estavam."
        )


def verificar_dias_repetidos(df):
    """Detecta um dia recente com as cotacoes de outro dia em todos os fundos.

    Foi o que aconteceu a 02/10/2026: o ficheiro trazia nesse dia as cotacoes
    de 12/06/2026 nos seis fundos.
    """
    pv = df.pivot(index="data", columns="fundo", values="nav").dropna()
    if pv.empty:
        return
    limite = pv.index.max() - timedelta(days=JANELA_REPETIDOS_DIAS)
    problemas = []
    for dia, linha in pv[pv.index >= limite].iterrows():
        outros = pv.drop(index=dia)
        iguais = outros[(outros - linha).abs().max(axis=1) < 1e-9]
        for outro in iguais.index:
            problemas.append(f"{dia} = {outro}")
    if problemas:
        raise ValueError(
            "Dias com as cotacoes de outro dia em todos os fundos: "
            f"{problemas}. Carga abortada; as tabelas ficam como estavam."
        )


def ensure_dataset_and_table(client, table):
    dataset_ref = bigquery.Dataset(f"{PROJECT_ID}.{DATASET}")
    dataset_ref.location = "EU"
    try:
        client.get_dataset(dataset_ref)
    except Exception:
        client.create_dataset(dataset_ref)
        print(f"Dataset '{DATASET}' criado.")

    ref = table_ref(table)
    try:
        client.get_table(ref)
        print(f"Tabela '{table}' ja existe.")
    except Exception:
        client.create_table(bigquery.Table(ref, schema=SCHEMA))
        print(f"Tabela '{table}' criada.")


def sanity_check(client, table, df):
    """A carga apaga e reescreve a tabela. Nao o fazer com um ficheiro suspeito."""
    ref = table_ref(table)
    try:
        existentes = client.get_table(ref).num_rows
    except Exception:
        return
    if existentes and len(df) < existentes * 0.9:
        raise ValueError(
            f"{table}: o ficheiro traz {len(df)} registos mas a tabela ja tem "
            f"{existentes}. Carga abortada para nao perder historico."
        )


def load_to_bq(client, table, df):
    ref = table_ref(table)
    job_config = bigquery.LoadJobConfig(
        schema=SCHEMA,
        write_disposition=bigquery.WriteDisposition.WRITE_TRUNCATE,
    )
    print(f"A carregar {len(df)} registos para {ref}...")
    client.load_table_from_dataframe(df, ref, job_config=job_config).result()
    print(f"Carga concluida. {table} tem {client.get_table(ref).num_rows} linhas.")


def main():
    client = bigquery.Client(project=PROJECT_ID)

    url, df_todos = escolher_melhor_ficheiro(resolve_excel_urls())
    verificar_atraso(df_todos)
    verificar_dias_repetidos(df_todos)

    df_legacy = df_todos[df_todos["fundo"] == LEGACY_FUND].reset_index(drop=True)

    for table, df in ((TABLE_ALL, df_todos), (TABLE_LEGACY, df_legacy)):
        ensure_dataset_and_table(client, table)
        sanity_check(client, table, df)
        load_to_bq(client, table, df)


if __name__ == "__main__":
    main()
