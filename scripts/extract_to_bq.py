import io
import os
import re
import time
import unicodedata
from datetime import datetime, timezone

import pandas as pd
import requests
from google.cloud import bigquery
from openpyxl import load_workbook

# O site da SGF foi refeito em Setembro de 2026. O Excel deixou de ter um URL
# fixo: passou a ser publicado com a data no nome (Historico-de-Cotacoes_MMDD.xlsx)
# dentro da pasta do mes do upload. Por isso o URL e descoberto em cada execucao
# a partir da pagina publica, em vez de estar escrito no codigo.
PAGE_URL = "https://goldensgf.pt/informacao-dos-fundos/"
MEDIA_API = "https://goldensgf.pt/wp-json/wp/v2/media"
FALLBACK_URL = "https://goldensgf.pt/wp-content/uploads/2026/09/Historico-de-Cotacoes_0916.xlsx"

FUND_NAME = "SGF DR FINANÇAS"
MIN_ROWS = 100

PROJECT_ID = os.environ["GCP_PROJECT_ID"]
DATASET = "PPR_SGF_DF"
TABLE = "sgf_dr_financas_nav"
TABLE_REF = f"{PROJECT_ID}.{DATASET}.{TABLE}"

HEADERS = {"User-Agent": "Mozilla/5.0"}

SCHEMA = [
    bigquery.SchemaField("data", "DATE"),
    bigquery.SchemaField("nav", "FLOAT64"),
    bigquery.SchemaField("fundo", "STRING"),
    bigquery.SchemaField("data_extracao", "TIMESTAMP"),
]


def norm(value):
    """Maiusculas, sem acentos e sem espacos a mais, para comparar nomes."""
    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(c for c in text if not unicodedata.combining(c))
    return " ".join(text.split()).upper()


def resolve_excel_url():
    """Descobre o URL actual do Excel. Tenta a pagina, depois a API do WordPress."""
    candidates = []

    try:
        print(f"A procurar o link do Excel em: {PAGE_URL}")
        html = requests.get(PAGE_URL, headers=HEADERS, timeout=60).text
        candidates += re.findall(
            r"https://goldensgf\.pt/wp-content/uploads/[^\"'\s>]+\.xlsx", html
        )
    except Exception as e:
        print(f"Nao foi possivel ler a pagina: {e}")

    if not candidates:
        try:
            print("A tentar a API de media do WordPress...")
            resp = requests.get(
                MEDIA_API,
                params={"search": "Historico", "per_page": 50,
                        "orderby": "date", "order": "desc"},
                headers=HEADERS,
                timeout=60,
            )
            resp.raise_for_status()
            candidates += [
                item.get("source_url", "")
                for item in resp.json()
                if str(item.get("source_url", "")).lower().endswith(".xlsx")
            ]
        except Exception as e:
            print(f"API de media falhou: {e}")

    # Ficar so com os ficheiros de cotacoes, mantendo a ordem de descoberta.
    cotacoes = [u for u in candidates if "COTAC" in norm(u)]
    ordered = cotacoes or candidates
    seen, urls = set(), []
    for u in ordered:
        if u not in seen:
            seen.add(u)
            urls.append(u)
    urls.append(FALLBACK_URL)

    print(f"Candidatos encontrados: {urls}")
    return urls


def download_excel(urls, retries=3):
    last_error = None
    for url in urls:
        for attempt in range(retries):
            try:
                print(f"A descarregar: {url} (tentativa {attempt + 1}/{retries})")
                response = requests.get(url, headers=HEADERS, timeout=120)
                response.raise_for_status()
                if not response.content.startswith(b"PK"):
                    raise ValueError("A resposta nao e um ficheiro xlsx.")
                print(f"Descarregado: {len(response.content)} bytes")
                return url, response.content
            except Exception as e:
                last_error = e
                print(f"Erro: {e}")
                if attempt < retries - 1:
                    time.sleep(10)
    raise RuntimeError(f"Nenhum URL funcionou. Ultimo erro: {last_error}")


def read_sheet(content):
    """Le a folha em modo streaming.

    O ficheiro novo declara ~1M de linhas, quase todas vazias. Em modo read_only
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
    print(f"Cabecalho: {header}")

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
    print(f"Linhas com dados: {len(df)}")
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
    print(f"Colunas: fundo={col_fundo}, nav={col_nav}, data={col_data}")
    return col_fundo, col_nav, col_data


def transform(df):
    df.columns = [str(c).strip() for c in df.columns]
    col_fundo, col_nav, col_data = pick_columns(df)

    alvo = norm(FUND_NAME)
    mask = df[col_fundo].map(norm) == alvo
    df_filtered = df[mask].copy()
    print(f"Registos apos filtro '{FUND_NAME}': {len(df_filtered)}")

    if df_filtered.empty:
        disponiveis = sorted(set(df[col_fundo].dropna().astype(str)))
        print(f"Fundos disponiveis no ficheiro: {disponiveis}")
        raise ValueError(f"Fundo '{FUND_NAME}' nao encontrado.")

    serie_data = df_filtered[col_data]
    if pd.api.types.is_datetime64_any_dtype(serie_data):
        datas = serie_data
    else:
        datas = pd.to_datetime(serie_data, dayfirst=True, errors="coerce")

    df_out = pd.DataFrame()
    df_out["data"] = datas.dt.date
    df_out["nav"] = pd.to_numeric(df_filtered[col_nav], errors="coerce")
    df_out["fundo"] = FUND_NAME
    df_out["data_extracao"] = datetime.now(timezone.utc)

    before = len(df_out)
    df_out = df_out.dropna(subset=["nav", "data"])
    print(f"Removidas {before - len(df_out)} linhas nulas. Total: {len(df_out)}")

    before = len(df_out)
    df_out = df_out.drop_duplicates(subset=["data"], keep="last")
    if before != len(df_out):
        print(f"Removidas {before - len(df_out)} datas duplicadas.")

    df_out = df_out.sort_values("data", ascending=False).reset_index(drop=True)
    print(f"Intervalo: {df_out['data'].min()} a {df_out['data'].max()}")
    return df_out


def ensure_dataset_and_table(client):
    dataset_ref = bigquery.Dataset(f"{PROJECT_ID}.{DATASET}")
    dataset_ref.location = "EU"
    try:
        client.get_dataset(dataset_ref)
        print(f"Dataset '{DATASET}' ja existe.")
    except Exception:
        client.create_dataset(dataset_ref)
        print(f"Dataset '{DATASET}' criado.")

    try:
        client.get_table(TABLE_REF)
        print(f"Tabela '{TABLE}' ja existe.")
    except Exception:
        client.create_table(bigquery.Table(TABLE_REF, schema=SCHEMA))
        print(f"Tabela '{TABLE}' criada.")


def sanity_check(client, df):
    """A carga apaga e reescreve a tabela. Nao o fazer com um ficheiro suspeito."""
    if len(df) < MIN_ROWS:
        raise ValueError(f"So {len(df)} registos (minimo {MIN_ROWS}). Carga abortada.")
    try:
        existentes = client.get_table(TABLE_REF).num_rows
    except Exception:
        return
    if existentes and len(df) < existentes * 0.9:
        raise ValueError(
            f"O ficheiro traz {len(df)} registos mas a tabela ja tem {existentes}. "
            "Carga abortada para nao perder historico."
        )


def load_to_bq(client, df):
    job_config = bigquery.LoadJobConfig(
        schema=SCHEMA,
        write_disposition=bigquery.WriteDisposition.WRITE_TRUNCATE,
    )
    print(f"A carregar {len(df)} registos para {TABLE_REF}...")
    client.load_table_from_dataframe(df, TABLE_REF, job_config=job_config).result()
    table = client.get_table(TABLE_REF)
    print(f"Carga concluida. Tabela tem {table.num_rows} linhas.")


def main():
    client = bigquery.Client(project=PROJECT_ID)
    url, content = download_excel(resolve_excel_url())
    print(f"Fonte utilizada: {url}")
    df_clean = transform(read_sheet(content))
    ensure_dataset_and_table(client)
    sanity_check(client, df_clean)
    load_to_bq(client, df_clean)


if __name__ == "__main__":
    main()
