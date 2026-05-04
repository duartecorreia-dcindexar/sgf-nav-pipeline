import os
import io
import time
import requests
import pandas as pd
from datetime import datetime, timezone
from google.cloud import bigquery

EXCEL_URL = "https://goldensgf.pt/wp-content/uploads/2024/08/HISTORICO-DE-COTACOES.xlsx"
FUND_NAME = "SGF DR FINANCAS"
PROJECT_ID = os.environ["GCP_PROJECT_ID"]
DATASET = "golden_sgf"
TABLE = "sgf_dr_financas_nav"
TABLE_REF = f"{PROJECT_ID}.{DATASET}.{TABLE}"


def download_excel(url, retries=3):
    for attempt in range(retries):
        try:
            print(f"A descarregar Excel de: {url} (tentativa {attempt + 1}/{retries})")
            headers = {"User-Agent": "Mozilla/5.0"}
            response = requests.get(url, headers=headers, timeout=60)
            response.raise_for_status()
            df = pd.read_excel(io.BytesIO(response.content), engine="openpyxl")
            print(f"Excel carregado: {len(df)} linhas, colunas: {list(df.columns)}")
            return df
        except Exception as e:
            print(f"Erro na tentativa {attempt + 1}: {e}")
            if attempt < retries - 1:
                print("A aguardar 10 segundos antes de repetir...")
                time.sleep(10)
            else:
                raise


def transform(df):
    df.columns = [c.strip() for c in df.columns]
    col_fundo, col_nav, col_data = df.columns[0], df.columns[1], df.columns[2]
    print(f"Colunas: fundo={col_fundo}, nav={col_nav}, data={col_data}")
    mask = df[col_fundo].str.strip().isin(["SGF DR FINANCAS", "SGF DR FINANÇAS"])
    df_filtered = df[mask].copy()
    print(f"Registos apos filtro: {len(df_filtered)}")
    if df_filtered.empty:
        print(f"Valores unicos: {df[col_fundo].unique()[:10]}")
        raise ValueError("Fundo nao encontrado.")
    df_out = pd.DataFrame()
    df_out["data"] = pd.to_datetime(df_filtered[col_data], dayfirst=True).dt.date
    df_out["nav"] = pd.to_numeric(df_filtered[col_nav], errors="coerce")
    df_out["fundo"] = df_filtered[col_fundo].str.strip()
    df_out["data_extracao"] = datetime.now(timezone.utc)
    before = len(df_out)
    df_out = df_out.dropna(subset=["nav", "data"])
    print(f"Removidas {before - len(df_out)} linhas nulas. Total: {len(df_out)}")
    df_out = df_out.sort_values("data", ascending=False).reset_index(drop=True)
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
    schema = [
        bigquery.SchemaField("data", "DATE"),
        bigquery.SchemaField("nav", "FLOAT64"),
        bigquery.SchemaField("fundo", "STRING"),
        bigquery.SchemaField("data_extracao", "TIMESTAMP"),
    ]
    table_obj = bigquery.Table(TABLE_REF, schema=schema)
    try:
        client.get_table(TABLE_REF)
        print(f"Tabela '{TABLE}' ja existe.")
    except Exception:
        client.create_table(table_obj)
        print(f"Tabela '{TABLE}' criada.")


def load_to_bq(client, df):
    job_config = bigquery.LoadJobConfig(
        schema=[
            bigquery.SchemaField("data", "DATE"),
            bigquery.SchemaField("nav", "FLOAT64"),
            bigquery.SchemaField("fundo", "STRING"),
            bigquery.SchemaField("data_extracao", "TIMESTAMP"),
        ],
        write_disposition=bigquery.WriteDisposition.WRITE_TRUNCATE,
    )
    print(f"A carregar {len(df)} registos para {TABLE_REF}...")
    job = client.load_table_from_dataframe(df, TABLE_REF, job_config=job_config)
    job.result()
    table = client.get_table(TABLE_REF)
    print(f"Carga concluida. Tabela tem {table.num_rows} linhas.")


def main():
    client = bigquery.Client(project=PROJECT_ID)
    df_raw = download_excel(EXCEL_URL)
    df_clean = transform(df_raw)
    ensure_dataset_and_table(client)
    load_to_bq(client, df_clean)


if __name__ == "__main__":
    main()
