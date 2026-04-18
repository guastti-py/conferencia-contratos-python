# =====================================================================
# UTILITÁRIOS DE ARQUIVOS E DADOS
# Funções de leitura, padronização, datas e conversão de colunas
# =====================================================================
import pandas as pd
from datetime import datetime


def validar_data(data_texto):
    try:
        datetime.strptime(data_texto, "%d.%m.%Y")
        return True
    except ValueError:
        return False


def limpar_valor_monetario(valor_str):
    if not valor_str:
        return 0.0

    valor_limpo = (
        str(valor_str)
        .replace("R$", "")
        .replace(".", "")
        .replace(",", ".")
        .strip()
    )

    try:
        return float(valor_limpo)
    except ValueError:
        return 0.0


def carregar_arquivo(caminho_arquivo):
    if caminho_arquivo.lower().endswith(".csv"):
        try:
            return pd.read_csv(caminho_arquivo, sep=None, engine="python", encoding="utf-8")
        except UnicodeDecodeError:
            return pd.read_csv(caminho_arquivo, sep=None, engine="python", encoding="latin-1")

    return pd.read_excel(caminho_arquivo)


def padronizar_colunas(df):
    df.columns = df.columns.str.strip()
    return df


def converter_colunas_numericas(df):
    colunas_monetarias = ["VALOR_BRUTO", "VALOR_LIQUIDO", "IOF", "CAD"]
    colunas_inteiras = ["NUM_PARC", "COD_CART", "COD_LF", "DATA_BASE"]
    colunas_decimais = ["COD_TAXA"]

    for col in colunas_monetarias:
        if col not in df.columns:
            if col != "CAD":
                df[col] = 0
            continue

        df[col] = (
            df[col]
            .astype(str)
            .str.replace("R$", "", regex=False)
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
            .str.strip()
        )
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    for col in colunas_inteiras:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(".0", "", regex=False)
                .str.strip()
            )
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    for col in colunas_decimais:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(".", "", regex=False)
                .str.replace(",", ".", regex=False)
                .str.strip()
            )
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    if "CAD" not in df.columns:
        df["CAD"] = 0

    return df


def converter_colunas_data(df):
    colunas_data = ["DT_EMISSAO", "VENC_INI"]

    for col in colunas_data:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)

    return df
