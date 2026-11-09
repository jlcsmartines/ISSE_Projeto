from io import BytesIO
from typing import List, Tuple
import logging
import unicodedata

import pandas as pd

if __package__:
    from .isse_calculator import DA_COLS, SE_COLS, SS_COLS, calculate_isse
else:
    from isse_calculator import DA_COLS, SE_COLS, SS_COLS, calculate_isse

logger = logging.getLogger(__name__)

MAX_FILE_SIZE_BYTES: int = 20 * 1024 * 1024
MAX_ROWS: int = 50_000

ID_COLS = ["setor", "municipio"]

REQUIRED_COLUMNS: List[str] = ID_COLS + SE_COLS + SS_COLS + DA_COLS


def _strip_accents(text: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(c)
    )


def _normalise_col(name: str) -> str:
    return _strip_accents(str(name).strip().lower())


def _read_raw_headers(xls: pd.ExcelFile, sheet_name: str) -> List[str]:
    header_df = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=1)
    if header_df.empty:
        return []
    _IGNORE = {"nan", "none", "nat", ""}
    return [
        _normalise_col(str(v))
        for v in header_df.iloc[0].tolist()
        if _normalise_col(str(v)) not in _IGNORE
    ]


def _parse_numeric(series: pd.Series) -> pd.Series:
    parsed = pd.to_numeric(series, errors="coerce")
    still_missing = parsed.isna() & series.notna()
    if still_missing.any():
        br = (
            series[still_missing]
            .astype(str)
            .str.strip()
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        parsed.loc[still_missing] = pd.to_numeric(br, errors="coerce")
    return parsed


def _anonimizar(df: pd.DataFrame) -> pd.DataFrame:
    """
    Remove ou generaliza campos que possam identificar indivíduos (LGPD).
    Nomes de setores são mantidos pois referem-se a regiões geográficas,
    não a pessoas. Colunas com dados pessoais diretos são descartadas se
    presentes acidentalmente na planilha.
    """
    PII_PATTERNS = ["cpf", "rg", "nome_morador", "email", "telefone", "endereco"]
    cols_pii = [c for c in df.columns if any(p in c for p in PII_PATTERNS)]
    if cols_pii:
        logger.warning("Colunas com possível dado pessoal removidas: %s", cols_pii)
        df = df.drop(columns=cols_pii)
    return df


def extract(file_bytes: bytes) -> pd.DataFrame:
    if len(file_bytes) > MAX_FILE_SIZE_BYTES:
        raise ValueError(
            f"Arquivo muito grande ({len(file_bytes) / 1_048_576:.1f} MB). "
            f"O limite é {MAX_FILE_SIZE_BYTES // 1_048_576} MB."
        )

    try:
        xls = pd.ExcelFile(BytesIO(file_bytes))
    except Exception as exc:
        raise ValueError(
            "Não foi possível abrir o arquivo como Excel. "
            "Verifique se não está corrompido ou com extensão trocada. "
            f"Detalhe: {exc}"
        ) from exc

    sheet = xls.sheet_names[0]
    for name in xls.sheet_names:
        if name.strip().lower() == "dados":
            sheet = name
            break

    raw_headers = _read_raw_headers(xls, sheet)
    seen: set = set()
    dupes: set = set()
    for h in raw_headers:
        (dupes if h in seen else seen).add(h)

    df = pd.read_excel(xls, sheet_name=sheet)

    if len(df) > MAX_ROWS:
        raise ValueError(
            f"A planilha tem {len(df):,} linhas, acima do limite de {MAX_ROWS:,}."
        )

    df.columns = [_normalise_col(c) for c in df.columns]
    df = df.loc[:, ~df.columns.str.startswith("unnamed")]
    df = df.dropna(axis=1, how="all")
    df.attrs["raw_duplicate_columns"] = sorted(dupes)
    return df


def find_duplicate_columns(df: pd.DataFrame) -> List[str]:
    return df.attrs.get("raw_duplicate_columns", [])


def validate_columns(df: pd.DataFrame) -> List[str]:
    errors: List[str] = []

    missing = [
        col for col in REQUIRED_COLUMNS
        if col not in df.columns or df[col].isna().all()
    ]
    if missing:
        errors.append(
            "Colunas obrigatórias ausentes ou totalmente vazias: " + ", ".join(missing)
        )

    dupes = find_duplicate_columns(df)
    if dupes:
        errors.append(
            "Colunas duplicadas (renomeie ou remova uma delas): " + ", ".join(dupes)
        )

    if df.empty:
        errors.append("A planilha está vazia.")

    return errors


def transform(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = _anonimizar(df)
    df = df.dropna(how="all")

    for col in ID_COLS:
        if col in df.columns:
            df[col] = df[col].where(
                ~df[col].apply(lambda v: isinstance(v, str) and not v.strip()),
                other=None,
            )

    df = df.dropna(subset=[c for c in ID_COLS if c in df.columns], how="any")
    df = df.reset_index(drop=True)

    numeric_cols = [c for c in df.columns if c not in ID_COLS]
    for col in numeric_cols:
        df[col] = _parse_numeric(df[col])

    cols_to_drop = []
    for col in numeric_cols:
        if pd.isna(df[col].mean()):
            logger.warning("Coluna '%s' descartada: sem valores numéricos.", col)
            cols_to_drop.append(col)
        elif df[col].isna().any():
            df[col] = df[col].fillna(df[col].mean())

    if cols_to_drop:
        df = df.drop(columns=cols_to_drop)

    return df


def run_pipeline(file_bytes: bytes) -> Tuple[pd.DataFrame, List[str]]:
    try:
        df = extract(file_bytes)
    except ValueError as exc:
        logger.error("Extração: %s", exc)
        return pd.DataFrame(), [str(exc)]

    errors = validate_columns(df)
    if errors:
        return df, errors

    try:
        df = transform(df)
    except Exception as exc:
        logger.exception("Transformação")
        return pd.DataFrame(), [f"Erro na transformação: {exc}"]

    if df.empty:
        return df, ["Nenhuma linha válida após o tratamento dos dados."]

    df = calculate_isse(df)
    return df, []
