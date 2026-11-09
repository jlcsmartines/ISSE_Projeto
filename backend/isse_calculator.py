import pandas as pd


SE_COLS = [
    "distancia_centro_km",
    "tempo_ape_min",
    "tempo_carro_min",
    "tempo_onibus_min",
]
SS_COLS = [
    "familias_bolsa_familia",
    "familias_extrema_pobreza",
]
DA_COLS = [
    "num_usf",
    "num_upa",
    "num_hospital_publico",
    "num_hospital_privado",
    "num_escolas_municipais",
    "num_escolas_estaduais",
    "num_escolas_privadas",
    "num_cemei",
    "num_creches",
    "dias_coleta_lixo",
    "linhas_onibus",
]

NOMINAL_COLS = {"linhas_onibus"}


def _cluster_score(series: pd.Series, *, higher_is_worse: bool = False) -> pd.Series:
    mean = series.mean()
    std = series.std(ddof=1)

    if pd.isna(mean) or pd.isna(std) or std == 0:
        return pd.Series(0.5, index=series.index)

    hi = mean + std
    lo = mean - std

    def _score(v):
        if pd.isna(v):
            return 0.5
        if higher_is_worse:
            if v < lo:
                return 0.00
            if v < mean:
                return 0.25
            if v < hi:
                return 0.75
            return 1.00
        else:
            if v > hi:
                return 0.00
            if v > mean:
                return 0.25
            if v > lo:
                return 0.75
            return 1.00

    return series.apply(_score)


def calculate_isse(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    present = lambda cols: [c for c in cols if c in df.columns]
    se_cols = present(SE_COLS)
    ss_cols = present(SS_COLS)
    da_cols = present(DA_COLS)

    scores = pd.DataFrame(index=df.index)

    for col in se_cols:
        scores[col] = _cluster_score(df[col], higher_is_worse=True)
    for col in ss_cols:
        scores[col] = _cluster_score(df[col], higher_is_worse=True)
    for col in da_cols:
        if col in NOMINAL_COLS:
            scores[col] = df[col].apply(lambda v: 0.0 if v == 1 else 1.0)
        else:
            scores[col] = _cluster_score(df[col], higher_is_worse=False)

    def _cat_score(cols):
        present_scores = [scores[c] for c in cols if c in scores.columns]
        if not present_scores:
            return pd.Series(0.0, index=df.index)
        return pd.concat(present_scores, axis=1).sum(axis=1) / len(present_scores)

    se_score = _cat_score(se_cols)
    ss_score = _cat_score(ss_cols)
    da_score = _cat_score(da_cols)

    df["isse"] = ((se_score + ss_score + da_score) / 3).round(4)

    # Categorização do ISSE conforme metodologia de Bastos Filho et al. (2019, p. 60)
    # A categorização em 4 grupos baseia-se na média e no desvio-padrão do próprio índice na amostra:
    # 1. Não Segregada: ISSE < (Média - 1 Desvio-Padrão)
    # 2. Baixa Segregação: (Média - 1 Desvio-Padrão) <= ISSE < Média
    # 3. Média Segregação: Média <= ISSE <= (Média + 1 Desvio-Padrão)
    # 4. Alta Segregação: ISSE > (Média + 1 Desvio-Padrão)
    mean_isse = df["isse"].mean()
    std_isse = df["isse"].std(ddof=1)

    if pd.isna(std_isse) or std_isse == 0:
        df["categoria_isse"] = "Média"
    else:
        lo = mean_isse - std_isse
        hi = mean_isse + std_isse

        def _classificar(v):
            if v < lo:
                return "Não Segregada"
            elif v < mean_isse:
                return "Baixa"
            elif v <= hi:
                return "Média"
            else:
                return "Alta"

        df["categoria_isse"] = df["isse"].apply(_classificar)

    return df
