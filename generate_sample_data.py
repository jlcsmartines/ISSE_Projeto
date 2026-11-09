#!/usr/bin/env python3
import os
import random

import pandas as pd

random.seed(42)

MUNICIPIO = "Passos"
OUT_PATH  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sample_data.xlsx")

BAIRROS = [
    "Centro", "Jardim das Flores", "Vila Nova", "Cohab I", "Cohab II",
    "Morro Alto", "Parque Industrial", "Vila Rica", "Bela Vista", "São José",
    "Recanto Verde", "Jardim América", "Boa Esperança", "Alto da Serra", "Vista Alegre",
]

rows = []
for bairro in BAIRROS:
    eh_central = bairro == "Centro"
    eh_periferia = bairro.startswith("Cohab") or bairro in ("Morro Alto", "Recanto Verde")

    dist = round(random.uniform(0.3, 1.5) if eh_central else
                 random.uniform(4.0, 12.0) if eh_periferia else
                 random.uniform(1.0, 6.0), 1)

    rows.append({
        "setor":                   bairro,
        "municipio":               MUNICIPIO,
        "distancia_centro_km":     dist,
        "tempo_ape_min":           round(dist * 12 + random.uniform(-2, 5)),
        "tempo_carro_min":         round(dist * 3  + random.uniform(1, 4)),
        "tempo_onibus_min":        round(dist * 5  + random.uniform(3, 10)),
        "familias_bolsa_familia":  random.randint(5, 20) if eh_central else
                                   random.randint(80, 250) if eh_periferia else
                                   random.randint(20, 90),
        "familias_extrema_pobreza":random.randint(1, 8) if eh_central else
                                   random.randint(30, 100) if eh_periferia else
                                   random.randint(8, 40),
        "num_usf":                 random.randint(0, 2) if eh_central else
                                   random.randint(0, 1),
        "num_upa":                 1 if eh_central else 0,
        "num_hospital_publico":    1 if eh_central else 0,
        "num_hospital_privado":    random.randint(2, 6) if eh_central else
                                   random.randint(0, 1),
        "num_escolas_municipais":  random.randint(1, 3),
        "num_escolas_estaduais":   random.randint(0, 2),
        "num_escolas_privadas":    random.randint(2, 8) if eh_central else
                                   random.randint(0, 2),
        "num_cemei":               random.randint(0, 2),
        "num_creches":             random.randint(0, 3),
        "dias_coleta_lixo":        random.randint(1, 3) if eh_periferia else
                                   random.randint(3, 6),
        "linhas_onibus":           0 if eh_periferia and random.random() < 0.4 else 1,
    })

pd.DataFrame(rows).to_excel(OUT_PATH, index=False)
print(f"Criado: {OUT_PATH} ({len(rows)} bairros)")
