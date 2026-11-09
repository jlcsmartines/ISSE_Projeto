import logging
import os
import sys

sys.path.append(os.path.dirname(__file__))

from fastapi import FastAPI, Request, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
import uvicorn

from etl import run_pipeline, MAX_FILE_SIZE_BYTES
from database import (
    init_db, upsert_municipio, salvar_analise,
    listar_municipios, buscar_analise, listar_analises_municipio
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(name)s — %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

init_db()

app = FastAPI(
    title="Plataforma ISSE",
    description="Análise do Índice de Segregação Socioespacial — UEMG Passos",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["GET", "POST", "DELETE"],
    allow_headers=["*"],
)


@app.get("/")
def health():
    return {"status": "ok", "versao": "1.0.0"}


@app.post("/upload")
async def upload(request: Request, file: UploadFile = File(...)):
    filename = file.filename or ""
    if not filename.lower().endswith((".xlsx", ".xls")):
        return JSONResponse(status_code=400, content={"errors": ["Envie um arquivo .xlsx ou .xls."]})

    cl = request.headers.get("content-length")
    if cl and int(cl) > MAX_FILE_SIZE_BYTES:
        return JSONResponse(status_code=413, content={"errors": ["Arquivo muito grande. Limite: 20 MB."]})

    file_bytes = await file.read()
    logger.info("Upload: %s (%d bytes)", filename, len(file_bytes))

    df, errors = run_pipeline(file_bytes)

    if errors:
        status = 400 if df.empty else 422
        return JSONResponse(status_code=status, content={"errors": errors})

    registros = df.to_dict(orient="records")
    municipio_nome = str(registros[0].get("municipio", "Desconhecido")).strip()
    mid = upsert_municipio(municipio_nome)
    salvar_analise(mid, filename, registros)

    logger.info("Salvo: %s — %d setores", municipio_nome, len(df))

    return {
        "errors": [],
        "municipio": municipio_nome,
        "row_count": len(df),
        "columns": list(df.columns),
        "data": registros,
    }


@app.get("/municipios")
def municipios():
    return {"municipios": listar_municipios()}


@app.get("/municipios/{nome}")
def municipio_detalhe(nome: str):
    data = buscar_analise(nome)
    if not data:
        return JSONResponse(status_code=404, content={"errors": [f"Município '{nome}' não encontrado."]})
    return data


@app.get("/municipios/{nome}/historico")
def historico(nome: str):
    rows = listar_analises_municipio(nome)
    if not rows:
        return JSONResponse(status_code=404, content={"errors": [f"Nenhuma análise para '{nome}'."]})
    return {"municipio": nome, "analises": rows}


@app.get("/comparar")
def comparar(municipios: str):
    nomes = [n.strip() for n in municipios.split(",") if n.strip()]
    if len(nomes) < 2:
        return JSONResponse(status_code=400, content={"errors": ["Informe ao menos 2 municípios separados por vírgula."]})

    resultado = []
    for nome in nomes:
        data = buscar_analise(nome)
        if not data:
            continue
        setores = data.get("setores", [])
        issees  = [s["isse"] for s in setores if s.get("isse") is not None]
        cats    = [s.get("categoria") for s in setores if s.get("categoria")]
        resultado.append({
            "municipio":    nome,
            "num_setores":  len(setores),
            "isse_medio":   round(sum(issees) / len(issees), 4) if issees else None,
            "isse_min":     round(min(issees), 4) if issees else None,
            "isse_max":     round(max(issees), 4) if issees else None,
            "distribuicao": {
                "Nula":  cats.count("Nula"),
                "Baixa": cats.count("Baixa"),
                "Média": cats.count("Média"),
                "Alta":  cats.count("Alta"),
            },
        })

    if not resultado:
        return JSONResponse(status_code=404, content={"errors": ["Nenhum dos municípios informados foi encontrado."]})

    return {"comparacao": resultado}


if __name__ == "__main__":
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True, log_level="info")
