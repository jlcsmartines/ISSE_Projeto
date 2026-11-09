import json
import os
import sqlite3
from typing import List, Optional

DB_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "isse.db")


def _conn():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    con = _conn()
    con.executescript("""
        CREATE TABLE IF NOT EXISTS municipios (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            nome      TEXT    NOT NULL,
            uf        TEXT,
            populacao INTEGER,
            criado_em TEXT    DEFAULT (datetime('now'))
        );

        CREATE TABLE IF NOT EXISTS analises (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            municipio_id INTEGER NOT NULL REFERENCES municipios(id),
            arquivo      TEXT,
            criado_em    TEXT DEFAULT (datetime('now'))
        );

        CREATE TABLE IF NOT EXISTS setores (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            analise_id  INTEGER NOT NULL REFERENCES analises(id),
            nome        TEXT    NOT NULL,
            isse        REAL,
            categoria   TEXT,
            indicadores TEXT
        );

        CREATE INDEX IF NOT EXISTS idx_setores_analise ON setores(analise_id);
        CREATE INDEX IF NOT EXISTS idx_analises_municipio ON analises(municipio_id);
    """)
    con.commit()
    con.close()


def upsert_municipio(nome: str, uf: str = None) -> int:
    con = _conn()
    row = con.execute("SELECT id FROM municipios WHERE nome = ?", (nome,)).fetchone()
    if row:
        mid = row["id"]
    else:
        cur = con.execute("INSERT INTO municipios (nome, uf) VALUES (?, ?)", (nome, uf))
        mid = cur.lastrowid
    con.commit()
    con.close()
    return mid


def salvar_analise(municipio_id: int, arquivo: str, setores: list) -> int:
    con = _conn()
    cur = con.execute(
        "INSERT INTO analises (municipio_id, arquivo) VALUES (?, ?)",
        (municipio_id, arquivo),
    )
    aid = cur.lastrowid
    for s in setores:
        indicadores = {k: v for k, v in s.items()
                       if k not in ("setor", "municipio", "isse", "categoria_isse")}
        con.execute(
            "INSERT INTO setores (analise_id, nome, isse, categoria, indicadores) "
            "VALUES (?, ?, ?, ?, ?)",
            (aid, s.get("setor", ""), s.get("isse"), s.get("categoria_isse"), json.dumps(indicadores)),
        )
    con.commit()
    con.close()
    return aid


def listar_municipios():
    con = _conn()
    rows = con.execute("""
        SELECT m.id, m.nome, m.uf,
               COUNT(DISTINCT a.id) as num_analises,
               MAX(a.criado_em)     as ultima_analise
        FROM municipios m
        LEFT JOIN analises a ON a.municipio_id = m.id
        GROUP BY m.id
        ORDER BY m.nome
    """).fetchall()
    con.close()
    return [dict(r) for r in rows]


def buscar_analise(municipio_nome: str):
    con = _conn()
    row = con.execute("""
        SELECT a.id, a.criado_em, m.nome as municipio, m.uf
        FROM analises a
        JOIN municipios m ON m.id = a.municipio_id
        WHERE m.nome = ?
        ORDER BY a.criado_em DESC
        LIMIT 1
    """, (municipio_nome,)).fetchone()
    if not row:
        con.close()
        return None
    aid = row["id"]
    setores = con.execute(
        "SELECT nome, isse, categoria, indicadores FROM setores WHERE analise_id = ?",
        (aid,)
    ).fetchall()
    con.close()
    result = dict(row)
    result["setores"] = [
        {**{"nome": s["nome"], "isse": s["isse"], "categoria": s["categoria"]},
         **json.loads(s["indicadores"] or "{}")}
        for s in setores
    ]
    return result


def listar_analises_municipio(municipio_nome: str):
    con = _conn()
    rows = con.execute("""
        SELECT a.id, a.arquivo, a.criado_em,
               COUNT(s.id) as num_setores,
               ROUND(AVG(s.isse), 4) as isse_medio
        FROM analises a
        JOIN municipios m ON m.id = a.municipio_id
        JOIN setores s ON s.analise_id = a.id
        WHERE m.nome = ?
        GROUP BY a.id
        ORDER BY a.criado_em DESC
    """, (municipio_nome,)).fetchall()
    con.close()
    return [dict(r) for r in rows]
