# Plataforma ISSE — UEMG Passos

desenvolvida como projeto de extensão universitária pela UEMG Passos.

Calcula o ISSE por bairro/setor conforme a metodologia de **Bastos Filho et al. (2019)**:

## Estrutura

```
isse-platform-fixed/
├── backend/
│   ├── main.py        API FastAPI — upload, consulta, comparação
│   ├── etl.py         Pipeline ETL com anonimização LGPD
│   └── database.py    Data warehouse SQLite (municípios, análises, setores)
├── frontend/
│   └── index.html     Dashboard completo — upload, gráficos, comparação
├── data/              Banco de dados gerado automaticamente (isse.db)
├── stress_tests/      Bateria de testes do pipeline ETL
├── install.py         Instalador automático (venv + dependências)
├── install.bat        Atalho Windows
├── install.sh         Atalho Linux/Mac
├── run.bat            Inicia servidor + abre frontend (Windows)
├── run.sh             Inicia servidor + abre frontend (Linux/Mac)
├── build_template.py  Gera o modelo Excel oficial (modelo_dados_isse.xlsx)
└── generate_sample_data.py  Gera planilha de exemplo para testes
```

## Instalação

**Windows:**
```
.\install.bat
```

**Linux / Mac:**
```
chmod +x install.sh && ./install.sh
```

## Uso

```
.\run.bat        (Windows)
./run.sh         (Linux/Mac)
```

O frontend abre automaticamente no navegador. O backend fica em `http://127.0.0.1:8000`.

## Endpoints da API

| Método | Rota | Descrição |
|---|---|---|
| `POST` | `/upload` | Processa planilha Excel e salva no banco |
| `GET` | `/municipios` | Lista municípios cadastrados |
| `GET` | `/municipios/{nome}` | Detalhe da última análise |
| `GET` | `/municipios/{nome}/historico` | Histórico de análises |
| `GET` | `/comparar?municipios=A,B` | Compara dois ou mais municípios |

Documentação interativa: `http://127.0.0.1:8000/docs`

## LGPD

O pipeline remove automaticamente colunas com padrões de dados pessoais
(CPF, RG, nome de morador, e-mail, telefone, endereço) antes de qualquer
processamento ou armazenamento.

