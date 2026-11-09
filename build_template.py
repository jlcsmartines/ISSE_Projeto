#!/usr/bin/env python3
import os

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FONT_NAME = "Arial"
OUT_PATH  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "modelo_dados_isse.xlsx")

HEADER_FILL   = PatternFill(start_color="1F2A38", end_color="1F2A38", fill_type="solid")
HEADER_FONT   = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
SE_FILL       = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
SS_FILL       = PatternFill(start_color="D5F5E3", end_color="D5F5E3", fill_type="solid")
DA_FILL       = PatternFill(start_color="FEF9E7", end_color="FEF9E7", fill_type="solid")
EXAMPLE_FILL  = PatternFill(start_color="F5EEF8", end_color="F5EEF8", fill_type="solid")
EXAMPLE_FONT  = Font(name=FONT_NAME, italic=True, color="6C3483", size=10)
BODY_FONT     = Font(name=FONT_NAME, size=10)
TITLE_FONT    = Font(name=FONT_NAME, bold=True, size=14, color="1F2A38")
SUBTITLE_FONT = Font(name=FONT_NAME, size=10, italic=True, color="555555")
SECTION_FONT  = Font(name=FONT_NAME, bold=True, size=11, color="1F2A38")
CAT_FONT      = Font(name=FONT_NAME, bold=True, size=10, color="FFFFFF")
NOTE_FONT     = Font(name=FONT_NAME, size=9, italic=True, color="777777")
KEY_FONT      = Font(name=FONT_NAME, bold=True, size=10)
SE_FONT       = Font(name=FONT_NAME, bold=True, size=10, color="1A5276")
SS_FONT       = Font(name=FONT_NAME, bold=True, size=10, color="1E8449")
DA_FONT       = Font(name=FONT_NAME, bold=True, size=10, color="7D6608")

SE_CAT_FILL = PatternFill(start_color="1A5276", end_color="1A5276", fill_type="solid")
SS_CAT_FILL = PatternFill(start_color="1E8449", end_color="1E8449", fill_type="solid")
DA_CAT_FILL = PatternFill(start_color="7D6608", end_color="7D6608", fill_type="solid")

_THIN  = Side(style="thin", color="CCCCCC")
BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

# (nome_coluna, descrição, tipo, exemplo, categoria, fill_col, font_col)
COLUMNS = [
    ("setor",                  "Identificador do bairro/setor (obrigatório)",                       "texto",    "Centro",   "ID",  None,    BODY_FONT),
    ("municipio",              "Nome do município (obrigatório)",                                    "texto",    "Passos",   "ID",  None,    BODY_FONT),
    ("distancia_centro_km",    "Distância do bairro até o centro urbano (km)",                      "número",   1.2,        "SE",  SE_FILL, SE_FONT),
    ("tempo_ape_min",          "Tempo de deslocamento a pé até o centro (minutos)",                  "número",   15,         "SE",  SE_FILL, SE_FONT),
    ("tempo_carro_min",        "Tempo de deslocamento de carro até o centro (minutos)",              "número",   5,          "SE",  SE_FILL, SE_FONT),
    ("tempo_onibus_min",       "Tempo de deslocamento de ônibus até o centro (minutos)",             "número",   12,         "SE",  SE_FILL, SE_FONT),
    ("familias_bolsa_familia", "Nº de famílias beneficiárias do Bolsa Família no bairro",           "inteiro",  45,         "SS",  SS_FILL, SS_FONT),
    ("familias_extrema_pobreza","Nº de famílias Bolsa Família em situação de extrema pobreza",      "inteiro",  18,         "SS",  SS_FILL, SS_FONT),
    ("num_usf",                "Nº de Unidades de Saúde da Família (USF) no bairro",                "inteiro",  1,          "DA",  DA_FILL, DA_FONT),
    ("num_upa",                "Nº de UPAs no bairro",                                              "inteiro",  0,          "DA",  DA_FILL, DA_FONT),
    ("num_hospital_publico",   "Nº de hospitais públicos no bairro",                                "inteiro",  1,          "DA",  DA_FILL, DA_FONT),
    ("num_hospital_privado",   "Nº de hospitais/clínicas privadas no bairro",                       "inteiro",  2,          "DA",  DA_FILL, DA_FONT),
    ("num_escolas_municipais", "Nº de escolas municipais no bairro",                                "inteiro",  2,          "DA",  DA_FILL, DA_FONT),
    ("num_escolas_estaduais",  "Nº de escolas estaduais no bairro",                                 "inteiro",  1,          "DA",  DA_FILL, DA_FONT),
    ("num_escolas_privadas",   "Nº de escolas privadas no bairro",                                  "inteiro",  3,          "DA",  DA_FILL, DA_FONT),
    ("num_cemei",              "Nº de Centros Municipais de Educação Infantil (CEMEI) no bairro",   "inteiro",  1,          "DA",  DA_FILL, DA_FONT),
    ("num_creches",            "Nº de creches (públicas e privadas) no bairro",                     "inteiro",  2,          "DA",  DA_FILL, DA_FONT),
    ("dias_coleta_lixo",       "Nº de dias por semana em que há coleta de lixo no bairro",          "inteiro",  3,          "DA",  DA_FILL, DA_FONT),
    ("linhas_onibus",          "O bairro é atendido por linha de ônibus? (1 = Sim / 0 = Não)",      "0 ou 1",   1,          "DA",  DA_FILL, DA_FONT),
]

INSTRUCTIONS = [
    '1. Preencha os dados na aba "Dados", uma linha por bairro/setor.',
    '2. NÃO altere os nomes das colunas na linha 1 nem sua ordem.',
    '3. A linha 2 é um EXEMPLO (roxo) — apague-a antes de enviar.',
    '4. Todas as 19 colunas são obrigatórias para o cálculo do ISSE.',
    '5. Linhas de ônibus (última coluna): 1 = bairro tem linha, 0 = não tem.',
    '6. Salve como .xlsx e envie pelo formulário da plataforma.',
]

CAT_LEGEND = [
    ("SE — Separação Espacial",    SE_CAT_FILL, "Distância e tempo de deslocamento até o centro urbano. Fonte: Google Maps por bairro."),
    ("SS — Separação Social",      SS_CAT_FILL, "Vulnerabilidade socioeconômica via Cadastro Único / CRAS."),
    ("DA — Desigualdade de Acesso",DA_CAT_FILL, "Equipamentos públicos e privados de saúde, educação e serviços urbanos. Fonte: Secretarias Municipais."),
]

FONTES = [
    ("Separação Espacial (SE)", "distancia_centro_km, tempo_ape_min, tempo_carro_min, tempo_onibus_min",
     "Google Maps — pesquisa manual por bairro. Medir a distância/tempo do ponto central do bairro até a Praça Central do município."),
    ("Separação Social (SS)", "familias_bolsa_familia, familias_extrema_pobreza",
     "Secretaria Municipal de Assistência Social ou CRAS. Dados do Cadastro Único (CadÚnico) por bairro."),
    ("Desigualdade de Acesso — Saúde", "num_usf, num_upa, num_hospital_publico, num_hospital_privado",
     "Secretaria Municipal de Saúde. Confirmar endereços pelo CNES (cnes.datasus.gov.br)."),
    ("Desigualdade de Acesso — Educação", "num_escolas_municipais, num_escolas_estaduais, num_escolas_privadas, num_cemei, num_creches",
     "Secretaria Municipal de Educação. Confirmar pelo Censo Escolar (inep.gov.br)."),
    ("Desigualdade de Acesso — Infraestrutura", "dias_coleta_lixo",
     "Empresa responsável pela coleta de lixo municipal (ex: ECSAM em Passos-MG)."),
    ("Desigualdade de Acesso — Mobilidade", "linhas_onibus",
     "Consultar rotas em moovitapp.com ou junto à empresa de transporte municipal. Valor: 1 = atendido, 0 = não atendido."),
]


def _fmt(cell, *, horizontal="left", wrap=False):
    cell.border = BORDER
    cell.alignment = Alignment(horizontal=horizontal, vertical="center", wrap_text=wrap)


def build_instrucoes(wb):
    ws = wb.active
    ws.title = "Instruções"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 50
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 14

    ws["A1"] = "Modelo de Dados — Plataforma ISSE"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Bastos Filho et al. (2019) — 17 indicadores, 3 categorias, peso 33,33% cada"
    ws["A2"].font = SUBTITLE_FONT

    ws["A4"] = "Como preencher"
    ws["A4"].font = SECTION_FONT
    r = 5
    for line in INSTRUCTIONS:
        ws[f"A{r}"] = line
        ws[f"A{r}"].font = BODY_FONT
        ws.merge_cells(f"A{r}:D{r}")
        r += 1

    r += 1
    ws[f"A{r}"] = "Categorias e cores"
    ws[f"A{r}"].font = SECTION_FONT
    r += 1
    for label, fill, desc in CAT_LEGEND:
        cell = ws.cell(row=r, column=1, value=label)
        cell.font  = CAT_FONT
        cell.fill  = fill
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="left", vertical="center")
        desc_cell = ws.cell(row=r, column=2, value=desc)
        desc_cell.font = BODY_FONT
        ws.merge_cells(f"B{r}:D{r}")
        _fmt(desc_cell, wrap=True)
        ws.row_dimensions[r].height = 28
        r += 1

    r += 1
    ws[f"A{r}"] = "Dicionário de colunas"
    ws[f"A{r}"].font = SECTION_FONT
    r += 1

    for i, h in enumerate(["Coluna", "Descrição", "Tipo", "Exemplo"]):
        cell = ws.cell(row=r, column=i + 1, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        _fmt(cell)
    r += 1

    for col_name, desc, tipo, example, cat, fill, font in COLUMNS:
        c1 = ws.cell(row=r, column=1, value=col_name)
        c1.font = font
        if fill:
            c1.fill = fill
        c2 = ws.cell(row=r, column=2, value=desc)
        c2.font = BODY_FONT
        c3 = ws.cell(row=r, column=3, value=tipo)
        c3.font = BODY_FONT
        c4 = ws.cell(row=r, column=4, value=example)
        c4.font = BODY_FONT
        for c in range(1, 5):
            _fmt(ws.cell(row=r, column=c), wrap=True)
        ws.row_dimensions[r].height = 20
        r += 1

    r += 2
    ws[f"A{r}"] = "Categorização do resultado (Bastos Filho et al., 2019, p. 60)"
    ws[f"A{r}"].font = SECTION_FONT
    r += 1
    for label, cutoff in [
        ("Não Segregada",                  "ISSE < (Média - 1 Desvio-Padrão)"),
        ("Baixa Segregação Socioespacial", "(Média - 1 Desvio-Padrão) ≤ ISSE < Média"),
        ("Média Segregação Socioespacial", "Média ≤ ISSE ≤ (Média + 1 Desvio-Padrão)"),
        ("Alta Segregação Socioespacial",  "ISSE > (Média + 1 Desvio-Padrão)"),
    ]:
        ws.cell(row=r, column=1, value=label).font = KEY_FONT
        ws.cell(row=r, column=2, value=cutoff).font = BODY_FONT
        for c in range(1, 3):
            _fmt(ws.cell(row=r, column=c))
        r += 1


def build_dados(wb):
    ws = wb.create_sheet("Dados")
    ws.sheet_view.showGridLines = False

    cat_colors = {"SE": SE_CAT_FILL, "SS": SS_CAT_FILL, "DA": DA_CAT_FILL, "ID": HEADER_FILL}
    cat_labels = {"SE": "Separação Espacial", "SS": "Separação Social", "DA": "Desigualdade de Acesso", "ID": "Identificação"}

    # Linha 1: rótulo de categoria
    current_cat = None
    cat_start   = 1
    for i, (_, _, _, _, cat, _, _) in enumerate(COLUMNS, start=1):
        if cat != current_cat:
            if current_cat is not None and i - 1 > cat_start:
                ws.merge_cells(start_row=1, start_column=cat_start, end_row=1, end_column=i - 1)
            cell = ws.cell(row=1, column=i, value=cat_labels[cat])
            cell.font  = CAT_FONT
            cell.fill  = cat_colors[cat]
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
            current_cat = cat
            cat_start   = i
    last_col = len(COLUMNS)
    if cat_start <= last_col:
        ws.merge_cells(start_row=1, start_column=cat_start, end_row=1, end_column=last_col)

    # Linha 2: nomes das colunas
    for i, (col_name, _, _, _, cat, fill, _) in enumerate(COLUMNS, start=1):
        letter = get_column_letter(i)
        ws.column_dimensions[letter].width = max(18, len(col_name) + 3)
        cell = ws.cell(row=2, column=i, value=col_name)
        cell.font  = HEADER_FONT
        cell.fill  = cat_colors[cat]
        _fmt(cell, horizontal="center")

    # Linha 3: exemplo
    for i, (_, _, _, example, _, _, _) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=3, column=i, value=example)
        cell.font = EXAMPLE_FONT
        cell.fill = EXAMPLE_FILL
        _fmt(cell, horizontal="center")

    ws.cell(row=3, column=1).comment = Comment(
        "Linha de EXEMPLO — apague antes de enviar a planilha.", "Modelo ISSE"
    )

    # Linhas 4–23: células em branco com bordas
    for r in range(4, 24):
        for c in range(1, len(COLUMNS) + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            cell.font   = BODY_FONT

    ws.freeze_panes = "A3"
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 30


def build_fontes(wb):
    ws = wb.create_sheet("Fontes de Dados")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 60

    ws["A1"] = "Guia de Coleta de Dados — ISSE"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Como e onde obter cada grupo de indicadores"
    ws["A2"].font = SUBTITLE_FONT

    r = 4
    for i, h in enumerate(["Grupo de indicadores", "Colunas afetadas", "Como obter / fonte"]):
        cell = ws.cell(row=r, column=i + 1, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        _fmt(cell, wrap=True)
    r += 1

    for grupo, colunas, como in FONTES:
        ws.cell(row=r, column=1, value=grupo).font   = KEY_FONT
        ws.cell(row=r, column=2, value=colunas).font = BODY_FONT
        ws.cell(row=r, column=3, value=como).font    = BODY_FONT
        for c in range(1, 4):
            _fmt(ws.cell(row=r, column=c), wrap=True)
        ws.row_dimensions[r].height = 48
        r += 1


def main():
    wb = Workbook()
    build_instrucoes(wb)
    build_dados(wb)
    build_fontes(wb)
    wb.active = wb.worksheets[1]
    wb.save(OUT_PATH)
    print(f"Gerado: {OUT_PATH}")


if __name__ == "__main__":
    main()
