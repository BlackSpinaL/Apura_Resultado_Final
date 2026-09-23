import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
import unicodedata
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# CONFIGURAÇÃO
# ============================================================
st.set_page_config(page_title="Boletim por Disciplina", layout="centered")
st.title("📊 Gerador de Boletim em Excel")
st.markdown(
    "Faça o upload do boletim em PDF e da planilha de apuração, "
    "escolha a etapa e clique em **Processar**."
)

# ============================================================
# MAPEAMENTO DE DISCIPLINAS
# ============================================================
MAPA_DISCIPLINAS = {
    "ARTE": "Arte",
    "BIOLOGIA": "Biologia",
    "BIOLOGIA NA PRATICA": "Biologia na Prática",
    "C.DA NATUREZA P ENEM": "C. da Natureza P/ ENEM",
    "CIENCIAS": "Ciências",
    "DESENV. SUSTENTAVEL": "Desenv. Sustentável",
    "ED. PARA PROFISSOES": "Ed. para Profissões",
    "ED.FISICA NA PRATICA": "Ed. Física na Prática",
    "ED.SOCIO.ENS.RELIG.": "Ed. Socio. Ens. Relig.",
    "EDUCACAO FINANCEIRA": "Educação Financeira",
    "EDUCACAO FISICA": "Educação Física",
    "FILOSOFIA": "Filosofia",
    "FISICA": "Física",
    "GEOGRAFIA": "Geografia",
    "HISTORIA": "História",
    "L.INGLESA NA PRATICA": "L. Inglesa na Prática",
    "LIN.PORTUGUESA": "Lin. Portuguesa",
    "LIN.PORTUGUESA 2": "Lin. Portuguesa 2",
    "LINGUA INGLESA": "Língua Inglesa",
    "MAT.E ESTATISTICA": "Mat. e Estatística",
    "MATEMATICA": "Matemática",
    "OFICINA DE TEXTO": "Oficina de Texto",
    "PROJETO DE VIDA": "Projeto de Vida",
    "QUIMICA": "Química",
    "QUIMICA NA PRATICA": "Química na Prática",
    "SOCIOLOGIA": "Sociologia",
}
DISCIPLINAS_VALIDAS = sorted(MAPA_DISCIPLINAS.keys(), key=len, reverse=True)

ETAPAS_CONFIG = {
    "1ª Etapa": {"max": 30, "min": 18},
    "2ª Etapa": {"max": 35, "min": 21},
    "3ª Etapa": {"max": 35, "min": 21},
}
MINIMO_TOTAL = 60

ETAPA_LIMIARES = {
    "1ª Etapa": 18,
    "2ª Etapa": 39,
    "3ª Etapa": 60,
    "Soma das 3 Etapas": 60,
}

# Mapeamento: etapa → coluna da verificação (J=10, K=11, L=12, M=13)
ETAPA_COLUNA = {
    "1ª Etapa": 10,
    "2ª Etapa": 11,
    "3ª Etapa": 12,
    "Soma das 3 Etapas": 13,
}

# ============================================================
# REGEX
# ============================================================
PADRAO_LINHA = re.compile(
    r'^(?P<nome>.+?)\s+(?P<n1>[\d,]+|-)\s+(?P<f1>\d+)\s+(?P<n2>[\d,]+|-)\s+(?P<f2>\d+)\s+'
    r'(?P<n3>[\d,]+|-)\s+(?P<f3>\d+)\s+(?P<tn>[\d,]+|-)\s+(?P<tf>\d+)$'
)
PADRAO_RESTO = re.compile(
    r'^([\d,]+|-)\s+(\d+)\s+([\d,]+|-)\s+(\d+)\s+([\d,]+|-)\s+(\d+)\s+([\d,]+|-)\s+(\d+)$'
)


def _num(s):
    return 0.0 if s == '-' else float(s.replace(',', '.'))


def extrair_linha(linha):
    linha = linha.strip()
    m = PADRAO_LINHA.match(linha)
    if not m:
        return None
    nome_bruto = m.group('nome')
    n1, f1 = _num(m.group('n1')), int(m.group('f1'))
    n2, f2 = _num(m.group('n2')), int(m.group('f2'))
    n3, f3 = _num(m.group('n3')), int(m.group('f3'))
    tn, tf = _num(m.group('tn')), int(m.group('tf'))
    if abs((n1 + n2 + n3) - tn) <= 0.02 and (f1 + f2 + f3) == tf:
        return nome_bruto, n1, n2, n3
    for disc in DISCIPLINAS_VALIDAS:
        if linha.startswith(disc):
            resto = linha[len(disc):].strip()
            m2 = PADRAO_RESTO.match(resto)
            if not m2:
                continue
            n1b, f1b = _num(m2.group(1)), int(m2.group(2))
            n2b, f2b = _num(m2.group(3)), int(m2.group(4))
            n3b, f3b = _num(m2.group(5)), int(m2.group(6))
            tnb, tfb = _num(m2.group(7)), int(m2.group(8))
            if abs((n1b + n2b + n3b) - tnb) <= 0.02 and (f1b + f2b + f3b) == tfb:
                return disc, n1b, n2b, n3b
    return None


# ============================================================
# EXTRAÇÃO DO PDF
# ============================================================
@st.cache_data
def extrair_dados_pdf(pdf_bytes):
    dados = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for pagina in pdf.pages:
            try:
                texto = pagina.extract_text()
            except Exception:
                continue
            if not texto or len(texto.strip()) < 10:
                continue
            padrao_cabecalho = (
                r"MATRÍCULA:\s*(\d+).*?ALUNO:\s*(.*?)\s*PERÍODO LETIVO:.*?TURMA:\s*(\d+)"
            )
            match = re.search(padrao_cabecalho, texto, re.DOTALL | re.IGNORECASE)
            if not match:
                continue
            matricula = match.group(1)
            nome = match.group(2).strip()
            turma = match.group(3).strip()
            for linha in texto.split('\n'):
                if not any(linha.strip().startswith(d) for d in DISCIPLINAS_VALIDAS):
                    continue
                r = extrair_linha(linha)
                if r is None:
                    continue
                disc, n1, n2, n3 = r
                dados.append({
                    "Turma": turma, "Matrícula": matricula, "Aluno": nome,
                    "Disciplina": MAPA_DISCIPLINAS.get(disc, disc),
                    "1ª Etapa": n1, "2ª Etapa": n2, "3ª Etapa": n3,
                })
    return pd.DataFrame(dados)


# ============================================================
# SITUAÇÕES
# ============================================================
def sit_etapa(nota, etapa):
    if nota is None or nota <= 0:
        return "—"
    return "Aprovado" if nota >= ETAPAS_CONFIG[etapa]["min"] else "Reprovado"


def gerar_boletim(df_aluno):
    linhas = []
    for _, row in df_aluno.iterrows():
        n1, n2, n3 = row["1ª Etapa"], row["2ª Etapa"], row["3ª Etapa"]
        linhas.append({
            "Disciplina": row["Disciplina"],
            "Nota 1ª": n1, "Sit. 1ª": sit_etapa(n1, "1ª Etapa"),
            "Nota 2ª": n2, "Sit. 2ª": sit_etapa(n2, "2ª Etapa"),
            "Nota 3ª": n3, "Sit. 3ª": sit_etapa(n3, "3ª Etapa"),
        })
    return pd.DataFrame(linhas)


# ============================================================
# HELPERS DE ESTILO
# ============================================================
def _style_range(ws, r1, c1, r2, c2, fill=None, border=None,
                 alignment=None, font=None):
    """Aplica estilo a um range retangular."""
    for rr in range(r1, r2 + 1):
        for cc in range(c1, c2 + 1):
            cell = ws.cell(row=rr, column=cc)
            if fill is not None:
                cell.fill = fill
            if border is not None:
                cell.border = border
            if alignment is not None:
                cell.alignment = alignment
            if font is not None:
                cell.font = font


def _aplicar_borda_amarela(ws, r1, c1, r2, c2):
    """Aplica borda amarela grossa no perímetro do range."""
    yellow = Side(style="medium", color="FFC000")
    for rr in range(r1, r2 + 1):
        for cc in range(c1, c2 + 1):
            cell = ws.cell(row=rr, column=cc)
            ex = cell.border
            left = yellow if cc == c1 else ex.left
            right = yellow if cc == c2 else ex.right
            top = yellow if rr == r1 else ex.top
            bottom = yellow if rr == r2 else ex.bottom
            cell.border = Border(left=left, right=right, top=top, bottom=bottom)


# ============================================================
# EXCEL — BOLETINS (formato Sugestão)
# ============================================================
def gerar_excel_unico(planilhas, etapa_selecionada="2ª Etapa"):
    """
    Gera Excel no formato da 'Sugestão':
    - Colunas A-J: boletim padrão
    - Colunas J-M: Verificação por Etapa (só a etapa escolhida tem dado)
    - Borda amarela ao redor da área de verificação
    - Linha 'Resultado' com fundo laranja
    - Linha 'Motivo' com fundo cinza (em branco quando Aprovado)
    - Canto sup. direito: 'RELATÓRIO APURA' + COUNTIF
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Boletins"

    # ---- Estilos ----
    title_font = Font(bold=True, size=11, color="FFFFFF")
    title_fill = PatternFill("solid", fgColor="1F4E78")
    hdr_fill = PatternFill("solid", fgColor="1F4E78")
    hdr_font = Font(bold=True, color="FFFFFF", size=11)
    sub_fill = PatternFill("solid", fgColor="D9E1F2")
    sub_font = Font(bold=True, size=10)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center")
    thin = Side(style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    apro_fill = PatternFill("solid", fgColor="C6EFCE")
    rec_fill = PatternFill("solid", fgColor="FFC7CE")
    rep_fill = PatternFill("solid", fgColor="FFC7CE")

    # Cores do rodapé
    orange_fill = PatternFill("solid", fgColor="FCE4D6")
    gray_fill = PatternFill("solid", fgColor="D9D9D9")
    orange_border = Border(
        left=Side(style="thin", color="C55A11"),
        right=Side(style="thin", color="C55A11"),
        top=Side(style="thin", color="C55A11"),
        bottom=Side(style="thin", color="C55A11"),
    )
    gray_border = Border(
        left=Side(style="thin", color="808080"),
        right=Side(style="thin", color="808080"),
        top=Side(style="thin", color="808080"),
        bottom=Side(style="thin", color="808080"),
    )

    # Coluna de verificação (J=10, K=11, L=12, M=13)
    col_verif = ETAPA_COLUNA.get(etapa_selecionada, 11)
    col_verif_letter = get_column_letter(col_verif)

    # Fórmulas para cada etapa
    formulas_etapa = {
        "1ª Etapa": '=IF(B{r}>=18,"Aprovado","Reprovado")',
        "2ª Etapa": '=IF(B{r}+D{r}>=39,"Aprovado","Reprovado")',
        "3ª Etapa": '=IF(B{r}+D{r}+F{r}>=60,"Aprovado","Reprovado")',
        "Soma das 3 Etapas": '=IF(B{r}+D{r}+F{r}>=60,"Aprovado","Reprovado")',
    }
    formula_verif = formulas_etapa.get(etapa_selecionada, formulas_etapa["2ª Etapa"])

    # Rótulo "Resultado da Xª etapa"
    if etapa_selecionada == "1ª Etapa":
        resultado_label = "Resultado da 1ª etapa"
    elif etapa_selecionada == "2ª Etapa":
        resultado_label = "Resultado da 2ª etapa"
    elif etapa_selecionada == "3ª Etapa":
        resultado_label = "Resultado da 3ª etapa"
    else:
        resultado_label = "Resultado Final"

    r = 1
    for aba in planilhas:
        topo_bloco = r

        # ---------- Linha 1 ----------
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
        cell = ws.cell(row=r, column=1,
                       value=f"MATRÍCULA: {aba['matricula']}   |   ALUNO: {aba['aluno']}   |   TURMA: {aba['turma']}")
        cell.font = title_font
        cell.fill = title_fill
        cell.alignment = left
        for col in range(1, 11):
            ws.cell(row=r, column=col).border = border

        # K1 = RELATÓRIO APURA
        ws.cell(row=r, column=11, value="RELATÓRIO APURA").font = Font(bold=True, size=10)
        ws.cell(row=r, column=11).alignment = center
        ws.cell(row=r, column=11).border = border
        ws.cell(row=r, column=12).border = border

        # M1 será preenchido com COUNTIF (só aqui é definido o valor)
        countif_cell = ws.cell(row=r, column=13)
        countif_cell.font = Font(bold=True)
        countif_cell.alignment = center
        countif_cell.border = border

        r += 1
        h1 = r          # linha 2 (cabeçalho grande)
        h2 = r + 1      # linha 3 (sub-cabeçalho)

        # ---------- Cabeçalhos mesclados (linha h1) ----------
        ws.cell(row=h1, column=1, value="Disciplina")
        ws.merge_cells(start_row=h1, start_column=1, end_row=h2, end_column=1)

        ws.cell(row=h1, column=2, value="1ª Etapa – Média de 18 pontos")
        ws.merge_cells(start_row=h1, start_column=2, end_row=h1, end_column=3)

        ws.cell(row=h1, column=4, value="2ª Etapa – Média 21 pontos")
        ws.merge_cells(start_row=h1, start_column=4, end_row=h1, end_column=5)

        ws.cell(row=h1, column=6, value="3ª Etapa – Média 21 pontos")
        ws.merge_cells(start_row=h1, start_column=6, end_row=h1, end_column=7)

        ws.cell(row=h1, column=8, value="Soma das 3 etapas")
        ws.merge_cells(start_row=h1, start_column=8, end_row=h1, end_column=9)

        ws.cell(row=h1, column=10, value="Verificação por Etapa (Mínimo 60% acumulado)")
        ws.merge_cells(start_row=h1, start_column=10, end_row=h1, end_column=13)

        # ---------- Sub-cabeçalhos (linha h2) ----------
        ws.cell(row=h2, column=2, value="Notas")
        ws.cell(row=h2, column=3, value="Situação na 1ª Etapa")
        ws.cell(row=h2, column=4, value="Notas")
        ws.cell(row=h2, column=5, value="Situação")
        ws.cell(row=h2, column=6, value="Notas")
        ws.cell(row=h2, column=7, value="Situação")
        ws.cell(row=h2, column=8, value="Notas")
        ws.cell(row=h2, column=10, value="Situação da 1ª Etapa")
        ws.cell(row=h2, column=11, value="Situação da 2ª Etapa")
        ws.cell(row=h2, column=12, value="Situação da 3ª Etapa")
        ws.cell(row=h2, column=13, value="Situação Final de Ano")

        for row in (h1, h2):
            for col in range(1, 14):
                c = ws.cell(row=row, column=col)
                c.font = hdr_font if row == h1 else sub_font
                c.fill = hdr_fill if row == h1 else sub_fill
                c.alignment = center
                c.border = border

        r = h2 + 1
        first_data = r

        # ---------- Linhas de dados ----------
        df = aba["df"]
        for _, row in df.iterrows():
            ws.cell(row=r, column=1, value=row["Disciplina"]).alignment = left

            ws.cell(row=r, column=2, value=row["Nota 1ª"]).number_format = "0.00"
            c3 = ws.cell(row=r, column=3, value=row["Sit. 1ª"]); c3.alignment = center

            ws.cell(row=r, column=4, value=row["Nota 2ª"]).number_format = "0.00"
            c5 = ws.cell(row=r, column=5, value=row["Sit. 2ª"]); c5.alignment = center

            ws.cell(row=r, column=6, value=row["Nota 3ª"]).number_format = "0.00"
            c7 = ws.cell(row=r, column=7, value=row["Sit. 3ª"]); c7.alignment = center

            # H = Soma das 3 etapas
            soma_cell = ws.cell(row=r, column=8, value=f"=B{r}+D{r}+F{r}")
            soma_cell.number_format = "0.00"
            soma_cell.alignment = center

            # Colunas de verificação J, K, L, M
            # Só a coluna da etapa selecionada recebe fórmula; as outras recebem "---"
            for col_idx in (10, 11, 12, 13):
                if col_idx == col_verif:
                    val = formula_verif.format(r=r)
                else:
                    val = "---"
                cc = ws.cell(row=r, column=col_idx, value=val)
                cc.alignment = center

            # Cores nas situações (C, E, G)
            for cc, val in ((c3, row["Sit. 1ª"]), (c5, row["Sit. 2ª"]), (c7, row["Sit. 3ª"])):
                if val == "Reprovado":
                    cc.fill = rep_fill
                elif val == "Aprovado":
                    cc.fill = apro_fill

            for col in range(1, 14):
                ws.cell(row=r, column=col).border = border
            r += 1

        last_data = r - 1

        # ---------- COUNTIF em M1 ----------
        ws.cell(row=topo_bloco, column=13,
                value=f'=COUNTIF({col_verif_letter}{first_data}:{col_verif_letter}{last_data},"Reprovado")')

        # ---------- Borda amarela grossa ao redor da verificação (J2:M{last_data}) ----------
        _aplicar_borda_amarela(ws, r1=h2, c1=10, r2=last_data, c2=13)

        # ---------- Rodapé: Resultado da Xª etapa ----------
        ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=10)
        label_cell = ws.cell(row=r, column=9, value=resultado_label)
        label_cell.font = Font(bold=True, size=10)
        label_cell.alignment = center

        ws.merge_cells(start_row=r, start_column=11, end_row=r, end_column=13)
        res_cell = ws.cell(
            row=r, column=11,
            value=(f'=IF(M{topo_bloco}>3,"Reprovado",'
                   f'IF(M{topo_bloco}>0,"Recuperação","Aprovado"))')
        )
        res_cell.alignment = center
        res_cell.font = Font(bold=True)

        # Fundo laranja em I:M
        _style_range(ws, r, 9, r, 13, fill=orange_fill, border=orange_border)
        r += 1

        # ---------- Rodapé: Motivo ----------
        ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=10)
        mot_label = ws.cell(row=r, column=9, value="Motivo")
        mot_label.font = Font(bold=True, size=10)
        mot_label.alignment = center

        ws.merge_cells(start_row=r, start_column=11, end_row=r, end_column=13)
        mot_cell = ws.cell(
            row=r, column=11,
            value=(
                f'=IF(COUNTIF({col_verif_letter}{first_data}:{col_verif_letter}{last_data},"Reprovado")=0,"",'
                f'IF(COUNTIF({col_verif_letter}{first_data}:{col_verif_letter}{last_data},"Reprovado")>=4,"Mais que 3 disciplinas",'
                f'IF(COUNTIF({col_verif_letter}{first_data}:{col_verif_letter}{last_data},"Reprovado")=3,"3 disciplinas",'
                f'IF(COUNTIF({col_verif_letter}{first_data}:{col_verif_letter}{last_data},"Reprovado")=2,"2 disciplinas",'
                f'IF(COUNTIF({col_verif_letter}{first_data}:{col_verif_letter}{last_data},"Reprovado")=1,"1 disciplina","")))))'
            )
        )
        mot_cell.alignment = center

        # Fundo cinza em I:M
        _style_range(ws, r, 9, r, 13, fill=gray_fill, border=gray_border)
        r += 3   # espaço entre alunos

    # ---------- Larguras ----------
    ws.column_dimensions["A"].width = 24
    for col in ("B", "D", "F", "H"):
        ws.column_dimensions[col].width = 9
    for col in ("C", "E", "G", "I"):
        ws.column_dimensions[col].width = 20
    ws.column_dimensions["J"].width = 18
    ws.column_dimensions["K"].width = 18
    ws.column_dimensions["L"].width = 18
    ws.column_dimensions["M"].width = 20

    ws.freeze_panes = "A3"
    return wb


# ============================================================
# COMPARAÇÃO APURA x BOLETIM
# ============================================================
def _norm(s):
    if s is None:
        return ""
    try:
        if isinstance(s, float) and pd.isna(s):
            return ""
    except Exception:
        pass
    s = str(s).strip().upper()
    s = "".join(c for c in unicodedata.normalize("NFD", s)
                if unicodedata.category(c) != "Mn")
    return s


def disc_apura_para_pretty(nome):
    n = _norm(nome)
    for k, v in MAPA_DISCIPLINAS.items():
        if _norm(k) == n:
            return v
    for k, v in MAPA_DISCIPLINAS.items():
        if _norm(v) == n:
            return v
    return str(nome).strip()


def parse_disciplinas_apura(s):
    if s is None:
        return []
    try:
        if isinstance(s, float) and pd.isna(s):
            return []
    except Exception:
        pass
    txt = str(s).replace("|", ";").replace(",", ";")
    return [disc_apura_para_pretty(d) for d in txt.split(";") if d.strip()]


def soma_ate_etapa(row, etapa):
    n1 = float(row.get("1ª Etapa") or 0)
    n2 = float(row.get("2ª Etapa") or 0)
    n3 = float(row.get("3ª Etapa") or 0)
    if etapa == "1ª Etapa":
        return n1
    if etapa == "2ª Etapa":
        return n1 + n2
    return n1 + n2 + n3


def gerar_comparacao(apura_df, boletim_df, etapa):
    limiar = ETAPA_LIMIARES[etapa]

    boletim_df = boletim_df.copy()
    boletim_df["_mat"] = (boletim_df["Matrícula"].astype(str)
                          .str.strip().str.replace(r"\.0$", "", regex=True))

    resultados = []
    for _, a in apura_df.iterrows():
        matricula = str(a.get("MATRICULA", "")).strip().replace(".0", "")
        nome = a.get("NOME ALUNO", "")
        turma = a.get("TURMA", "")
        resultado_apura = a.get("RESULTADO", "")
        motivo = a.get("MOTIVO", "")
        disciplinas = parse_disciplinas_apura(a.get("DISCIPLINAS", ""))

        aluno_b = boletim_df[boletim_df["_mat"] == matricula]

        if aluno_b.empty:
            resultados.append({
                "MATRICULA": matricula, "NOME ALUNO": nome, "TURMA": turma,
                "RESULTADO APURA": resultado_apura, "MOTIVO APURA": motivo,
                "DISCIPLINA": "(aluno não encontrado no boletim)",
                "Nota 1ª": None, "Nota 2ª": None, "Nota 3ª": None,
                "Soma": None, "Limiar": limiar, "Status": "SEM DADOS",
            })
            continue

        for disc in disciplinas:
            match = aluno_b[aluno_b["Disciplina"].apply(_norm) == _norm(disc)]
            if match.empty:
                resultados.append({
                    "MATRICULA": matricula, "NOME ALUNO": nome, "TURMA": turma,
                    "RESULTADO APURA": resultado_apura, "MOTIVO APURA": motivo,
                    "DISCIPLINA": disc,
                    "Nota 1ª": None, "Nota 2ª": None, "Nota 3ª": None,
                    "Soma": None, "Limiar": limiar,
                    "Status": "DISCIPLINA NÃO ENCONTRADA",
                })
                continue

            r = match.iloc[0]
            n1 = float(r.get("1ª Etapa") or 0)
            n2 = float(r.get("2ª Etapa") or 0)
            n3 = float(r.get("3ª Etapa") or 0)
            soma = soma_ate_etapa(r, etapa)
            status = "REGULARIZADO" if soma >= limiar else "PENDENTE"

            resultados.append({
                "MATRICULA": matricula, "NOME ALUNO": nome, "TURMA": turma,
                "RESULTADO APURA": resultado_apura, "MOTIVO APURA": motivo,
                "DISCIPLINA": disc,
                "Nota 1ª": round(n1, 2),
                "Nota 2ª": round(n2, 2),
                "Nota 3ª": round(n3, 2),
                "Soma": round(soma, 2),
                "Limiar": limiar,
                "Status": status,
            })

    return pd.DataFrame(resultados)


def gerar_excel_comparacao(df_comp, etapa):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Comparacao {etapa}"

    title_font = Font(bold=True, size=12, color="FFFFFF")
    title_fill = PatternFill("solid", fgColor="1F4E78")
    hdr_font = Font(bold=True, color="FFFFFF", size=10)
    hdr_fill = PatternFill("solid", fgColor="1F4E78")
    ok_fill = PatternFill("solid", fgColor="C6EFCE")
    pend_fill = PatternFill("solid", fgColor="FFC7CE")
    warn_fill = PatternFill("solid", fgColor="FFEB9C")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center")
    thin = Side(style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    cols = list(df_comp.columns)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
    c = ws.cell(row=1, column=1,
                value=f"COMPARAÇÃO APURA x BOLETIM — {etapa}  (Limiar: {ETAPA_LIMIARES[etapa]} pontos)")
    c.font = title_font
    c.fill = title_fill
    c.alignment = center

    for j, col in enumerate(cols, start=1):
        cell = ws.cell(row=2, column=j, value=col)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = center
        cell.border = border

    for i, (_, row) in enumerate(df_comp.iterrows(), start=3):
        for j, col in enumerate(cols, start=1):
            cell = ws.cell(row=i, column=j, value=row[col])
            cell.border = border
            cell.alignment = center if col not in ("NOME ALUNO", "DISCIPLINA", "MOTIVO APURA") else left

        st_idx = cols.index("Status") + 1
        st_val = row["Status"]
        if st_val == "REGULARIZADO":
            ws.cell(row=i, column=st_idx).fill = ok_fill
        elif st_val == "PENDENTE":
            ws.cell(row=i, column=st_idx).fill = pend_fill
        else:
            ws.cell(row=i, column=st_idx).fill = warn_fill

    larguras = {
        "MATRICULA": 12, "NOME ALUNO": 32, "TURMA": 10,
        "RESULTADO APURA": 16, "MOTIVO APURA": 28, "DISCIPLINA": 22,
        "Nota 1ª": 9, "Nota 2ª": 9, "Nota 3ª": 9,
        "Soma": 9, "Limiar": 9, "Status": 18,
    }
    for j, col in enumerate(cols, start=1):
        ws.column_dimensions[get_column_letter(j)].width = larguras.get(col, 14)

    ws.freeze_panes = "A3"

    if not df_comp.empty:
        ws2 = wb.create_sheet("Resumo por Aluno")
        resumo = (df_comp.groupby(["MATRICULA", "NOME ALUNO", "TURMA", "RESULTADO APURA"])
                  .agg(
                      Total_Disciplinas=("DISCIPLINA", "count"),
                      Pendentes=("Status", lambda s: (s == "PENDENTE").sum()),
                      Regularizadas=("Status", lambda s: (s == "REGULARIZADO").sum()),
                  ).reset_index())
        resumo["Situação"] = resumo.apply(
            lambda r: "OK — APURA CONFIRMADA" if r["Pendentes"] == 0 else
                      ("ATENÇÃO — AINDA PENDENTE" if r["Regularizadas"] > 0 else
                       "AINDA REPROVADO"),
            axis=1
        )

        for j, col in enumerate(resumo.columns, start=1):
            hc = ws2.cell(row=1, column=j, value=col)
            hc.font = hdr_font
            hc.fill = hdr_fill
            hc.alignment = center
            hc.border = border

        for i, (_, row) in enumerate(resumo.iterrows(), start=2):
            for j, col in enumerate(resumo.columns, start=1):
                cell = ws2.cell(row=i, column=j, value=row[col])
                cell.border = border
                cell.alignment = center if col != "NOME ALUNO" else left
            if row["Pendentes"] == 0:
                ws2.cell(row=i, column=len(resumo.columns)).fill = ok_fill
            else:
                ws2.cell(row=i, column=len(resumo.columns)).fill = pend_fill

        for j, col in enumerate(resumo.columns, start=1):
            ws2.column_dimensions[get_column_letter(j)].width = \
                30 if col == "NOME ALUNO" else 18

    return wb


# ============================================================
# INTERFACE
# ============================================================
col1, col2 = st.columns(2)
with col1:
    uploaded = st.file_uploader("📂 Upload do Boletim (PDF)", type="pdf", key="pdf")
with col2:
    apura_uploaded = st.file_uploader(
        "📊 Upload da Apuração (Excel)", type=["xlsx", "xls"], key="apura"
    )

etapa_compare = st.selectbox(
    "🎯 Etapa para verificar (aparecerá como coluna única na verificação):",
    ["1ª Etapa", "2ª Etapa", "3ª Etapa", "Soma das 3 Etapas"],
    index=1,
    disabled=(uploaded is None or apura_uploaded is None),
)

# ---------- Botão de processar ----------
processar = st.button(
    "🚀 Processar",
    type="primary",
    use_container_width=True,
    disabled=(uploaded is None or apura_uploaded is None),
    key="btn_processar",
)

if uploaded is None or apura_uploaded is None:
    st.info("👈 Faça o upload do PDF do boletim **e** da planilha de apuração para habilitar o processamento.")

# ============================================================
# PROCESSAMENTO
# ============================================================
if processar and uploaded is not None and apura_uploaded is not None:
    # ---------- 1) Extrai o PDF ----------
    with st.spinner("Processando o boletim..."):
        df_notas = extrair_dados_pdf(uploaded.getvalue())

    if df_notas.empty:
        st.error("❌ Nenhum dado pôde ser extraído do PDF.")
        st.stop()

    # ---------- 2) Gera o Excel dos boletins ----------
    planilhas = []
    for (turma, mat, aluno), df_aluno in df_notas.groupby(
        ["Turma", "Matrícula", "Aluno"], sort=False
    ):
        planilhas.append({
            "matricula": mat, "aluno": aluno, "turma": turma,
            "df": gerar_boletim(df_aluno),
        })

    with st.spinner("Gerando Excel dos boletins..."):
        wb = gerar_excel_unico(planilhas, etapa_selecionada=etapa_compare)
        buf = io.BytesIO()
        wb.save(buf)
        st.session_state["wb_boletins_bytes"] = buf.getvalue()
        st.session_state["n_alunos"] = len(planilhas)

    # ---------- 3) Lê a planilha de apuração (com normalização de colunas) ----------
    try:
        apura_df = pd.read_excel(apura_uploaded)
    except Exception as e:
        st.error(f"Erro ao ler a planilha de apuração: {e}")
        st.stop()

    # Normaliza nomes das colunas: remove BOM, espaços extras e deixa UPPER
    apura_df.columns = (
        apura_df.columns.astype(str)
        .str.replace("\ufeff", "", regex=False)
        .str.strip()
        .str.upper()
    )

    obrig = {"MATRICULA", "NOME ALUNO", "RESULTADO", "DISCIPLINAS"}
    faltando = obrig - set(apura_df.columns)
    if faltando:
        st.error(
            f"❌ Colunas ausentes na planilha de apuração: **{sorted(faltando)}**\n\n"
            f"Colunas encontradas: {sorted(apura_df.columns.tolist())}"
        )
        st.stop()

    # ---------- 4) Gera a comparação ----------
    with st.spinner("Comparando disciplinas..."):
        df_comp = gerar_comparacao(apura_df, df_notas, etapa_compare)

    if df_comp.empty:
        st.warning("Nenhuma comparação pôde ser gerada.")
        st.stop()

    with st.spinner("Gerando Excel da comparação..."):
        wb_comp = gerar_excel_comparacao(df_comp, etapa_compare)
        buf2 = io.BytesIO()
        wb_comp.save(buf2)
        st.session_state["df_comp_cache"] = df_comp
        st.session_state["wb_comp_bytes"] = buf2.getvalue()
        st.session_state["etapa_cache"] = etapa_compare
        st.session_state["wb_comp_name"] = (
            f"comparacao_apura_{etapa_compare.replace(' ', '_').replace('ª','a')}.xlsx"
        )

    st.success("✅ Processamento concluído!")

# ============================================================
# DOWNLOADS (persistem após processar)
# ============================================================
if "wb_boletins_bytes" in st.session_state:
    st.markdown("---")
    st.subheader("📄 Boletins (formato Sugestão)")
    st.caption(f"Alunos processados: {st.session_state.get('n_alunos', 0)} — "
               f"Etapa selecionada: **{st.session_state.get('etapa_cache','')}**")
    st.download_button(
        label="📥 Baixar Excel dos Boletins",
        data=st.session_state["wb_boletins_bytes"],
        file_name="boletim_por_disciplina.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="dl_boletins",
    )

if "wb_comp_bytes" in st.session_state:
    st.markdown("---")
    st.subheader(f"🔍 Comparação com Apuração — {st.session_state['etapa_cache']}")

    df_comp = st.session_state["df_comp_cache"]
    st.dataframe(df_comp, use_container_width=True, hide_index=True)

    total = len(df_comp)
    ok = int((df_comp["Status"] == "REGULARIZADO").sum())
    pend = int((df_comp["Status"] == "PENDENTE").sum())
    c1, c2, c3 = st.columns(3)
    c1.metric("Total aluno-disciplina", total)
    c2.metric("✅ Regularizados", ok)
    c3.metric("⚠️ Pendentes", pend)

    st.download_button(
        label=f"📥 Baixar Excel da Comparação ({st.session_state['etapa_cache']})",
        data=st.session_state["wb_comp_bytes"],
        file_name=st.session_state["wb_comp_name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="dl_comparacao",
    )