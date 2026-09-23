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
st.set_page_config(page_title="Boletim x Apura", layout="centered")
st.title("📊 Apura Resultado: Análise 📊")
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

LIMIAR_ETAPA = {
    "1ª Etapa": 18,
    "2ª Etapa": 39,
    "3ª Etapa": 60,
    "Soma das 3 Etapas": 60,
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
# HELPERS DE NORMALIZAÇÃO
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


def _nome_aba_seguro(nome):
    """Garante nome válido de sheet (≤31 chars, sem : \\ / ? * [ ])."""
    nome = str(nome).strip() or "Turma"
    for ch in [":", "\\", "/", "?", "*", "[", "]"]:
        nome = nome.replace(ch, "-")
    return nome[:31]


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
# DESENHO DE UM BLOCO DE ALUNO EM UMA ABA
# ============================================================
def _desenhar_bloco_aluno(ws, r_inicial, aba, reprovadas_norm,
                          etapa_selecionada, formula_etapa, limiar,
                          label_situacao, label_resultado, label_comparacao,
                          estilos):
    """
    Desenha o bloco (cabeçalho + dados + rodapé) de um aluno a partir
    de r_inicial (linha com 1 espaço em branco já contabilizado).
    Retorna a próxima linha livre.
    """
    (
        title_font, title_fill, hdr_fill, hdr_font,
        sub_fill, sub_font, analysis_fill, analysis_font,
        dark_fill, countif_fill, center, left, border,
        apro_fill, rep_fill, orange_fill, gray_fill, yellow_fill,
    ) = estilos

    r = r_inicial + 1
    topo = r

    # ---------- Linha do topo ----------
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
    cell = ws.cell(row=r, column=1,
                   value=f"MATRÍCULA: {aba['matricula']}   |   ALUNO: {aba['aluno']}   |   TURMA: {aba['turma']}")
    cell.font = title_font
    cell.fill = title_fill
    cell.alignment = center

    ws.merge_cells(start_row=r, start_column=10, end_row=r, end_column=11)
    an = ws.cell(row=r, column=10, value="Análise do Boletim x Apura")
    an.font = analysis_font
    an.fill = analysis_fill
    an.alignment = center

    ws.cell(row=r, column=12).fill = analysis_fill

    m_cell = ws.cell(row=r, column=13)
    m_cell.font = Font(bold=True)
    m_cell.alignment = center
    m_cell.fill = countif_fill

    for col in range(1, 14):
        ws.cell(row=r, column=col).border = border

    # ---------- Cabeçalhos ----------
    r += 1
    h1 = r
    h2 = r + 1

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

    ws.cell(row=h1, column=10, value=label_situacao)
    ws.merge_cells(start_row=h1, start_column=10, end_row=h2, end_column=10)

    ws.cell(row=h1, column=11, value="Relatório do Apura")
    ws.merge_cells(start_row=h1, start_column=11, end_row=h2, end_column=11)

    ws.cell(row=h1, column=12, value=label_comparacao)
    ws.merge_cells(start_row=h1, start_column=12, end_row=h2, end_column=12)

    # M — SEM MERGE (evita conflito)

    ws.cell(row=h2, column=2, value="Notas")
    ws.cell(row=h2, column=3, value="Situação na 1ª Etapa")
    ws.cell(row=h2, column=4, value="Notas")
    ws.cell(row=h2, column=5, value="Situação")
    ws.cell(row=h2, column=6, value="Notas")
    ws.cell(row=h2, column=7, value="Situação")
    ws.cell(row=h2, column=8, value="Notas")

    for row in (h1, h2):
        for col in range(1, 14):
            c = ws.cell(row=row, column=col)
            c.font = hdr_font if row == h1 else sub_font
            c.fill = hdr_fill if row == h1 else sub_fill
            c.alignment = center
            c.border = border

    for row in (h1, h2):
        ws.cell(row=row, column=7).fill = gray_fill
        ws.cell(row=row, column=9).fill = dark_fill
        ws.cell(row=row, column=13).fill = dark_fill

    r = h2 + 1
    first_data = r

    # ---------- Dados ----------
    df = aba["df"]
    for _, row in df.iterrows():
        n1 = row["Nota 1ª"]
        n2 = row["Nota 2ª"]
        n3 = row["Nota 3ª"]

        ws.cell(row=r, column=1, value=row["Disciplina"]).alignment = left

        ws.cell(row=r, column=2, value=n1).number_format = "0.00"
        c3 = ws.cell(row=r, column=3, value=row["Sit. 1ª"]); c3.alignment = center

        ws.cell(row=r, column=4, value=n2).number_format = "0.00"
        c5 = ws.cell(row=r, column=5, value=row["Sit. 2ª"]); c5.alignment = center

        ws.cell(row=r, column=6, value=n3).number_format = "0.00"
        c7 = ws.cell(row=r, column=7, value=row["Sit. 3ª"]); c7.alignment = center

        soma_cell = ws.cell(row=r, column=8, value=f"=B{r}+D{r}+F{r}")
        soma_cell.number_format = "0.00"
        soma_cell.alignment = center

        j_cell = ws.cell(row=r, column=10, value=formula_etapa.format(r=r))
        j_cell.alignment = center

        disc_norm = _norm(row["Disciplina"])
        k_val = "Reprovado" if disc_norm in reprovadas_norm else "Aprovado"
        k_cell = ws.cell(row=r, column=11, value=k_val)
        k_cell.alignment = center
        k_cell.fill = rep_fill if k_val == "Reprovado" else apro_fill

        l_cell = ws.cell(row=r, column=12,
                         value=f'=IF(J{r}=K{r},"Ok","Divergente")')
        l_cell.alignment = center

        if etapa_selecionada == "1ª Etapa":
            soma_etapa = n1
        elif etapa_selecionada == "2ª Etapa":
            soma_etapa = n1 + n2
        else:
            soma_etapa = n1 + n2 + n3
        j_result = "Aprovado" if soma_etapa >= limiar else "Reprovado"
        if j_result != k_val:
            l_cell.fill = yellow_fill

        for cc, val in ((c3, row["Sit. 1ª"]), (c5, row["Sit. 2ª"]), (c7, row["Sit. 3ª"])):
            if val == "Reprovado":
                cc.fill = rep_fill
            elif val == "Aprovado":
                cc.fill = apro_fill
            elif val == "—":
                cc.fill = gray_fill

        ws.cell(row=r, column=9).fill = dark_fill
        ws.cell(row=r, column=9).border = border
        ws.cell(row=r, column=9).alignment = center
        ws.cell(row=r, column=13).fill = dark_fill
        ws.cell(row=r, column=13).border = border
        ws.cell(row=r, column=13).alignment = center

        for col in range(1, 14):
            ws.cell(row=r, column=col).border = border
        r += 1

    last_data = r - 1

    # ---------- COUNTIF em M (linha topo) ----------
    ws.cell(row=topo, column=13,
            value=f'=COUNTIF(J{first_data}:J{last_data},"Reprovado")')

    # ---------- Rodapé: Resultado ----------
    ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=10)
    rc = ws.cell(row=r, column=9, value=label_resultado)
    rc.font = Font(bold=True)
    rc.alignment = center

    ws.merge_cells(start_row=r, start_column=11, end_row=r, end_column=13)
    res_cell = ws.cell(
        row=r, column=11,
        value=(f'=IF(COUNTIF(J{first_data}:J{last_data},"Reprovado")>3,"Reprovado",'
               f'IF(COUNTIF(J{first_data}:J{last_data},"Reprovado")>0,"Recuperação","Aprovado"))')
    )
    res_cell.alignment = center
    res_cell.font = Font(bold=True)
    for col in range(9, 14):
        ws.cell(row=r, column=col).fill = orange_fill
        ws.cell(row=r, column=col).border = border
    r += 1

    # ---------- Rodapé: Motivo ----------
    ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=10)
    mc = ws.cell(row=r, column=9, value="Motivo")
    mc.font = Font(bold=True)
    mc.alignment = center

    ws.merge_cells(start_row=r, start_column=11, end_row=r, end_column=13)
    mot_cell = ws.cell(
        row=r, column=11,
        value=(
            f'=IF(COUNTIF(J{first_data}:J{last_data},"Reprovado")=0,"",'
            f'IF(COUNTIF(J{first_data}:J{last_data},"Reprovado")>=4,"Mais que 3 disciplinas",'
            f'IF(COUNTIF(J{first_data}:J{last_data},"Reprovado")=3,"3 disciplinas",'
            f'IF(COUNTIF(J{first_data}:J{last_data},"Reprovado")=2,"2 disciplinas",'
            f'IF(COUNTIF(J{first_data}:J{last_data},"Reprovado")=1,"1 disciplina","")))))'
        )
    )
    mot_cell.alignment = center
    for col in range(9, 14):
        ws.cell(row=r, column=col).fill = orange_fill
        ws.cell(row=r, column=col).border = border
    r += 3

    return r


# ============================================================
# EXCEL ÚNICO — uma aba por turma
# ============================================================
def gerar_excel_unico(planilhas, apura_map, etapa_selecionada="2ª Etapa"):
    wb = openpyxl.Workbook()
    # Remove a aba default
    wb.remove(wb.active)

    # ---- Paleta ----
    title_font = Font(bold=True, size=11, color="FFFFFF")
    title_fill = PatternFill("solid", fgColor="1F4E78")
    hdr_fill = PatternFill("solid", fgColor="1F4E78")
    hdr_font = Font(bold=True, color="FFFFFF", size=11)
    sub_fill = PatternFill("solid", fgColor="D9E1F2")
    sub_font = Font(bold=True, size=10)

    analysis_fill = PatternFill("solid", fgColor="404040")
    analysis_font = Font(bold=True, color="FFFFFF", size=10)
    dark_fill = PatternFill("solid", fgColor="404040")
    countif_fill = PatternFill("solid", fgColor="FFEB9C")

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center")
    thin = Side(style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    apro_fill = PatternFill("solid", fgColor="C6EFCE")
    rep_fill = PatternFill("solid", fgColor="FFC7CE")
    orange_fill = PatternFill("solid", fgColor="FCE4D6")
    gray_fill = PatternFill("solid", fgColor="BFBFBF")
    yellow_fill = PatternFill("solid", fgColor="FFFF00")

    estilos = (
        title_font, title_fill, hdr_fill, hdr_font,
        sub_fill, sub_font, analysis_fill, analysis_font,
        dark_fill, countif_fill, center, left, border,
        apro_fill, rep_fill, orange_fill, gray_fill, yellow_fill,
    )

    # ---- Fórmula da etapa ----
    formulas_etapa = {
        "1ª Etapa": '=IF(B{r}>=18,"Aprovado","Reprovado")',
        "2ª Etapa": '=IF(B{r}+D{r}>=39,"Aprovado","Reprovado")',
        "3ª Etapa": '=IF(B{r}+D{r}+F{r}>=60,"Aprovado","Reprovado")',
        "Soma das 3 Etapas": '=IF(B{r}+D{r}+F{r}>=60,"Aprovado","Reprovado")',
    }
    formula_etapa = formulas_etapa.get(etapa_selecionada, formulas_etapa["2ª Etapa"])
    limiar = LIMIAR_ETAPA[etapa_selecionada]

    if etapa_selecionada == "Soma das 3 Etapas":
        label_situacao = "Situação Final"
        label_resultado = "Resultado Final"
        label_comparacao = "Comparação: Final x Apura"
    else:
        label_situacao = f"Situação da {etapa_selecionada}"
        label_resultado = f"Resultado da {etapa_selecionada}"
        label_comparacao = f"Comparação: {etapa_selecionada} x Apura"

    # ---- Agrupa por turma ----
    turmas = {}
    for p in planilhas:
        t = str(p["turma"]).strip() or "SEM TURMA"
        turmas.setdefault(t, []).append(p)

    # Ordena as turmas alfabeticamente
    for turma_nome in sorted(turmas.keys()):
        alunos_da_turma = turmas[turma_nome]

        nome_aba = _nome_aba_seguro(turma_nome)
        # Garante nome único
        base = nome_aba
        contador = 1
        while nome_aba in wb.sheetnames:
            contador += 1
            sufixo = f"_{contador}"
            nome_aba = (base[:31 - len(sufixo)] + sufixo)
        ws = wb.create_sheet(title=nome_aba)

        r = 0  # começa em 0 porque o helper já faz `r_inicial + 1`
        for aba in alunos_da_turma:
            mat_norm = str(aba["matricula"]).strip()
            reprovadas_norm = apura_map.get(mat_norm, set())
            r = _desenhar_bloco_aluno(
                ws, r, aba, reprovadas_norm,
                etapa_selecionada, formula_etapa, limiar,
                label_situacao, label_resultado, label_comparacao,
                estilos,
            )

        # ---- Larguras ----
        ws.column_dimensions["A"].width = 24
        for col in ("B", "D", "F", "H"):
            ws.column_dimensions[col].width = 9
        for col in ("C", "E"):
            ws.column_dimensions[col].width = 18
        ws.column_dimensions["G"].width = 12
        ws.column_dimensions["I"].width = 8
        ws.column_dimensions["J"].width = 20
        ws.column_dimensions["K"].width = 18
        ws.column_dimensions["L"].width = 22
        ws.column_dimensions["M"].width = 10

    # Se nenhuma turma foi criada, cria uma vazia
    if not wb.sheetnames:
        wb.create_sheet(title="Vazio")

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
    "🎯 Etapa para verificar (aparecerá na coluna 'Situação da Xª Etapa'):",
    ["1ª Etapa", "2ª Etapa", "3ª Etapa", "Soma das 3 Etapas"],
    index=1,
    disabled=(uploaded is None or apura_uploaded is None),
)

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
    with st.spinner("Processando o boletim..."):
        df_notas = extrair_dados_pdf(uploaded.getvalue())

    if df_notas.empty:
        st.error("❌ Nenhum dado pôde ser extraído do PDF.")
        st.stop()

    try:
        apura_df = pd.read_excel(apura_uploaded)
    except Exception as e:
        st.error(f"Erro ao ler a planilha de apuração: {e}")
        st.stop()

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

    apura_map = {}
    for _, a in apura_df.iterrows():
        mat = str(a.get("MATRICULA", "")).strip().replace(".0", "")
        disc_list = parse_disciplinas_apura(a.get("DISCIPLINAS", ""))
        apura_map[mat] = set(_norm(d) for d in disc_list)

    planilhas = []
    for (turma, mat, aluno), df_aluno in df_notas.groupby(
        ["Turma", "Matrícula", "Aluno"], sort=False
    ):
        planilhas.append({
            "matricula": mat, "aluno": aluno, "turma": turma,
            "df": gerar_boletim(df_aluno),
        })

    with st.spinner("Gerando Excel 'Boletim x Apura' (uma aba por turma)..."):
        wb = gerar_excel_unico(planilhas, apura_map, etapa_selecionada=etapa_compare)
        buf = io.BytesIO()
        wb.save(buf)
        st.session_state["wb_bytes"] = buf.getvalue()
        st.session_state["n_alunos"] = len(planilhas)
        st.session_state["etapa_cache"] = etapa_compare
        # Lista as turmas presentes
        turmas_presentes = sorted({str(p["turma"]).strip() for p in planilhas})
        st.session_state["turmas_cache"] = turmas_presentes

    st.success(
        f"✅ Processamento concluído! {len(planilhas)} aluno(s) em "
        f"{len(st.session_state['turmas_cache'])} turma(s)."
    )

# ============================================================
# DOWNLOAD
# ============================================================
if "wb_bytes" in st.session_state:
    st.markdown("---")
    st.subheader("📄 Arquivo gerado: **Boletim x Apura**")
    st.caption(
        f"Alunos: {st.session_state.get('n_alunos', 0)} — "
        f"Etapa analisada: **{st.session_state.get('etapa_cache','')}**"
    )
    turmas = st.session_state.get("turmas_cache", [])
    if turmas:
        st.caption(f"Abas geradas (uma por turma): **{', '.join(turmas)}**")

    nome_arq = f"boletim_x_apura_{st.session_state.get('etapa_cache','').replace(' ', '_').replace('ª','a')}.xlsx"
    st.download_button(
        label="📥 Baixar 'Boletim x Apura'",
        data=st.session_state["wb_bytes"],
        file_name=nome_arq,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="dl_boletim_apura",
    )
