import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import os
import pdfplumber
from datetime import datetime

# Configuração da página e tema
st.set_page_config(page_title="18º BPM — Extrajornada", page_icon="🛡️", layout="centered")

# 🎨 MODO ESCURO (CSS Personalizado)
st.markdown("""
<style>
    /* Fundo Escuro */
    .stApp {
        background-color: #0E1117;
        color: #E0E6ED;
    }
    /* Estilização dos Containers e Caixas */
    div[data-testid="stFileUploader"], div[data-testid="stTextInput"] {
        background-color: #1E222D;
        border-radius: 10px;
        padding: 10px;
        border: 1px solid #2E364A;
    }
    /* Botões operacionais */
    .stButton>button {
        background-color: #002060;
        color: #FFFFFF;
        font-weight: bold;
        border-radius: 8px;
        border: 1px solid #1E3A8A;
        width: 100%;
        padding: 10px;
    }
    .stButton>button:hover {
        background-color: #1E40AF;
        border-color: #3B82F6;
    }
    /* Títulos */
    h1, h2, h3 {
        color: #F3F4F6 !important;
    }
</style>
""", unsafe_allow_html=True)

# 🏷️ ASSINATURA NO CANTINHO INFERIOR DIREITO
st.markdown("""
<div style="position: fixed; bottom: 15px; right: 20px; text-align: right; color: #9CA3AF; font-size: 12px; font-family: sans-serif; z-index: 999999; line-height: 1.4; background-color: rgba(14, 17, 23, 0.85); padding: 6px 12px; border-radius: 6px; border: 1px solid #2E364A;">
    Desenvolvido por:<br>
    <strong style="color: #60A5FA; font-size: 13px;">Nathan Wenzel</strong>
</div>
""", unsafe_allow_html=True)

# 🔒 CONFIGURAÇÃO DA SENHA DE ACESSO
SENHA_CORRETA = "deusa"

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

# Tela de Login
if not st.session_state.autenticado:
    if os.path.exists("brasao.png"):
        col1, col2, col3 = st.columns(3)
        with col2:
            st.image("brasao.png", width=130)

    st.title("🔒 Acesso Restrito — 18º BPM")
    st.write("Digite a senha de acesso para utilizar o Gerador de Relatórios Extrajornada.")
    
    senha_input = st.text_input("Senha de acesso:", type="password")
    
    if st.button("Entrar"):
        if senha_input == SENHA_CORRETA:
            st.session_state.autenticado = True
            st.success("Acesso liberado!")
            st.rerun()
        else:
            st.error("Senha incorreta! Verifique e tente novamente.")
    st.stop()

# =========================================================
# FUNÇÕES DE SUPORTE E FORMATOS DE DATA
# =========================================================

DIAS_SEMANA = {
    0: "segunda-feira", 1: "terça-feira", 2: "quarta-feira",
    3: "quinta-feira", 4: "sexta-feira", 5: "sábado", 6: "domingo"
}

MESES = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
    5: "maio", 6: "junho", 7: "julho", 8: "agosto",
    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
}

def formatar_data_para_tela_inicial(val_str):
    """Converte '23/09/2026' para '23 de setembro de 2026 (quarta-feira)' APENAS para a caixa da tela inicial"""
    if not val_str or str(val_str).strip().lower() in ['none', 'nan', '']:
        return "23 de setembro de 2026 (quarta-feira)"
    
    # Pega apenas o texto limpo da data (evita passar lista para o strptime)
    s = str(val_str).strip().split()[0]
    
    for fmt in ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y"]:
        try:
            dt = datetime.strptime(s, fmt)
            dia = dt.day
            mes = MESES[dt.month]
            ano = dt.year
            dia_sem = DIAS_SEMANA[dt.weekday()]
            return f"{dia} de {mes} de {ano} ({dia_sem})"
        except ValueError:
            pass
            
    return str(val_str).strip()

def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), hex_color)
    shd.set(qn('w:val'), 'clear')
    tcPr.append(shd)

def ler_arquivo_pdf(file_bytes):
    data = []
    with pdfplumber.open(file_bytes) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    row_clean = [str(cell).replace('\n', ' ').strip() if cell is not None else '' for cell in row]
                    if any(row_clean):
                        data.append(row_clean)
    if not data:
        return pd.DataFrame()
    
    header_idx = 0
    for idx, row in enumerate(data):
        row_str = " ".join(row).upper()
        if "VOLCHER" in row_str or "VOUCHER" in row_str or "CIDADE" in row_str or "DATA" in row_str or "N°" in row_str:
            header_idx = idx
            break
            
    headers = [str(h).strip() for h in data[header_idx]]
    rows = data[header_idx + 1:]
    return pd.DataFrame(rows, columns=headers)

def processar_dataframe(df):
    cols = {str(c).upper().strip(): c for c in df.columns}
    
    def encontrar_coluna(termos, df_cols):
        for termo in termos:
            for col_upper, col_orig in df_cols.items():
                if termo in col_upper:
                    return col_orig
        return None

    col_volcher = encontrar_coluna(["VOLCHER", "VOUCHER", "VOLCHE", "VOUCHE", "N°", "Nº"], cols)
    col_cidade = encontrar_coluna(["CIDADE", "MUNICÍPIO", "LOCAL"], cols)
    col_data = encontrar_coluna(["DATA", "DIA"], cols)
    col_hora = encontrar_coluna(["HORA"], cols)

    return df, col_volcher, col_cidade, col_data, col_hora

def gerar_relatorio_word(df_escala, data_extenso="23 de setembro de 2026 (quarta-feira)"):
    """Gera o documento Word exatamente igual ao modelo v8 aprovado"""
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.7)
        section.right_margin = Inches(0.7)

    # Inserção do Brasão no Documento Word (se existir)
    if os.path.exists("brasao.png"):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_after = Pt(4)
        run_img = p_img.add_run()
        run_img.add_picture("brasao.png", width=Inches(0.9))

    # Cabeçalho Institucional Oficial PM
    p_hdr = doc.add_paragraph()
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_hdr.paragraph_format.space_after = Pt(2)
    p_hdr.paragraph_format.line_spacing = 1.15
    
    r_estado = p_hdr.add_run("ESTADO DO PARANÁ\nPOLÍCIA MILITAR\n2º COMANDO REGIONAL DE POLÍCIA MILITAR\n18° BATALHÃO DE POLÍCIA MILITAR\n")
    r_estado.bold = True
    r_estado.font.size = Pt(11)
    r_estado.font.name = "Arial"
    
    r_prog = p_hdr.add_run("PROGRAMAÇÃO EXTRAJORNADA VOLUNTÁRIA\n")
    r_prog.bold = True
    r_prog.font.size = Pt(12)
    r_prog.font.name = "Arial"
    r_prog.font.color.rgb = RGBColor(0, 32, 96)

    p_data = doc.add_paragraph()
    p_data.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_data.paragraph_format.space_after = Pt(6)
    r_dt = p_data.add_run(data_extenso)
    r_dt.bold = True
    r_dt.font.size = Pt(11)
    r_dt.font.name = "Arial"

    p_obs = doc.add_paragraph()
    p_obs.paragraph_format.space_after = Pt(12)
    r_obs = p_obs.add_run("* As equipes ficarão à disposição do CPU, Adjunto e COPOM para atendimentos de ocorrências. Na ausência de ocorrências, deverão seguir o cartão de programa.")
    r_obs.italic = True
    r_obs.font.size = Pt(9.5)
    r_obs.font.name = "Arial"

    df, col_v, col_c, col_d, col_h = processar_dataframe(df_escala)

    group_cols = [c for c in [col_v, col_c] if c is not None]
    
    for group_idx, (chaves, grupo) in enumerate(df.groupby(group_cols if group_cols else df.columns)):
        primeiro = grupo.iloc[0]
        
        volcher_raw = str(primeiro.get(col_v, "")).replace('.0', '').replace('None', '').replace('nan', '').strip() if col_v else ""
        volcher_val = volcher_raw if volcher_raw else str(group_idx + 1)
        
        cidade_val = str(primeiro.get(col_c, "")).replace('None', '').replace('nan', '').strip() if col_c else ""
        data_str = str(primeiro.get(col_d, "23/09/2026")).strip() if col_d else "23/09/2026"
        
        hora_str = str(primeiro.get(col_h, "18:00")).strip() if col_h else "18:00"
        if "às" not in hora_str.lower() and "das" not in hora_str.lower():
            hora_str = f"Das {hora_str} às 23:59"
        
        table = doc.add_table(rows=5, cols=2)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        for row in table.rows:
            row.cells[0].width = Inches(2.0)
            row.cells[1].width = Inches(4.5)

        campos = [
            ("CIDADE/VOLCHER", f"{cidade_val} - VOLCHER - {volcher_val}"),
            ("DATA", data_str),
            ("LOCAL", cidade_val),
            ("HORÁRIO", hora_str)
        ]

        for i, (label, val) in enumerate(campos):
            r = table.rows[i]
            
            c0 = r.cells[0]
            p0 = c0.paragraphs[0]
            p0.paragraph_format.space_after = Pt(2)
            p0.paragraph_format.space_before = Pt(2)
            r0 = p0.add_run(label)
            r0.bold = True
            r0.font.name = "Arial"
            r0.font.size = Pt(10)
            set_cell_background(c0, "D9E1F2")

            c1 = r.cells[1]
            p1 = c1.paragraphs[0]
            p1.paragraph_format.space_after = Pt(2)
            p1.paragraph_format.space_before = Pt(2)
            r1 = p1.add_run(val)
            r1.bold = True if label in ["CIDADE/VOLCHER", "LOCAL"] else False
            r1.font.name = "Arial"
            r1.font.size = Pt(10)
            set_cell_background(c1, "FFFFFF")

        # Quadro de observações sem amarelo
        r4 = table.rows[4]
        c0 = r4.cells[0]
        c1 = r4.cells[1]
        c0.merge(c1)
        p_obs_tbl = c0.paragraphs[0]
        p_obs_tbl.paragraph_format.space_after = Pt(3)
        p_obs_tbl.paragraph_format.space_before = Pt(3)
        p_obs_tbl.paragraph_format.line_spacing = 1.15

        r_l1 = p_obs_tbl.add_run("A equipe ficará a Disposição do COPOM e CPU ou Adjunto.\n")
        r_l1.font.name = "Arial"
        r_l1.font.size = Pt(9.5)

        r_l2 = p_obs_tbl.add_run("SISGCOP 61076\n")
        r_l2.bold = True
        r_l2.font.name = "Arial"
        r_l2.font.size = Pt(9.5)

        r_l3 = p_obs_tbl.add_run("    • Equipe deverá fazer contato com o Adjunto ao assumir serviço.\n")
        r_l3.font.name = "Arial"
        r_l3.font.size = Pt(9.5)

        r_l4 = p_obs_tbl.add_run("    • A equipe além realizar o atendimento de ocorrências, deverá realizar o Patrulhamento Ostensivo e Preventivo na área designada para atuar.")
        r_l4.font.name = "Arial"
        r_l4.font.size = Pt(9.5)

        set_cell_background(c0, "FAFAFA")

        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# Interface Principal
if os.path.exists("brasao.png"):
    col1, col2, col3 = st.columns(3)
    with col2:
        st.image("brasao.png", width=120)

st.title("🛡️ Gerador de Relatório Extrajornada")
st.caption("18º Batalhão de Polícia Militar — PMPR")
st.write("Envie a tabela da escala em **.pdf**, **.xlsx** ou **.csv**.")

arquivo = st.file_uploader("Envie o arquivo da escala", type=["xlsx", "csv", "pdf"])

if arquivo:
    ext = arquivo.name.split(".")[-1].lower()
    if ext == "xlsx":
        df = pd.read_excel(arquivo)
    elif ext == "csv":
        df = pd.read_csv(arquivo)
    elif ext == "pdf":
        df = ler_arquivo_pdf(arquivo)

    if not df.empty:
        st.success(f"Arquivo carregado com sucesso! {len(df)} registros encontrados.")
        
        # Sugestão inteligente de data APENAS para preenchimento da caixinha na tela inicial
        _, _, _, col_d, _ = processar_dataframe(df)
        data_sugerida_tela = "23 de setembro de 2026 (quarta-feira)"
        if col_d and not df[col_d].dropna().empty:
            primeira_data_val = str(df[col_d].dropna().iloc[0]).strip()
            data_sugerida_tela = formatar_data_para_tela_inicial(primeira_data_val)

        data_cabecalho = st.text_input("Data para o cabeçalho do relatório", value=data_sugerida_tela)
        
        if st.button("Gerar Documento Word"):
            docx_bytes = gerar_relatorio_word(df, data_cabecalho)
            st.download_button("📥 Baixar Relatório Preenchido (.docx)", docx_bytes, "RELATORIO_EXTRAJORNADA.docx")
    else:
        st.error("Não foi possível extrair dados da tabela. Verifique o arquivo enviado.")
