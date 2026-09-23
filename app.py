import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import pdfplumber

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
    col_data = encontrar_coluna(["DATA"], cols)
    col_hora = encontrar_coluna(["HORA"], cols)

    return df, col_volcher, col_cidade, col_data, col_hora

def gerar_relatorio_word(df_escala, data_extenso="23 de setembro de 2026 (quarta-feira)"):
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.7)
        section.right_margin = Inches(0.7)

    # Cabeçalho Institucional
    p_hdr = doc.add_paragraph()
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p_hdr.add_run("ESTADO DO PARANÁ POLÍCIA MILITAR\n2º COMANDO REGIONAL DE POLÍCIA MILITAR\n18° BATALHÃO DE POLÍCIA MILITAR\n")
    r1.bold = True
    
    r2 = p_hdr.add_run("PROGRAMAÇÃO EXTRAJORNADA VOLUNTÁRIA\n")
    r2.bold = True
    r2.font.color.rgb = RGBColor(0, 32, 96)

    p_data = doc.add_paragraph()
    p_data.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_data.add_run(data_extenso).bold = True

    p_obs = doc.add_paragraph()
    p_obs.add_run("* As equipes ficarão à disposição do CPU, Adjunto e COPOM para atendimentos de ocorrências. Na ausência de ocorrências, deverão seguir o cartão de programa.").italic = True

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
        
        campos = [
            ("CIDADE/VOLCHER", f"{cidade_val} - VOLCHER - {volcher_val}"),
            ("DATA", data_str),
            ("LOCAL", cidade_val),
            ("HORÁRIO", hora_str)
        ]

        for i, (label, val) in enumerate(campos):
            r = table.rows[i]
            
            p0 = r.cells[0].paragraphs[0]
            p0.add_run(label).bold = True
            
            p1 = r.cells[1].paragraphs[0]
            p1.add_run(val)
            
            set_cell_background(r.cells[0], "E9EEF4")
            set_cell_background(r.cells[1], "F8FAFC")

        # Texto fixo de observação simplificado
        r4 = table.rows[4]
        c0 = r4.cells[0]
        c1 = r4.cells[1]
        c0.merge(c1)
        p_obs_tbl = c0.paragraphs[0]
        p_obs_tbl.text = "A equipe ficará a Disposição do COPOM e CPU ou Adjunto. | SISGCOP 61076"
        set_cell_background(c0, "FAFAFA")

        doc.add_paragraph()

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

st.title("🛡️ Gerador de Relatório Extrajornada PM")
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
        data_cabecalho = st.text_input("Data para o cabeçalho do relatório", "23 de setembro de 2026 (quarta-feira)")
        
        if st.button("Gerar Documento Word"):
            docx_bytes = gerar_relatorio_word(df, data_cabecalho)
            st.download_button("📥 Baixar Relatório Preenchido (.docx)", docx_bytes, "RELATORIO_EXTRAJORNADA.docx")
    else:
        st.error("Não foi possível extrair dados da tabela. Verifique o arquivo enviado.")
