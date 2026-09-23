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
        if "VOLUNTÁRIO" in row_str or "NOME" in row_str or "CPF" in row_str or "VOUCHER" in row_str:
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

    col_voucher = encontrar_coluna(["VOUCHER"], cols) or df.columns[0]
    col_cidade = encontrar_coluna(["CIDADE", "MUNICÍPIO", "LOCAL"], cols) or df.columns[1]
    col_grad = encontrar_coluna(["POSTO", "GRAD", "POSTO/GRAD"], cols)
    col_nome = encontrar_coluna(["VOLUNTÁRIO", "PM", "NOME"], cols)
    col_cpf = encontrar_coluna(["CPF"], cols)
    col_data = encontrar_coluna(["DATA"], cols)
    col_hora = encontrar_coluna(["HORA", "HORÁRIO"], cols)

    return df, col_voucher, col_cidade, col_grad, col_nome, col_cpf, col_data, col_hora

def gerar_relatorio_word(df_escala, data_extenso="23 de setembro de 2026 (quarta-feira)"):
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

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

    df, col_v, col_c, col_g, col_n, col_cpf, col_d, col_h = processar_dataframe(df_escala)

    for (voucher, cidade), grupo in df.groupby([col_v, col_c]):
        primeiro = grupo.iloc[0]
        data_str = str(primeiro.get(col_d, "23/09/2026")) if col_d else "23/09/2026"
        hora_str = str(primeiro.get(col_h, "18:00 às 23:59")) if col_h else "18:00 às 23:59"
        if "às" not in hora_str.lower() and "das" not in hora_str.lower():
            hora_str = f"Das {hora_str} às 23:59"
        
        table = doc.add_table(rows=7 + len(grupo), cols=2)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        campos = [
            ("CIDADE/VOUCHER", f"{cidade} - VOUCHER {voucher}"),
            ("DATA", data_str),
            ("LOCAL", str(cidade)),
            ("HORÁRIO", hora_str)
        ]

        for i, (label, val) in enumerate(campos):
            r = table.rows[i]
            r.cells[0].paragraphs[0].add_run(label).bold = True
            r.cells[1].paragraphs[0].add_run(val)
            set_cell_background(r.cells[0], "F2F2F2")
            set_cell_background(r.cells[1], "F2F2F2")

        r4 = table.rows[4]
        c_m = r4.cells[0]
        c_m.merge(r4.cells[1])
        c_m.paragraphs[0].add_run("A equipe ficará a Disposição do COPOM e CPU ou Adjunto. | SISGCOP 61076\nEquipe deverá fazer contato com o Adjunto ao assumir serviço. A equipe além de realizar o atendimento de ocorrências, deverá realizar o Patrulhamento Ostensivo e Preventivo na área designada para atuar.")
        set_cell_background(c_m, "FAFAFA")

        r5 = table.rows[5]
        c_pm = r5.cells[0]
        c_pm.merge(r5.cells[1])
        c_pm.paragraphs[0].add_run("POLICIAIS MILITARES ESCALADOS (PM VOLUNTÁRIOS)").bold = True
        set_cell_background(c_pm, "D9E1F2")

        for idx, (_, row) in enumerate(grupo.iterrows()):
            r_pm = table.rows[6 + idx]
            grad = str(row.get(col_g, "PM")) if col_g else "PM"
            nome = str(row.get(col_n, "")) if col_n else ""
            cpf = str(row.get(col_cpf, "")) if col_cpf else ""
            r_pm.cells[0].paragraphs[0].add_run(grad).bold = True
            r_pm.cells[1].paragraphs[0].add_run(f"{nome} — CPF: {cpf}")

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
        st.error("Não foi possível extrair a tabela do arquivo PDF. Verifique se o PDF contém texto selecionável de tabela.")
