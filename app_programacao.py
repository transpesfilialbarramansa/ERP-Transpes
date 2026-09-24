import io
import base64
import datetime
import hashlib
import json
import psycopg
import pandas as pd
import streamlit as st
import re
import openpyxl
import urllib.parse

# Importações para geração de PDF e Excel
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

def formatar_cpf(valor: str) -> str:
    if not valor:
        return ""
    digits = re.sub(r'\D', '', str(valor))
    if len(digits) == 11:
        return f"{digits[:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:]}"
    return str(valor).strip()

def formatar_telefone(valor: str) -> str:
    if not valor:
        return ""
    digits = re.sub(r'\D', '', str(valor))
    if len(digits) == 11:
        return f"({digits[:2]}) {digits[2:7]}-{digits[7:]}"
    elif len(digits) == 10:
        return f"({digits[:2]}) {digits[2:6]}-{digits[6:]}"
    return str(valor).strip()

# ==========================================
# 1. CONFIGURAÇÃO DA PÁGINA (Sempre o primeiro comando)
# ==========================================
st.set_page_config(
    page_title="TRANSPES - Sistema de Gestão",
    page_icon="logo_transpes.png",
    layout="wide"
)

# ==========================================
# APLICAÇÃO DE ESTILO E LAYOUT ULTRACOMPACTO (CSS CUSTOMIZADO)
# ==========================================
st.markdown("""
    <style>
    /* Fundo geral e densidade de tela */
    .stApp {
        background-color: #F4F6F9;
    }

    /* Redução de padding das seções principais do Streamlit */
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
        max-width: 100% !important;
    }

    /* Estilização e largura da Sidebar */
    [data-testid="stSidebar"] {
        background-color: #0B2136 !important;
        min-width: 220px !important;
        max-width: 220px !important;
    }

    [data-testid="stSidebar"] * {
        color: #E2E8F0 !important;
    }

    /* Botões do Menu Lateral Compactos */
    [data-testid="stSidebar"] .stButton > button {
        background-color: transparent !important;
        color: #CBD5E1 !important;
        border: none !important;
        text-align: left !important;
        justify-content: flex-start !important;
        font-weight: 500 !important;
        font-size: 0.82rem !important;
        border-radius: 4px !important;
        padding: 5px 10px !important;
        margin-bottom: -4px !important;
        width: 100% !important;
    }

    /* Botão Selecionado */
    [data-testid="stSidebar"] .stButton > button[kind="primary"],
    [data-testid="stSidebar"] .stButton > button:focus {
        background-color: #D99B26 !important;
        color: #0B2136 !important;
        font-weight: bold !important;
    }

    /* Hover Sidebar */
    [data-testid="stSidebar"] .stButton > button:hover {
        background-color: #1A365D !important;
        color: #FFFFFF !important;
    }

    /* Títulos compactos */
    h1 {
        color: #0F172A !important;
        font-weight: 700 !important;
        font-size: 1.25rem !important;
        margin-bottom: 0.5rem !important;
        padding-bottom: 0px !important;
    }

    h2, h3 {
        color: #0F172A !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        margin-top: 0.4rem !important;
        margin-bottom: 0.3rem !important;
    }

    /* Inputs e Rótulos ultracompactos */
    div[data-baseweb="input"], div[data-baseweb="select"] {
        min-height: 32px !important;
    }

    input, select, textarea {
        font-size: 0.8rem !important;
        padding: 4px 8px !important;
    }

    label {
        font-size: 0.78rem !important;
        font-weight: 600 !important;
        margin-bottom: 2px !important;
    }

    /* Métricas Compactas */
    div[data-testid="metric-container"] {
        background-color: #FFFFFF !important;
        border-radius: 6px !important;
        padding: 6px 12px !important;
        border: 1px solid #E2E8F0 !important;
        box-shadow: 0px 1px 3px rgba(0, 0, 0, 0.03) !important;
    }

    div[data-testid="metric-container"] label {
        font-size: 0.72rem !important;
    }

    div[data-testid="metric-container"] div[data-testid="stMetricValue"] {
        font-size: 1.1rem !important;
        font-weight: 700 !important;
    }

    /* Tabelas densas e compactas */
    div[data-testid="stDataFrame"] {
        font-size: 0.75rem !important;
    }

    /* Botões Form / Ação reduzidos */
    .stButton > button {
        font-size: 0.82rem !important;
        padding: 4px 10px !important;
    }

    /* Divisores compactos */
    hr {
        margin: 0.6rem 0 !important;
    }

    header[data-testid="stHeader"] {
        background: transparent !important;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# CONEXÃO E BANCO DE DADOS (SUPABASE / POSTGRESQL / NEON)
# ==========================================

def get_connection():
    db_url = st.secrets["DB_URL"]
    if db_url.startswith("postgres://"):
        db_url = db_url.replace("postgres://", "postgresql://", 1)
    return psycopg.connect(db_url)

def hash_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

FORNECEDORES_INICIAIS = [
    ('99', 'ALEXANDRE SILVANO DE MELO', 'ALEMAO MUNCK', 'CAETITE', 'BA', '28.332.244/0001-90', '77 9925-3952'),
    ('108', 'TERMOSOL CONSTRUTORA E COMERCIO LTDA', 'CONSTRUTORA TERMOSOL', 'CAETITE', 'BA', '06.872.066/0001-58', '77 3454-3659'),
    ('43', 'AGM CONSTRUTORA LTDA', 'AGM CONSTRUTORA', 'CALMON', 'BA', '11.051.592/0001-97', '74 36212416'),
    ('151', 'PEZINHO GUINCHO LOCACAO E TRANSPORTE LTDA - ME', 'PEZINHO GUINCHO LOCACAO E TRANSPORTE', 'FEIRA DE SANTANA', 'BA', '13.332.190/0001-96', '75 36234413')
]

@st.cache_resource
def init_db():
    with get_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS usuarios (
                    id SERIAL PRIMARY KEY,
                    usuario TEXT UNIQUE NOT NULL,
                    senha TEXT NOT NULL,
                    nome TEXT NOT NULL,
                    perfil TEXT NOT NULL
                )
            """)

            usuarios_iniciais = [
                ("admin", "Transpes@1966", "Administrador do Sistema", "ADMIN"),
                ("deyves.teixeira", "Transpes@26", "Deyves Teixeira", "ADMIN"),
                ("felipe.cesario", "Transpes@26", "Felipe Cesario", "ADMIN")
            ]

            for usr, pwd, nome, perf in usuarios_iniciais:
                cursor.execute("SELECT COUNT(*) FROM usuarios WHERE usuario = %s", (usr,))
                if cursor.fetchone()[0] == 0:
                    cursor.execute(
                        "INSERT INTO usuarios (usuario, senha, nome, perfil) VALUES (%s, %s, %s, %s)",
                        (usr, hash_senha(pwd), nome, perf)
                    )
                else:
                    cursor.execute(
                        "UPDATE usuarios SET senha = %s WHERE usuario = %s",
                        (hash_senha(pwd), usr)
                    )

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS cargas (
                    id SERIAL PRIMARY KEY,
                    numero_carga TEXT UNIQUE NOT NULL,
                    status TEXT DEFAULT 'PENDENTE_PROGRAMACAO',
                    data_criacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

            novas_colunas_cargas = [
                ("numero_set", "TEXT"),
                ("origens_json", "TEXT"),
                ("destinos_json", "TEXT"),
                ("notas_fiscais_comercial", "TEXT"),
                ("receita_frete", "REAL DEFAULT 0.00"),
                ("receita_pedagio", "REAL DEFAULT 0.00"),
                ("receita_taxa_descarga", "REAL DEFAULT 0.00"),
                ("data_previsao_coleta", "TEXT"),
                ("data_previsao_entrega", "TEXT"),
                ("numero_tp", "TEXT"),
                ("nome_motorista", "TEXT"),
                ("cpf_motorista", "TEXT"),
                ("telefone_motorista", "TEXT"),
                ("tipo_motorista", "TEXT"),
                ("tipo_veiculo", "TEXT"),
                ("quantidade_eixos", "INTEGER"),
                ("placa_cavalo", "TEXT"),
                ("placa_carreta", "TEXT"),
                ("valor_rpa", "REAL DEFAULT 0.00"),
                ("data_coleta", "TEXT"),
                ("previsao_descarga", "TEXT"),
                ("observacoes_programacao", "TEXT"),
                ("tipo_pedagio", "TEXT"),
                ("plataforma", "TEXT"),
                ("vinculo", "TEXT")
            ]

            for col_nome, col_tipo in novas_colunas_cargas:
                cursor.execute(f"ALTER TABLE cargas ADD COLUMN IF NOT EXISTS {col_nome} {col_tipo};")

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS carga_expedicao (
                    carga_id INTEGER PRIMARY KEY REFERENCES cargas(id) ON DELETE CASCADE,
                    numero_cte TEXT,
                    numero_viagem TEXT,
                    numero_mdfe TEXT,
                    numero_contrato TEXT,
                    valor_pedagio_pago REAL DEFAULT 0.00,
                    notas_fiscais_expedicao TEXT,
                    data_saida_filial TEXT NOT NULL,
                    observacoes_expedicao TEXT,
                    data_emissao TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS carga_operacional (
                    carga_id INTEGER PRIMARY KEY REFERENCES cargas(id) ON DELETE CASCADE,
                    fornecedor_descarga TEXT NOT NULL,
                    custo_fornecedor_descarga REAL DEFAULT 0.00,
                    data_agendamento_descarga TEXT,
                    forma_pagamento TEXT,
                    prazo_pagamento_dias INTEGER,
                    data_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS carga_administracao (
                    carga_id INTEGER PRIMARY KEY REFERENCES cargas(id) ON DELETE CASCADE,
                    valor_adiantamento REAL,
                    valor_saldo REAL,
                    comprovante_entregue INTEGER DEFAULT 0,
                    data_liberacao_saldo TEXT,
                    status_pagamento_saldo TEXT DEFAULT 'PENDENTE'
                )
            """)

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS fornecedores (
                    id SERIAL PRIMARY KEY,
                    nr_contrato TEXT,
                    razao_social TEXT NOT NULL,
                    nome_fantasia TEXT,
                    local_atendimento TEXT,
                    uf TEXT,
                    cpf_cnpj TEXT,
                    contato TEXT,
                    data_cadastro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

            cursor.execute("SELECT COUNT(*) FROM fornecedores")
            if cursor.fetchone()[0] == 0:
                for row in FORNECEDORES_INICIAIS:
                    cursor.execute("""
                        INSERT INTO fornecedores (nr_contrato, razao_social, nome_fantasia, local_atendimento, uf, cpf_cnpj, contato)
                        VALUES (%s, %s, %s, %s, %s, %s, %s)
                    """, row)

            conn.commit()

init_db()

# ==========================================
# FUNÇÕES AUXILIARES DE EXPORTAÇÃO E IMAGEM
# ==========================================
def get_base64_image(image_path):
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except Exception:
        return None

def gerar_numero_carga_novo():
    with get_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute("SELECT COUNT(*) FROM cargas")
            qtd = cursor.fetchone()[0] + 1
    ano = datetime.datetime.now().year
    return f"TRP-{ano}-{qtd:03d}"

def processar_json_lista(json_str):
    if not json_str:
        return []
    try:
        data = json.loads(json_str)
        return data if isinstance(data, list) else [str(data)]
    except Exception:
        return [s.strip() for s in str(json_str).split(",") if s.strip()]

def formatar_origens_destinos(json_str, apenas_locais=False):
    itens = processar_json_lista(json_str)
    if not itens:
        return "-"
    res = []
    for item in itens:
        if isinstance(item, dict):
            cli = item.get("cliente", "")
            cid = item.get("cidade", "")
            uf = item.get("estado", "")
            if apenas_locais:
                res.append(f"{cid}/{uf}")
            else:
                res.append(f"{cli} ({cid}/{uf})")
        else:
            res.append(str(item))
    return " | ".join(res)

def limpar_formato_json_lista(valor, separador=", "):
    if not valor or str(valor).strip() in ["-", "None", "", "[]"]:
        return "-"
    if isinstance(valor, str) and (valor.startswith("[") and valor.endswith("]")):
        try:
            dados = json.loads(valor)
            if isinstance(dados, list):
                itens = [str(i).strip() for i in dados if str(i).strip()]
                return separador.join(itens) if itens else "-"
        except Exception:
            pass
    limpo = str(valor).replace("[", "").replace("]", "").replace('"', '').replace("'", "").strip()
    return limpo if limpo else "-"

def gerar_excel_geral(df_dados):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_dados.to_excel(writer, index=False, sheet_name="Detalhamento de Cargas")
        workbook = writer.book
        worksheet = writer.sheets["Detalhamento de Cargas"]
        worksheet.views.sheetView[0].showGridLines = True

        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
        from openpyxl.utils import get_column_letter

        fill_cabecalho = PatternFill(start_color="0F2A4A", end_color="0F2A4A", fill_type="solid")
        font_cabecalho = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        font_corpo = Font(name="Calibri", size=9)

        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")

        borda_fina = Side(border_style="thin", color="D9D9D9")
        borda_caixa = Border(left=borda_fina, right=borda_fina, top=borda_fina, bottom=borda_fina)

        cols_moeda = ["Receita Total (R$)", "RPA", "Pedágio Pago", "Custo Descarga", "Custo Total (R$)", "Margem (R$)"]
        cols_centro = ["Nº TP", "Nº SET", "Data Programada", "Status", "CPF Motorista", "Tipo Motorista", "Placa Cavalo", "Placa Carreta", "CT-e", "Viagem", "MDF-e", "Contrato", "Nota Fiscal", "Tipo de Pedágio", "Plataforma", "Vínculo", "Data Agendamento Descarga", "Data Pagamento Saldo"]

        for col_num, col_name in enumerate(df_dados.columns, start=1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.fill = fill_cabecalho
            cell.font = font_cabecalho
            cell.alignment = align_center

        for row_num in range(2, len(df_dados) + 2):
            for col_num, col_name in enumerate(df_dados.columns, start=1):
                cell = worksheet.cell(row=row_num, column=col_num)
                cell.font = font_corpo
                cell.border = borda_caixa
                if col_name in cols_moeda:
                    cell.alignment = align_right
                    cell.number_format = 'R$ #,##0.00'
                elif col_name == "Margem (%)":
                    cell.alignment = align_right
                    cell.number_format = '0%'
                elif col_name in cols_centro:
                    cell.alignment = align_center
                else:
                    cell.alignment = align_left

        for col in worksheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                if len(val_str) > max_len:
                    max_len = len(val_str)
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 10)

    output.seek(0)
    return output.getvalue()

def gerar_pdf_geral(df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=15, leftMargin=15, topMargin=15, bottomMargin=15)
    elements = []
    styles = getSampleStyleSheet()
    
    style_title = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=11, leading=13, alignment=1, textColor=colors.HexColor("#0f2a4a"), spaceAfter=8)
    style_cell = ParagraphStyle('Cell', parent=styles['Normal'], fontName='Helvetica', fontSize=5.5, leading=6.5, alignment=0)
    style_header = ParagraphStyle('Header', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=6, leading=7, textColor=colors.white, alignment=1)

    elements.append(Paragraph("Relatório Geral de Cargas - Transpes", style_title))

    cols_pdf = ["Nº TP", "Nº SET", "Data Programada", "Status", "Motorista", "Origem", "Destino", "Receita Total (R$)", "Custo Total (R$)", "Margem (R$)", "Margem (%)"]
    df_pdf = df.copy()
    
    for c in ["Receita Total (R$)", "Custo Total (R$)", "Margem (R$)"]:
        if c in df_pdf.columns:
            df_pdf[c] = df_pdf[c].apply(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if isinstance(x, (int, float)) else str(x))
    
    if "Margem (%)" in df_pdf.columns:
        df_pdf["Margem (%)"] = df_pdf["Margem (%)"].apply(lambda x: f"{x:.1f}%" if isinstance(x, (int, float)) else str(x))

    table_data = [[Paragraph(col, style_header) for col in cols_pdf]]

    for _, row in df_pdf.iterrows():
        linha = []
        for col in cols_pdf:
            val = str(row[col]) if pd.notnull(row[col]) and str(row[col]) != "" else "-"
            linha.append(Paragraph(val, style_cell))
        table_data.append(linha)

    col_widths = [45, 55, 55, 60, 95, 150, 150, 65, 65, 65, 45]
    pdf_table = Table(table_data, colWidths=col_widths, repeatRows=1)
    pdf_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0f2a4a")),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#cccccc")),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8f9fa")]),
    ]))

    elements.append(pdf_table)
    doc.build(elements)
    return buffer.getvalue()

# ==========================================
# DIÁLOGOS / POP-UPS
# ==========================================
@st.dialog("💼 Dados Comerciais Registrados!")
def exibir_popup_comercial(num_carga):
    st.success(f"A carga **{num_carga}** foi cadastrada pelo Comercial com sucesso!")
    if st.button("OK / Nova Carga", type="primary", use_container_width=True):
        st.session_state["exibir_modal_comercial"] = False
        st.session_state["com_form_version"] = st.session_state.get("com_form_version", 0) + 1
        st.rerun()

@st.dialog("📌 Programação Concluída!")
def exibir_popup_programacao(num_tp):
    st.success(f"A carga TP **{num_tp}** foi programada com sucesso!")
    if st.button("OK / Próxima Programação", type="primary", use_container_width=True):
        st.session_state["exibir_modal_programacao"] = False
        st.session_state["prog_form_version"] = st.session_state.get("prog_form_version", 0) + 1
        st.rerun()

@st.dialog("📦 Carga Expedida!")
def exibir_popup_expedicao(num_tp):
    st.success(f"A carga TP **{num_tp}** foi expedida e está **EM TRÂNSITO**!")
    if st.button("OK / Próxima Carga", type="primary", use_container_width=True):
        st.session_state["exibir_modal_expedicao"] = False
        st.session_state["exp_form_version"] = st.session_state.get("exp_form_version", 0) + 1
        st.rerun()

@st.dialog("⚙️ Agendamento Confirmado!")
def exibir_popup_operacional(num_tp):
    st.success(f"A carga TP **{num_tp}** teve os dados operacionais salvos!")
    if st.button("OK / Próxima Carga", type="primary", use_container_width=True):
        st.session_state["exibir_modal_operacional"] = False
        st.session_state["op_form_version"] = st.session_state.get("op_form_version", 0) + 1
        st.rerun()

@st.dialog("💼 Processo Finalizado!")
def exibir_popup_administracao(num_tp):
    st.success(f"A carga TP **{num_tp}** teve o acerto concluído!")
    if st.button("OK / Concluir", type="primary", use_container_width=True):
        st.session_state["exibir_modal_administracao"] = False
        st.session_state["adm_form_version"] = st.session_state.get("adm_form_version", 0) + 1
        st.rerun()

@st.dialog("➕ Novo Fornecedor")
def modal_cadastrar_fornecedor():
    with st.form("form_modal_fornecedor"):
        f1, f2 = st.columns(2)
        nr_contrato = f1.text_input("Nr. Contrato").strip()
        cpf_cnpj = f2.text_input("CPF/CNPJ*").strip()

        razao_social = st.text_input("Razão Social*").strip().upper()
        nome_fantasia = st.text_input("Nome Fantasia").strip().upper()

        l1, l2 = st.columns([3, 1])
        local_atendimento = l1.text_input("Local Atendimento*").strip().upper()
        uf = l2.text_input("UF*").strip().upper()

        contato = st.text_input("Contato / Telefone").strip()

        btn_salvar_forn = st.form_submit_button("💾 Salvar Fornecedor", type="primary", use_container_width=True)

    if btn_salvar_forn:
        if not razao_social or not local_atendimento or not uf:
            st.error("Preencha os campos obrigatórios.")
        else:
            try:
                with get_connection() as conn:
                    with conn.cursor() as cursor:
                        cursor.execute("""
                            INSERT INTO fornecedores (nr_contrato, razao_social, nome_fantasia, local_atendimento, uf, cpf_cnpj, contato)
                            VALUES (%s, %s, %s, %s, %s, %s, %s)
                        """, (nr_contrato, razao_social, nome_fantasia, local_atendimento, uf, cpf_cnpj, contato))
                        conn.commit()
                st.success("Fornecedor cadastrado com sucesso!")
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao cadastrar fornecedor: {e}")

# ==========================================
# AUTENTICAÇÃO E SESSÃO
# ==========================================
if "logado" not in st.session_state:
    st.session_state["logado"] = False

if not st.session_state["logado"]:
    bg_base64 = get_base64_image("fundo_transpes.png")
    bg_style = f"background-image: url('data:image/jpeg;base64,{bg_base64}'); background-size: cover; background-position: center; background-repeat: no-repeat; background-attachment: fixed;" if bg_base64 else "background-color: #0E2F56;"

    logo_base64 = get_base64_image("logo_transpes.png")
    logo_html = f'<img src="data:image/png;base64,{logo_base64}" style="max-width: 150px; margin-bottom: 8px;">' if logo_base64 else '<div style="color: #0E2F56; font-size: 22px; font-weight: bold;">ERP Transpes</div>'

    st.markdown(
        f"""
        <style>
        [data-testid="stSidebar"], [data-testid="stHeader"] {{ display: none; }}
        .stApp {{ {bg_style} }}
        [data-testid="stMainBlockContainer"] {{ max-width: 380px !important; padding-top: 4rem !important; margin: auto !important; }}
        .login-card {{ background: rgba(255, 255, 255, 0.95); padding: 20px 18px; border-radius: 12px; box-shadow: 0 6px 24px rgba(0, 0, 0, 0.3); text-align: center; margin-bottom: 15px; }}
        </style>
        """,
        unsafe_allow_html=True
    )

    st.markdown(f'<div class="login-card">{logo_html}<div style="color: #555; font-size: 12px;">Acesse o sistema</div></div>', unsafe_allow_html=True)

    with st.form("form_login"):
        usuario_input = st.text_input("Usuário", placeholder="Digite seu usuário")
        senha_input = st.text_input("Senha", type="password", placeholder="Digite sua senha")
        btn_login = st.form_submit_button("Entrar", use_container_width=True, type="primary")

    if btn_login:
        with get_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(
                    "SELECT usuario, nome, perfil FROM usuarios WHERE usuario = %s AND senha = %s",
                    (usuario_input.strip().lower(), hash_senha(senha_input))
                )
                usr = cursor.fetchone()
        
        if usr:
            st.session_state["logado"] = True
            st.session_state["usuario"] = usr[0]
            st.session_state["usuario_nome"] = usr[1]
            st.session_state["usuario_perfil"] = usr[2]
            st.rerun()
        else:
            st.error("Usuário ou senha incorretos.")

    st.stop()

# ==========================================
# NAVEGAÇÃO E SIDEBAR ESTILIZADA
# ==========================================
if "menu" not in st.session_state:
    st.session_state["menu"] = "Visão Geral"

menu_selecionado = st.session_state["menu"]
perfil = st.session_state.get("usuario_perfil", "ADMIN")

opcoes_perfil = ["Visão Geral"]
if perfil == "ADMIN":
    opcoes_perfil.extend(["Comercial", "Programação", "Expedição", "Operacional", "Administração", "Fornecedores", "Excluir Cargas", "Usuários"])
elif perfil == "COMERCIAL":
    opcoes_perfil.append("Comercial")
elif perfil == "PROGRAMACAO":
    opcoes_perfil.append("Programação")
elif perfil == "EXPEDICAO":
    opcoes_perfil.append("Expedição")
elif perfil == "OPERACIONAL":
    opcoes_perfil.extend(["Operacional", "Fornecedores"])
elif perfil == "ADMINISTRATIVO":
    opcoes_perfil.extend(["Administração", "Fornecedores"])

logo_base64_sidebar = get_base64_image("logo_transpes.png")

with st.sidebar:
    # 1. Logo + Cabeçalho na Sidebar
    if logo_base64_sidebar:
        st.markdown(f"""
            <div style="text-align: center; padding: 5px 0px 10px 0px;">
                <img src="data:image/png;base64,{logo_base64_sidebar}" style="max-width: 120px; height: auto; margin-bottom: 5px;">
                <h3 style="color: #FFFFFF; margin: 0; font-weight: 800; font-size: 1.1rem; letter-spacing: 1px;">TRANSPES</h3>
                <p style="color: #94A3B8; font-size: 0.65rem; margin: 0;">SISTEMA DE GESTÃO</p>
            </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
            <div style="text-align: center; padding: 5px 0px 10px 0px;">
                <h3 style="color: #FFFFFF; margin: 0; font-weight: 800; font-size: 1.1rem; letter-spacing: 1px;">TRANSPES</h3>
                <p style="color: #94A3B8; font-size: 0.65rem; margin: 0;">SISTEMA DE GESTÃO</p>
            </div>
        """, unsafe_allow_html=True)

    # 2. Navegação
    menus_icones = {
        "Visão Geral": "📊 Visão Geral",
        "Comercial": "💲 Comercial",
        "Programação": "📅 Programação",
        "Expedição": "📦 Expedição",
        "Operacional": "⚙️ Operacional",
        "Administração": "💼 Administração",
        "Fornecedores": "🚚 Fornecedores",
        "Excluir Cargas": "🗑️ Excluir Cargas",
        "Usuários": "👥 Usuários"
    }

    for item in opcoes_perfil:
        if st.button(menus_icones.get(item, item), type="primary" if menu_selecionado == item else "secondary", use_container_width=True):
            st.session_state["menu"] = item
            st.rerun()

    # 3. Rodapé do Usuário Logado
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown(f"<small style='color: #CBD5E1;'>👤 <b>{st.session_state.get('usuario_nome', 'Admin')}</b> ({st.session_state.get('usuario_perfil', 'ADMIN')})</small>", unsafe_allow_html=True)

    if st.button("🚪 Sair", use_container_width=True):
        st.session_state.clear()
        st.rerun()

# ==========================================
# PÁGINAS DO SISTEMA
# ==========================================

# 1. VISÃO GERAL
if menu_selecionado == "Visão Geral":
    st.title("📊 Visão Geral e Relatórios")

    query = """
        SELECT 
            c.id AS id_carga,
            COALESCE(c.numero_tp, c.numero_carga) AS "Nº TP",
            c.numero_set AS "Nº SET",
            c.data_coleta AS "Data Programada",
            c.status AS "Status",
            COALESCE(c.nome_motorista, '-') AS "Motorista",
            COALESCE(c.cpf_motorista, '-') AS "CPF Motorista",
            COALESCE(c.tipo_motorista, '-') AS "Tipo Motorista",
            COALESCE(c.placa_cavalo, '-') AS "Placa Cavalo",
            COALESCE(c.placa_carreta, '-') AS "Placa Carreta",
            c.origens_json,
            c.destinos_json,
            COALESCE(c.receita_frete, 0.0) AS receita_frete,
            COALESCE(c.receita_pedagio, 0.0) AS receita_pedagio,
            COALESCE(c.receita_taxa_descarga, 0.0) AS receita_taxa_descarga,
            COALESCE(c.valor_rpa, 0.0) AS "RPA",
            COALESCE(e.valor_pedagio_pago, 0.0) AS "Pedágio Pago",
            COALESCE(o.custo_fornecedor_descarga, 0.0) AS "Custo Descarga",
            COALESCE(o.fornecedor_descarga, '-') AS "Fornecedor Descarga",
            COALESCE(o.data_agendamento_descarga, '-') AS "Data Agendamento Descarga",
            COALESCE(e.numero_cte, '-') AS "CT-e",
            COALESCE(e.numero_viagem, '-') AS "Viagem",
            COALESCE(e.numero_mdfe, '-') AS "MDF-e",
            COALESCE(e.numero_contrato, '-') AS "Contrato",
            CASE 
                WHEN e.notas_fiscais_expedicao IS NOT NULL AND e.notas_fiscais_expedicao != '' AND e.notas_fiscais_expedicao != '[]' 
                THEN e.notas_fiscais_expedicao
                ELSE c.notas_fiscais_comercial
            END AS "Nota Fiscal",
            COALESCE(c.tipo_pedagio, '-') AS "Tipo de Pedágio",
            COALESCE(c.plataforma, '-') AS "Plataforma",
            COALESCE(c.vinculo, '-') AS "Vínculo",
            COALESCE(a.data_liberacao_saldo, '-') AS "Data Pagamento Saldo"
        FROM cargas c
        LEFT JOIN carga_expedicao e ON c.id = e.carga_id
        LEFT JOIN carga_operacional o ON c.id = o.carga_id
        LEFT JOIN carga_administracao a ON c.id = a.carga_id
        ORDER BY c.id ASC
    """
    with get_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(query)
            dados = cursor.fetchall()
            colunas = [desc.name for desc in cursor.description]
            df = pd.DataFrame(dados, columns=colunas)

    if df.empty:
        st.info("Nenhuma carga cadastrada no momento.")
    else:
        def formatar_real(valor):
            return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        df["Origem"] = df["origens_json"].apply(lambda x: formatar_origens_destinos(x))
        df["Destino"] = df["destinos_json"].apply(lambda x: formatar_origens_destinos(x))
        df["Cliente Origem"] = df["origens_json"].apply(lambda x: " | ".join([i.get("cliente","") for i in processar_json_lista(x) if isinstance(i, dict)]))
        df["Cliente Destino"] = df["destinos_json"].apply(lambda x: " | ".join([i.get("cliente","") for i in processar_json_lista(x) if isinstance(i, dict)]))

        df["Receita Total (R$)"] = df["receita_frete"] + df["receita_pedagio"] + df["receita_taxa_descarga"]
        df["Custo Total (R$)"] = df["RPA"] + df["Pedágio Pago"] + df["Custo Descarga"]
        df["Margem (R$)"] = df["Receita Total (R$)"] - df["Custo Total (R$)"]
        df["Margem (%)"] = df.apply(lambda r: (r["Margem (R$)"] / r["Receita Total (R$)"] * 100) if r["Receita Total (R$)"] > 0 else 0.0, axis=1)

        df["CT-e"] = df["CT-e"].apply(lambda x: limpar_formato_json_lista(x, separador=", "))
        df["Viagem"] = df["Viagem"].apply(lambda x: limpar_formato_json_lista(x, separador=", "))
        df["MDF-e"] = df["MDF-e"].apply(lambda x: limpar_formato_json_lista(x, separador=", "))
        df["Nota Fiscal"] = df["Nota Fiscal"].apply(lambda x: limpar_formato_json_lista(x, separador=" / "))

        cols_final = [
            "Nº TP", "Nº SET", "Data Programada", "Status", "Motorista", 
            "CPF Motorista", "Tipo Motorista", "Placa Cavalo", "Placa Carreta",
            "Origem", "Destino", "Cliente Origem", "Cliente Destino", 
            "Receita Total (R$)", "RPA", "Pedágio Pago", "Custo Descarga", 
            "Fornecedor Descarga", "Data Agendamento Descarga", "CT-e", 
            "Viagem", "MDF-e", "Contrato", "Nota Fiscal", "Tipo de Pedágio", 
            "Plataforma", "Vínculo", "Custo Total (R$)", "Margem (R$)", 
            "Margem (%)", "Data Pagamento Saldo"
        ]

        rec_tot = df["Receita Total (R$)"].sum()
        custo_tot = df["Custo Total (R$)"].sum()
        margem_tot = df["Margem (R$)"].sum()

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Total Cargas", len(df))
        m2.metric("Receita Total", formatar_real(rec_tot))
        m3.metric("Custos Totais", formatar_real(custo_tot))
        m4.metric("Margem Total", formatar_real(margem_tot))
        m5.metric("Margem Média", f"{round((margem_tot / rec_tot * 100) if rec_tot > 0 else 0.0)}%")

        st.markdown("---")

        c_exp1, c_exp2, _ = st.columns([1, 1, 3])
        with c_exp1:
            st.download_button("📗 Exportar Excel", data=gerar_excel_geral(df[cols_final]), file_name="Relatorio_Geral_Transpes.xlsx", use_container_width=True, key="btn_export_excel")
        with c_exp2:
            st.download_button("📄 Exportar PDF", data=gerar_pdf_geral(df), file_name="Relatorio_Geral_Transpes.pdf", use_container_width=True, key="btn_export_pdf")

        df_exib = df.copy()
        for c in ["Receita Total (R$)", "RPA", "Pedágio Pago", "Custo Descarga", "Custo Total (R$)", "Margem (R$)"]:
            df_exib[c] = df_exib[c].apply(formatar_real)
        df_exib["Margem (%)"] = df_exib["Margem (%)"].apply(lambda v: f"{round(v)}%")

        st.dataframe(df_exib[cols_final], use_container_width=True, hide_index=True)

# 2. COMERCIAL
elif menu_selecionado == "Comercial":
    st.title("💼 Comercial - Novo Frete")

    if "com_form_version" not in st.session_state:
        st.session_state["com_form_version"] = 0

    v_com = st.session_state["com_form_version"]
    if st.session_state.get("exibir_modal_comercial", False):
        exibir_popup_comercial(st.session_state.get("carga_comercial_num", ""))

    c_q1, c_q2, c_q3, _ = st.columns([1, 1, 1, 2])
    qtd_origens = c_q1.number_input("Qtd. Origens*", 1, 10, 1, key=f"com_q_orig_{v_com}")
    qtd_destinos = c_q2.number_input("Qtd. Destinos*", 1, 10, 1, key=f"com_q_dest_{v_com}")
    qtd_nfs = c_q3.number_input("Qtd. NFs", 0, 10, 0, key=f"com_q_nfs_{v_com}")

    with st.form(f"form_comercial_{v_com}"):
        numero_set = st.text_input("Número do SET*", placeholder="Ex: SET-2026-001").strip().upper()

        st.subheader("Origens")
        lista_origens = []
        for i in range(int(qtd_origens)):
            co1, co2, co3 = st.columns([2, 2, 1])
            cli = co1.text_input(f"Cliente Origem {i+1}*", key=f"com_cli_orig_{i}_{v_com}").upper()
            cid = co2.text_input(f"Cidade {i+1}*", key=f"com_cid_orig_{i}_{v_com}").upper()
            uf = co3.text_input(f"UF {i+1}*", key=f"com_uf_orig_{i}_{v_com}").upper()
            lista_origens.append({"cliente": cli, "cidade": cid, "estado": uf})

        st.subheader("Destinos")
        lista_destinos = []
        for j in range(int(qtd_destinos)):
            cd1, cd2, cd3 = st.columns([2, 2, 1])
            cli = cd1.text_input(f"Cliente Destino {j+1}*", key=f"com_cli_dest_{j}_{v_com}").upper()
            cid = cd2.text_input(f"Cidade {j+1}*", key=f"com_cid_dest_{j}_{v_com}").upper()
            uf = cd3.text_input(f"UF {j+1}*", key=f"com_uf_dest_{j}_{v_com}").upper()
            lista_destinos.append({"cliente": cli, "cidade": cid, "estado": uf})

        lista_nfs_com = []
        if qtd_nfs > 0:
            st.subheader("Notas Fiscais")
            cols_nf = st.columns(min(int(qtd_nfs), 4))
            for k in range(int(qtd_nfs)):
                val_nf = cols_nf[k % 4].text_input(f"NF {k+1}", key=f"com_nf_{k}_{v_com}").upper()
                if val_nf.strip():
                    lista_nfs_com.append(val_nf.strip())

        st.subheader("Valores do Cliente")
        v1, v2, v3 = st.columns(3)
        receita_frete = v1.number_input("Frete (R$)*", min_value=0.0, value=0.0, step=100.0)
        receita_pedagio = v2.number_input("Pedágio (R$)", min_value=0.0, value=0.0, step=50.0)
        receita_taxa_descarga = v3.number_input("Taxa Descarga (R$)", min_value=0.0, value=0.0, step=50.0)

        salvar_com = st.form_submit_button("💾 Salvar Comercial", type="primary", use_container_width=True)

    if salvar_com:
        valid_orig = all(o["cliente"] and o["cidade"] and o["estado"] for o in lista_origens)
        valid_dest = all(d["cliente"] and d["cidade"] and d["estado"] for d in lista_destinos)

        if not numero_set:
            st.error("O campo SET é obrigatório.")
        elif not valid_orig or not valid_dest:
            st.error("Preencha todos os campos de origens e destinos.")
        elif receita_frete <= 0:
            st.error("O Valor do Frete deve ser maior que zero.")
        else:
            num_carga_novo = gerar_numero_carga_novo()
            try:
                with get_connection() as conn:
                    with conn.cursor() as cursor:
                        cursor.execute("""
                            INSERT INTO cargas (
                                numero_carga, numero_set, origens_json, destinos_json, notas_fiscais_comercial,
                                receita_frete, receita_pedagio, receita_taxa_descarga,
                                status, nome_motorista
                            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, 'PENDENTE_PROGRAMACAO', '-')
                        """, (
                           num_carga_novo, numero_set, json.dumps(lista_origens), json.dumps(lista_destinos),
                           json.dumps(lista_nfs_com) if lista_nfs_com else "",
                           receita_frete, receita_pedagio, receita_taxa_descarga
                       ))
                        conn.commit()

                st.session_state["exibir_modal_comercial"] = True
                st.session_state["carga_comercial_num"] = num_carga_novo
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao salvar: {e}")

# 3. PROGRAMAÇÃO
elif menu_selecionado == "Programação":
    st.title("📌 Programação de Cargas")

    if "prog_form_version" not in st.session_state:
        st.session_state["prog_form_version"] = 0

    v_prog = st.session_state["prog_form_version"]
    if st.session_state.get("exibir_modal_programacao", False):
        exibir_popup_programacao(st.session_state.get("tp_programado_num", ""))

    query_prog = "SELECT id, numero_carga, numero_set, origens_json, destinos_json, notas_fiscais_comercial FROM cargas WHERE status = 'PENDENTE_PROGRAMACAO'"
    with get_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(query_prog)
            dados = cursor.fetchall()
            cargas_prog_df = pd.DataFrame(dados, columns=[desc.name for desc in cursor.description])

    if cargas_prog_df.empty:
        st.info("Nenhuma carga pendente de programação.")
    else:
        opcoes_cargas = {f"SET: {r['numero_set']} - ORIGEM: {formatar_origens_destinos(r['origens_json'], True)} - DESTINO: {formatar_origens_destinos(r['destinos_json'], True)}": r['id'] for _, r in cargas_prog_df.iterrows()}
        carga_id = opcoes_cargas[st.selectbox("Selecione a Carga*", list(opcoes_cargas.keys()), key=f"prog_sel_{v_prog}")]

        with st.form(f"form_programacao_{v_prog}"):
            c1, c2, c3, c4 = st.columns(4)
            numero_tp = c1.text_input("Número TP*", placeholder="TP-1234").strip().upper()
            nome_motorista = c2.text_input("Motorista*").upper()
            cpf_motorista = c3.text_input("CPF Motorista")
            telefone_motorista = c4.text_input("Telefone Motorista")

            c5, c6, c7, c8 = st.columns(4)
            tipo_motorista = c5.selectbox("Tipo Motorista*", ["TERCEIRO", "FROTA", "AGREGADO"])
            tipo_veiculo = c6.selectbox("Tipo Veículo*", ["TRUCK", "CARRETA 13m", "CARRETA 15m", "BITREM", "RODOTREM"])
            eixos = c7.number_input("Eixos*", 2, 9, 6)
            valor_rpa = c8.number_input("RPA (R$)*", min_value=0.0, value=0.0, step=100.0)

            c9, c10, c11, c12 = st.columns(4)
            placa_cavalo = c9.text_input("Placa Cavalo*").upper()
            placa_carreta = c10.text_input("Placa Carreta*").upper()
            data_coleta = c11.date_input("Data Coleta*", datetime.date.today(), format="DD/MM/YYYY")
            previsao_descarga = c12.date_input("Previsão Entrega*", datetime.date.today() + datetime.timedelta(days=2), format="DD/MM/YYYY")

            c13, c14, c15 = st.columns(3)
            tipo_pedagio = c13.text_input("Tipo Pedágio").strip().upper()
            plataforma = c14.selectbox("Plataforma*", ["APP", "Card"])
            vinculo = c15.selectbox("Vínculo*", ["CPF", "CNPJ"])

            obs_prog = st.text_area("Observação").upper()

            salvar_prog = st.form_submit_button("💾 Salvar Programação", type="primary", use_container_width=True)

        if salvar_prog:
            if not numero_tp or not nome_motorista or not placa_cavalo or not placa_carreta:
                st.error("Preencha os campos obrigatórios (*).")
            else:
                try:
                    with get_connection() as conn:
                        with conn.cursor() as cursor:
                            cursor.execute("""
                                UPDATE cargas SET
                                    numero_tp = %s, nome_motorista = %s, cpf_motorista = %s,
                                    telefone_motorista = %s, tipo_motorista = %s, tipo_veiculo = %s,
                                    quantidade_eixos = %s, placa_cavalo = %s, placa_carreta = %s,
                                    valor_rpa = %s, data_coleta = %s, previsao_descarga = %s,
                                    tipo_pedagio = %s, plataforma = %s, vinculo = %s, 
                                    observacoes_programacao = %s, status = 'PROGRAMADA'
                                WHERE id = %s
                            """, (
                                numero_tp, nome_motorista, formatar_cpf(cpf_motorista), formatar_telefone(telefone_motorista), tipo_motorista,
                                tipo_veiculo, eixos, placa_cavalo, placa_carreta, valor_rpa,
                                data_coleta.strftime("%d/%m/%Y"), previsao_descarga.strftime("%d/%m/%Y"),
                                tipo_pedagio, plataforma, vinculo, obs_prog, carga_id
                            ))
                            conn.commit()

                    st.session_state["exibir_modal_programacao"] = True
                    st.session_state["tp_programado_num"] = numero_tp
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao salvar: {e}")

# 4. EXPEDIÇÃO
elif menu_selecionado == "Expedição":
    st.title("📦 Expedição")

    if "exp_form_version" not in st.session_state:
        st.session_state["exp_form_version"] = 0

    v_exp = st.session_state["exp_form_version"]
    if st.session_state.get("exibir_modal_expedicao", False):
        exibir_popup_expedicao(st.session_state.get("exp_tp_num", ""))

    query_exp = "SELECT id, numero_set, numero_tp, nome_motorista, origens_json, destinos_json, notas_fiscais_comercial, valor_rpa FROM cargas WHERE status = 'PROGRAMADA'"
    with get_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(query_exp)
            dados = cursor.fetchall()
            cargas_exp_df = pd.DataFrame(dados, columns=[desc.name for desc in cursor.description])

    if cargas_exp_df.empty:
        st.info("Nenhuma carga aguardando expedição.")
    else:
        opcoes_cargas = {f"SET: {r['numero_set']} - TP: {r['numero_tp']} - MOTORISTA: {r['nome_motorista']}": r['id'] for _, r in cargas_exp_df.iterrows()}
        carga_id = opcoes_cargas[st.selectbox("Selecione a Carga Programada*", list(opcoes_cargas.keys()), key=f"exp_sel_{v_exp}")]
        
        row_sel = cargas_exp_df[cargas_exp_df['id'] == carga_id].iloc[0]
        v_rpa = row_sel['valor_rpa'] or 0.0

        e_c1, e_c2, e_c3, e_c4 = st.columns(4)
        qtd_ctes = e_c1.number_input("Qtd. CT-es*", 1, 10, 1, key=f"exp_q_cte_{v_exp}")
        qtd_viagens = e_c2.number_input("Qtd. Viagens", 1, 10, 1, key=f"exp_q_v_{v_exp}")
        qtd_mdfes = e_c3.number_input("Qtd. MDF-es*", 1, 10, 1, key=f"exp_q_mdfe_{v_exp}")
        qtd_nfs_exp = e_c4.number_input("Qtd. NFs livres", 0, 10, 0, key=f"exp_q_nf_{v_exp}")

        with st.form(f"form_expedicao_{v_exp}"):
            st.text_input("Adiantamento (70% RPA)", value=f"R$ {v_rpa * 0.70:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), disabled=True)

            lista_ctes = []
            cols_cte = st.columns(min(int(qtd_ctes), 4))
            for i in range(int(qtd_ctes)):
                v_c = cols_cte[i % 4].text_input(f"CT-e {i+1}*", key=f"exp_cte_{i}_{v_exp}").upper()
                if v_c.strip(): lista_ctes.append(v_c.strip())

            lista_viagens = []
            cols_v = st.columns(min(int(qtd_viagens), 4))
            for j in range(int(qtd_viagens)):
                v_v = cols_v[j % 4].text_input(f"Viagem {j+1}", key=f"exp_viag_{j}_{v_exp}").upper()
                if v_v.strip(): lista_viagens.append(v_v.strip())

            f1, f2, f3 = st.columns(3)
            lista_mdfes = []
            cols_mdfe = f1.columns(min(int(qtd_mdfes), 2))
            for k in range(int(qtd_mdfes)):
                v_m = cols_mdfe[k % 2].text_input(f"MDF-e {k+1}*", key=f"exp_mdfe_{k}_{v_exp}").upper()
                if v_m.strip(): lista_mdfes.append(v_m.strip())

            numero_contrato = f2.text_input("Número do Contrato*", placeholder="CTR-12345").strip().upper()
            data_saida_filial = f3.date_input("Data Saída Filial*", datetime.date.today(), format="DD/MM/YYYY")

            valor_pedagio_pago = st.number_input("Pedágio Pago Motorista (R$)", min_value=0.0, value=0.0, step=50.0)

            lista_nfs_exp = []
            if qtd_nfs_exp > 0:
                cols_nf = st.columns(min(int(qtd_nfs_exp), 4))
                for n in range(int(qtd_nfs_exp)):
                    v_nf = cols_nf[n % 4].text_input(f"NF {n+1}", key=f"exp_nf_{n}_{v_exp}").upper()
                    if v_nf.strip(): lista_nfs_exp.append(v_nf.strip())

            obs_exp = st.text_area("Observação").upper()

            salvar_exp = st.form_submit_button("🚚 Confirmar Saída / Expedição", type="primary", use_container_width=True)

        if salvar_exp:
            if len(lista_ctes) < int(qtd_ctes) or len(lista_mdfes) < int(qtd_mdfes) or not numero_contrato:
                st.error("Preencha todos os campos obrigatórios (*).")
            else:
                try:
                    with get_connection() as conn:
                        with conn.cursor() as cursor:
                            cursor.execute("""
                                INSERT INTO carga_expedicao (
                                    carga_id, numero_cte, numero_viagem, numero_mdfe, numero_contrato,
                                    valor_pedagio_pago, notas_fiscais_expedicao, data_saida_filial, observacoes_expedicao
                                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                            """, (
                                carga_id, json.dumps(lista_ctes), json.dumps(lista_viagens) if lista_viagens else "",
                                json.dumps(lista_mdfes), numero_contrato, valor_pedagio_pago,
                                json.dumps(lista_nfs_exp) if lista_nfs_exp else "", data_saida_filial.strftime("%d/%m/%Y"), obs_exp
                            ))
                            cursor.execute("UPDATE cargas SET status = 'EM TRÂNSITO' WHERE id = %s", (carga_id,))
                            conn.commit()

                    st.session_state["exibir_modal_expedicao"] = True
                    st.session_state["exp_tp_num"] = row_sel['numero_tp']
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao salvar: {e}")

# 5. OPERACIONAL
elif menu_selecionado == "Operacional":
    st.title("⚙️ Operacional - Descarga")

    if "op_form_version" not in st.session_state:
        st.session_state["op_form_version"] = 0

    v_op = st.session_state["op_form_version"]
    if st.session_state.get("exibir_modal_operacional", False):
        exibir_popup_operacional(st.session_state.get("op_tp_num", ""))

    query_op = """
        SELECT c.id, c.numero_tp, c.nome_motorista, c.origens_json, c.destinos_json, c.previsao_descarga, e.data_saida_filial
        FROM cargas c
        INNER JOIN carga_expedicao e ON c.id = e.carga_id
        WHERE c.status = 'EM TRÂNSITO'
    """
    with get_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(query_op)
            dados = cursor.fetchall()
            cargas_op_df = pd.DataFrame(dados, columns=[desc.name for desc in cursor.description])

    if cargas_op_df.empty:
        st.info("Nenhuma carga em trânsito no momento.")
    else:
        opcoes_cargas = {f"TP: {r['numero_tp']} - MOTORISTA: {r['nome_motorista']}": r['id'] for _, r in cargas_op_df.iterrows()}
        carga_id = opcoes_cargas[st.selectbox("Selecione a Carga em Trânsito*", list(opcoes_cargas.keys()), key=f"op_sel_{v_op}")]
        row_sel = cargas_op_df[cargas_op_df['id'] == carga_id].iloc[0]

        with get_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute("SELECT razao_social, nome_fantasia FROM fornecedores ORDER BY razao_social")
                forn_df = pd.DataFrame(cursor.fetchall(), columns=[desc.name for desc in cursor.description])
        
        lista_fornecedores = [f"{r['razao_social']} ({r['nome_fantasia']})" if r['nome_fantasia'] else r['razao_social'] for _, r in forn_df.iterrows()]

        inf1, inf2, inf3 = st.columns(3)
        inf1.text_input("Origem/Destino", value=f"{formatar_origens_destinos(row_sel['origens_json'], True)} -> {formatar_origens_destinos(row_sel['destinos_json'], True)}", disabled=True)
        inf2.text_input("Saída Filial", value=row_sel['data_saida_filial'], disabled=True)
        inf3.text_input("Prev. Descarga", value=row_sel['previsao_descarga'], disabled=True)

        with st.form(f"form_operacional_{v_op}"):
            f1, f2 = st.columns(2)
            fornecedor_sel = f1.selectbox("Fornecedor Descarga*", lista_fornecedores) if lista_fornecedores else f1.text_input("Fornecedor Descarga*").upper()
            custo_descarga = f2.number_input("Custo Descarga (R$)*", min_value=0.0, value=0.0, step=50.0)

            p1, p2, p3 = st.columns(3)
            data_agendamento = p1.date_input("Data Agendamento*", datetime.date.today(), format="DD/MM/YYYY")
            forma_pagamento = p2.selectbox("Forma Pagamento*", ["PIX", "BOLETO", "DEPÓSITO BANCÁRIO"])
            prazo_dias = p3.number_input("Prazo (dias)*", min_value=0, value=30, step=5)

            salvar_op = st.form_submit_button("✅ Salvar Operacional", type="primary", use_container_width=True)

        if salvar_op:
            if not fornecedor_sel:
                st.error("Selecione um fornecedor.")
            else:
                try:
                    with get_connection() as conn:
                        with conn.cursor() as cursor:
                            cursor.execute("""
                                INSERT INTO carga_operacional (
                                    carga_id, fornecedor_descarga, custo_fornecedor_descarga,
                                    data_agendamento_descarga, forma_pagamento, prazo_pagamento_dias
                                ) VALUES (%s, %s, %s, %s, %s, %s)
                            """, (
                                carga_id, fornecedor_sel, custo_descarga,
                                data_agendamento.strftime("%d/%m/%Y"), forma_pagamento, prazo_dias
                            ))
                            cursor.execute("UPDATE cargas SET status = 'ENTREGUE' WHERE id = %s", (carga_id,))
                            conn.commit()

                    st.session_state["exibir_modal_operacional"] = True
                    st.session_state["op_tp_num"] = row_sel['numero_tp']
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao salvar: {e}")

# 6. ADMINISTRAÇÃO
elif menu_selecionado == "Administração":
    st.title("💼 Administração - Acerto Financeiro")

    if "adm_form_version" not in st.session_state:
        st.session_state["adm_form_version"] = 0

    v_adm = st.session_state["adm_form_version"]
    if st.session_state.get("exibir_modal_administracao", False):
        exibir_popup_administracao(st.session_state.get("adm_tp_num", ""))

    query_adm = "SELECT c.id, c.numero_tp, c.nome_motorista, COALESCE(c.valor_rpa, 0.0) AS valor_rpa, e.numero_cte FROM cargas c LEFT JOIN carga_expedicao e ON c.id = e.carga_id WHERE c.status = 'ENTREGUE'"
    with get_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(query_adm)
            dados = cursor.fetchall()
            cargas_adm_df = pd.DataFrame(dados, columns=[desc.name for desc in cursor.description])

    if cargas_adm_df.empty:
        st.info("Nenhuma carga entregue aguardando acerto.")
    else:
        opcoes_cargas = {f"TP: {r['numero_tp']} - MOTORISTA: {r['nome_motorista']}": r['id'] for _, r in cargas_adm_df.iterrows()}
        carga_id = opcoes_cargas[st.selectbox("Selecione a Carga Entregue*", list(opcoes_cargas.keys()), key=f"adm_sel_{v_adm}")]
        row_sel = cargas_adm_df[cargas_adm_df['id'] == carga_id].iloc[0]

        v_rpa = row_sel['valor_rpa']
        adiantamento_calc = v_rpa * 0.70
        saldo_calc = v_rpa * 0.30

        with st.form(f"form_administracao_{v_adm}"):
            a1, a2, a3, a4 = st.columns(4)
            a1.text_input("Adiantamento (70%)", value=f"R$ {adiantamento_calc:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), disabled=True)
            a2.text_input("Saldo (30%)", value=f"R$ {saldo_calc:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), disabled=True)
            status_pag = a3.selectbox("Status Saldo", ["PENDENTE", "PAGO", "CANCELADO"])
            data_lib = a4.date_input("Data Liberação", datetime.date.today(), format="DD/MM/YYYY")

            comprovante = st.checkbox("Comprovante de Entregue Recebido")

            salvar_adm = st.form_submit_button("💰 Finalizar Acerto Financeiro", type="primary", use_container_width=True)

        if salvar_adm:
            try:
                with get_connection() as conn:
                    with conn.cursor() as cursor:
                        cursor.execute("""
                            INSERT INTO carga_administracao (
                                carga_id, valor_adiantamento, valor_saldo, comprovante_entregue,
                                data_liberacao_saldo, status_pagamento_saldo
                            ) VALUES (%s, %s, %s, %s, %s, %s)
                        """, (carga_id, adiantamento_calc, saldo_calc, 1 if comprovante else 0, data_lib.strftime("%d/%m/%Y"), status_pag))
                        cursor.execute("UPDATE cargas SET status = 'FINALIZADA' WHERE id = %s", (carga_id,))
                        conn.commit()

                st.session_state["exibir_modal_administracao"] = True
                st.session_state["adm_tp_num"] = row_sel['numero_tp']
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao salvar: {e}")

# 7. FORNECEDORES
elif menu_selecionado == "Fornecedores":
    st.title("🚚 Gestão de Fornecedores")

    col_btn1, _ = st.columns([2, 3])
    with col_btn1:
        if st.button("➕ Cadastrar Novo Fornecedor", type="primary", use_container_width=True):
            modal_cadastrar_fornecedor()

    with st.expander("📥 Importação em Lote (Excel / CSV)", expanded=False):
        df_modelo = pd.DataFrame({"nr_contrato": ["12345"], "razao_social": ["EXEMPLO LTDA"], "nome_fantasia": ["EXEMPLO"], "local_atendimento": ["SÃO PAULO"], "uf": ["SP"], "cpf_cnpj": ["00.000.000/0001-00"], "contato": ["(11) 99999-9999"]})
        buffer_modelo = io.BytesIO()
        with pd.ExcelWriter(buffer_modelo, engine='openpyxl') as writer:
            df_modelo.to_excel(writer, index=False, sheet_name='Fornecedores')
        
        st.download_button("📄 Baixar Modelo (.xlsx)", data=buffer_modelo.getvalue(), file_name="Modelo_Fornecedores.xlsx")

        arquivo_carregado = st.file_uploader("Arquivo Excel ou CSV", type=["xlsx", "csv"])
        if arquivo_carregado is not None:
            try:
                df_import = pd.read_csv(arquivo_carregado, dtype=str) if arquivo_carregado.name.endswith('.csv') else pd.read_excel(arquivo_carregado, dtype=str)
                df_import = df_import.fillna("")
                st.dataframe(df_import.head(5), use_container_width=True)

                if st.button("🚀 Confirmar Importação"):
                    sucessos = 0
                    with get_connection() as conn:
                        with conn.cursor() as cursor:
                            for _, row in df_import.iterrows():
                                if str(row.get("razao_social", "")).strip():
                                    cursor.execute("INSERT INTO fornecedores (nr_contrato, razao_social, nome_fantasia, local_atendimento, uf, cpf_cnpj, contato) VALUES (%s, %s, %s, %s, %s, %s, %s)",
                                                   (row.get("nr_contrato", ""), row.get("razao_social", ""), row.get("nome_fantasia", ""), row.get("local_atendimento", ""), row.get("uf", ""), row.get("cpf_cnpj", ""), row.get("contato", "")))
                                    sucessos += 1
                            conn.commit()
                    st.success(f"✅ {sucessos} fornecedores importados!")
                    st.rerun()
            except Exception as e:
                st.error(f"Erro na importação: {e}")

    query_forn = 'SELECT nr_contrato AS "Nr. Contrato", razao_social AS "Razão Social", nome_fantasia AS "Nome Fantasia", local_atendimento AS "Local", uf AS "UF", cpf_cnpj AS "CPF/CNPJ", contato AS "Contato" FROM fornecedores ORDER BY id DESC'
    with get_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(query_forn)
            dados = cursor.fetchall()
            df_fornecedores = pd.DataFrame(dados, columns=[desc.name for desc in cursor.description])

    st.dataframe(df_fornecedores, use_container_width=True, hide_index=True)

# 8. EXCLUIR CARGAS
elif menu_selecionado == "Excluir Cargas":
    st.title("🗑️ Excluir Cargas")
    
    with get_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute("SELECT id, numero_carga, COALESCE(numero_tp, '-') AS tp, COALESCE(nome_motorista, '-') AS mot, status FROM cargas")
            cargas_df = pd.DataFrame(cursor.fetchall(), columns=[desc.name for desc in cursor.description])
    
    if cargas_df.empty:
        st.info("Nenhuma carga cadastrada.")
    else:
        opcoes_cargas = {f"Carga: {r['numero_carga']} | TP: {r['tp']} | Motorista: {r['mot']} | Status: {r['status']}": r['id'] for _, r in cargas_df.iterrows()}
        carga_id = opcoes_cargas[st.selectbox("Selecione a Carga*", list(opcoes_cargas.keys()))]

        if st.button("❌ Excluir Definitivamente", type="primary"):
            with get_connection() as conn:
                with conn.cursor() as cursor:
                    cursor.execute("DELETE FROM carga_administracao WHERE carga_id = %s", (carga_id,))
                    cursor.execute("DELETE FROM carga_operacional WHERE carga_id = %s", (carga_id,))
                    cursor.execute("DELETE FROM carga_expedicao WHERE carga_id = %s", (carga_id,))
                    cursor.execute("DELETE FROM cargas WHERE id = %s", (carga_id,))
                    conn.commit()
            st.success("Carga excluída!")
            st.rerun()

# 9. USUÁRIOS
elif menu_selecionado == "Usuários":
    st.title("👥 Gestão de Usuários")
    
    with get_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute("SELECT id, usuario, nome, perfil FROM usuarios")
            users_df = pd.DataFrame(cursor.fetchall(), columns=[desc.name for desc in cursor.description])
            
    st.dataframe(users_df, use_container_width=True, hide_index=True)

    st.subheader("➕ Novo Usuário")
    with st.form("form_novo_usuario"):
        u1, u2, u3, u4 = st.columns(4)
        novo_usr = u1.text_input("Usuário*").strip().lower()
        novo_nome = u2.text_input("Nome*").strip().upper()
        nova_senha = u3.text_input("Senha*", type="password")
        novo_perfil = u4.selectbox("Perfil*", ["ADMIN", "COMERCIAL", "PROGRAMACAO", "EXPEDICAO", "OPERACIONAL", "ADMINISTRATIVO"])

        salvar_usr = st.form_submit_button("💾 Salvar Usuário", type="primary", use_container_width=True)

        if salvar_usr:
            if not (novo_usr and novo_nome and nova_senha and novo_perfil):
                st.error("Preencha todos os campos.")
            else:
                try:
                    with get_connection() as conn:
                        with conn.cursor() as cursor:
                            cursor.execute("INSERT INTO usuarios (usuario, senha, nome, perfil) VALUES (%s, %s, %s, %s)",
                                           (novo_usr, hash_senha(nova_senha), novo_nome, novo_perfil))
                            conn.commit()
                    st.success("Usuário criado com sucesso!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao salvar: {e}")
