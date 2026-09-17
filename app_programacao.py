import io
import base64
import datetime
import hashlib
import json
import psycopg2
import pandas as pd
import streamlit as st
import re
import openpyxl
from streamlit_option_menu import option_menu

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
# CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(
    page_title="Transpes - Sistema de Gestão",
    page_icon="logo_transpes.png",
    layout="wide"
)

# ==========================================
# CONEXÃO E BANCO DE DADOS (SUPABASE / POSTGRESQL)
# ==========================================

def get_connection():
    return psycopg2.connect(st.secrets["DB_URL"])

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
            # 1. Tabela de Usuários
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

            # 2. Tabela Principal de Cargas
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS cargas (
                    id SERIAL PRIMARY KEY,
                    numero_carga TEXT UNIQUE NOT NULL,
                    status TEXT DEFAULT 'PENDENTE_PROGRAMACAO',
                    data_criacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

            # 3. Adiciona as colunas novas caso ainda não existam no banco Supabase/Postgres
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
                ("observacoes_programacao", "TEXT")
            ]

            for col_nome, col_tipo in novas_colunas_cargas:
                cursor.execute(f"ALTER TABLE cargas ADD COLUMN IF NOT EXISTS {col_nome} {col_tipo};")

            # 4. Tabela Carga Expedição
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

            # 5. Tabela Carga Operacional
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

            # 6. Tabela Carga Administração
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

            # 7. Tabela Fornecedores
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
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

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

def gerar_excel_geral(df):
    output = io.BytesIO()
    
    # 1. Seleciona estritamente as colunas do layout da imagem
    cols_excel = [
        "Nº TP", "Nº SET", "Data Programada", "Status", "Motorista", 
        "Origem", "Destino", "Receita Total (R$)", "Custo Total (R$)", 
        "Margem (R$)", "Margem (%)"
    ]
    
    df_excel = df.copy()
    
    # Garante que as colunas existam no DataFrame
    cols_presentes = [c for c in cols_excel if c in df_excel.columns]
    df_excel = df_excel[cols_presentes]

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Escreve a tabela a partir da linha 4 para dar espaço ao título
        df_excel.to_excel(writer, index=False, sheet_name="Relatório Geral", startrow=3)
        
        workbook = writer.book
        worksheet = writer.sheets["Relatório Geral"]
        
        # Desativa as linhas de grade padrão para dar o visual limpo do relatório
        worksheet.views.sheetView[0].showGridLines = True

        # --- ESTILOS ---
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        
        font_titulo = Font(name="Arial", size=14, bold=True, color="0F2A4A")
        font_cabecalho = Font(name="Arial", size=9, bold=True, color="FFFFFF")
        font_dados = Font(name="Arial", size=8)
        
        align_titulo = Alignment(horizontal="center", vertical="center")
        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        align_right = Alignment(horizontal="right", vertical="center")

        fill_cabecalho = PatternFill(start_color="0F2A4A", end_color="0F2A4A", fill_type="solid")
        fill_zebra = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
        fill_branca = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

        borda_fina = Side(border_style="thin", color="CCCCCC")
        borda_caixa = Border(left=borda_fina, right=borda_fina, top=borda_fina, bottom=borda_fina)

        # 2. Insere e formata o Título
        worksheet.merge_cells(start_row=1, start_column=1, end_row=2, end_column=len(cols_presentes))
        cell_titulo = worksheet.cell(row=1, column=1)
        cell_titulo.value = "Relatório Geral de Cargas - Transpes"
        cell_titulo.font = font_titulo
        cell_titulo.alignment = align_titulo

        # 3. Formata o Cabeçalho (Linha 4 do Excel)
        for col_num in range(1, len(cols_presentes) + 1):
            cell = worksheet.cell(row=4, column=col_num)
            cell.font = font_cabecalho
            cell.fill = fill_cabecalho
            cell.alignment = align_center
            cell.border = borda_caixa

        # 4. Formata as Linhas de Dados
        num_linhas = len(df_excel)
        for row_idx in range(5, 5 + num_linhas):
            fill_atual = fill_zebra if row_idx % 2 == 0 else fill_branca
            
            for col_idx, col_nome in enumerate(cols_presentes, start=1):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                cell.font = font_dados
                cell.fill = fill_atual
                cell.border = borda_caixa
                
                # Alinhamento e formatação por tipo de dado
                if col_nome in ["Nº TP", "Nº SET", "Data Programada", "Status"]:
                    cell.alignment = align_center
                elif col_nome in ["Motorista", "Origem", "Destino"]:
                    cell.alignment = align_left
                elif col_nome in ["Receita Total (R$)", "Custo Total (R$)", "Margem (R$)"]:
                    cell.alignment = align_right
                    cell.number_format = 'R$ #,##0.00'
                elif col_nome == "Margem (%)":
                    cell.alignment = align_right
                    cell.number_format = '0%'

        # 5. Ajuste de Largura das Colunas
        larguras_fixas = {
            "Nº TP": 12, "Nº SET": 12, "Data Programada": 15, "Status": 14,
            "Motorista": 30, "Origem": 45, "Destino": 45,
            "Receita Total (R$)": 18, "Custo Total (R$)": 18, "Margem (R$)": 18, "Margem (%)": 12
        }

        for idx, col in enumerate(cols_presentes, start=1):
            col_letter = openpyxl.utils.get_column_letter(idx)
            worksheet.column_dimensions[col_letter].width = larguras_fixas.get(col, 15)

    return output.getvalue()


from reportlab.lib.pagesizes import letter, landscape, A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

def gerar_pdf_geral(df):
    buffer = io.BytesIO()
    
    # 1. Configura a folha explicitamente na horizontal (A4 Paisagem) com margens finas
    doc = SimpleDocTemplate(
        buffer, 
        pagesize=landscape(A4), 
        rightMargin=15, 
        leftMargin=15, 
        topMargin=20, 
        bottomMargin=20
    )
    elements = []

    styles = getSampleStyleSheet()
    
    style_title = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=14,
        alignment=1,
        textColor=colors.HexColor("#0f2a4a"),
        spaceAfter=10
    )
    
    style_cell = ParagraphStyle(
        'Cell',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=6,
        leading=7,
        alignment=0
    )
    
    style_header = ParagraphStyle(
        'Header',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=6.5,
        leading=8,
        textColor=colors.white,
        alignment=1
    )

    elements.append(Paragraph("Relatório Geral de Cargas - Transpes", style_title))

    # 2. Seleção estrita das colunas essenciais
    cols_pdf = [
        "Nº TP", "Nº SET", "Data Programada", "Status", "Motorista", 
        "Origem", "Destino", "Receita Total (R$)", "Custo Total (R$)", 
        "Margem (R$)", "Margem (%)"
    ]
    
    df_pdf = df.copy()
    
    # Formatação visual dos números no relatório
    for c in ["Receita Total (R$)", "Custo Total (R$)", "Margem (R$)"]:
        if c in df_pdf.columns:
            df_pdf[c] = df_pdf[c].apply(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if isinstance(x, (int, float)) else str(x))
    
    if "Margem (%)" in df_pdf.columns:
        df_pdf["Margem (%)"] = df_pdf["Margem (%)"].apply(lambda x: f"{x:.1f}%" if isinstance(x, (int, float)) else str(x))

    table_data = []
    
    # Monta os cabeçalhos
    headers = [Paragraph(col, style_header) for col in cols_pdf]
    table_data.append(headers)

    # Monta as linhas de dados
    for _, row in df_pdf.iterrows():
        linha = []
        for col in cols_pdf:
            val = str(row[col]) if pd.notnull(row[col]) and str(row[col]) != "" else "-"
            linha.append(Paragraph(val, style_cell))
        table_data.append(linha)

    # 3. Distribuição das larguras das colunas em pontos para fechar 810pt (largura exata da folha A4 em modo paisagem)
    col_widths = [45, 55, 55, 60, 95, 150, 150, 65, 65, 65, 45]

    pdf_table = Table(table_data, colWidths=col_widths, repeatRows=1)
    pdf_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0f2a4a")),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#cccccc")),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
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
    st.write("Os dados foram salvos e liberados para a equipe de Programação.")
    if st.button("OK / Próximo Registro", type="primary", use_container_width=True):
        st.session_state["exibir_modal_comercial"] = False
        st.session_state["com_form_version"] = st.session_state.get("com_form_version", 0) + 1
        st.rerun()

@st.dialog("📌 Programação Concluída!")
def exibir_popup_programacao(num_tp):
    st.success(f"A carga TP **{num_tp}** foi programada com sucesso!")
    st.write("Status atualizado para **PROGRAMADA**. Enviado para a Expedição.")
    if st.button("OK / Próxima Programação", type="primary", use_container_width=True):
        st.session_state["exibir_modal_programacao"] = False
        st.session_state["prog_form_version"] = st.session_state.get("prog_form_version", 0) + 1
        st.rerun()

@st.dialog("📦 Carga Expedida!")
def exibir_popup_expedicao(num_tp):
    st.success(f"A carga TP **{num_tp}** foi expedida com sucesso e está **EM TRÂNSITO**!")
    if st.button("OK / Próxima Carga", type="primary", use_container_width=True):
        st.session_state["exibir_modal_expedicao"] = False
        st.session_state["exp_form_version"] = st.session_state.get("exp_form_version", 0) + 1
        st.rerun()

@st.dialog("⚙️ Agendamento de Descarga Confirmado!")
def exibir_popup_operacional(num_tp):
    st.success(f"A carga TP **{num_tp}** teve os dados operacionais salvos e status alterado para **ENTREGUE**!")
    if st.button("OK / Próxima Carga", type="primary", use_container_width=True):
        st.session_state["exibir_modal_operacional"] = False
        st.session_state["op_form_version"] = st.session_state.get("op_form_version", 0) + 1
        st.rerun()

@st.dialog("💼 Processo Finalizado!")
def exibir_popup_administracao(num_tp):
    st.success(f"A carga TP **{num_tp}** teve o acerto concluído e o status alterado para **FINALIZADA**!")
    if st.button("OK / Concluir", type="primary", use_container_width=True):
        st.session_state["exibir_modal_administracao"] = False
        st.session_state["adm_form_version"] = st.session_state.get("adm_form_version", 0) + 1
        st.rerun()

@st.dialog("➕ Cadastrar Novo Fornecedor")
def modal_cadastrar_fornecedor():
    with st.form("form_modal_fornecedor"):
        f1, f2 = st.columns(2)
        nr_contrato = f1.text_input("Nr. Contrato").strip()
        cpf_cnpj = f2.text_input("CPF/CNPJ*").strip()

        razao_social = st.text_input("Razão Social*").strip().upper()
        nome_fantasia = st.text_input("Nome Fantasia").strip().upper()

        l1, l2 = st.columns([3, 1])
        local_atendimento = l1.text_input("Local de Atendimento*").strip().upper()
        uf = l2.text_input("UF*").strip().upper()

        contato = st.text_input("Contato / Telefone").strip()

        btn_salvar_forn = st.form_submit_button("💾 Salvar Fornecedor", type="primary", use_container_width=True)

    if btn_salvar_forn:
        if not razao_social or not local_atendimento or not uf:
            st.error("Preencha os campos obrigatórios (Razão Social, Local de Atendimento e UF).")
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
    try:
        bg_base64 = get_base64_image("fundo_transpes.png")
        bg_style = f"background-image: url('data:image/jpeg;base64,{bg_base64}'); background-size: cover; background-position: center; background-repeat: no-repeat; background-attachment: fixed;"
    except Exception:
        bg_style = "background-color: #0E2F56;"

    try:
        logo_base64 = get_base64_image("logo_transpes.png")
        logo_html = f'<img src="data:image/png;base64,{logo_base64}" style="max-width: 180px; margin-bottom: 10px;">'
    except Exception:
        logo_html = '<div class="login-title" style="color: #0E2F56; font-size: 26px; font-weight: bold;">ERP Transpes</div>'

    st.markdown(
        f"""
        <style>
        [data-testid="stSidebar"], [data-testid="stHeader"] {{ display: none; }}
        .stApp {{ {bg_style} }}
        [data-testid="stMainBlockContainer"] {{ max-width: 420px !important; padding-top: 5rem !important; margin: auto !important; }}
        .login-card {{ background: rgba(255, 255, 255, 0.95); padding: 25px 20px 18px 20px; border-radius: 14px; box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.4); text-align: center; margin-bottom: 20px; }}
        </style>
        """,
        unsafe_allow_html=True
    )

    st.markdown(f'<div class="login-card">{logo_html}<div style="color: #444; font-size: 13px;">Entre com suas credenciais</div></div>', unsafe_allow_html=True)

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
            st.session_state["nome"] = usr[1]
            st.session_state["perfil"] = usr[2]
            st.rerun()
        else:
            st.error("Usuário ou senha incorretos.")

    st.stop()

# ==========================================
# BARRA LATERAL E NAVEGAÇÃO
# ==========================================
st.sidebar.markdown(f"### 👤 {st.session_state['nome']}")
st.sidebar.caption(f"Perfil: **{st.session_state['perfil']}**")

if st.sidebar.button("🚪 Sair", use_container_width=True):
    st.session_state.clear()
    st.rerun()

perfil = st.session_state["perfil"]

opcoes = ["Visão Geral"]
icones = ["bar-chart"]

if perfil == "ADMIN":
    opcoes.extend(["Comercial", "Programação", "Expedição", "Operacional", "Administração", "Fornecedores", "Excluir Cargas", "Usuários"])
    icones.extend(["currency-dollar", "clipboard-plus", "file-earmark-text", "tools", "briefcase", "truck", "trash", "people"])
else:
    if perfil == "COMERCIAL":
        opcoes.append("Comercial")
        icones.append("currency-dollar")
    elif perfil == "PROGRAMACAO":
        opcoes.append("Programação")
        icones.append("clipboard-plus")
    elif perfil == "EXPEDICAO":
        opcoes.append("Expedição")
        icones.append("file-earmark-text")
    elif perfil == "OPERACIONAL":
        opcoes.extend(["Operacional", "Fornecedores"])
        icones.extend(["tools", "truck"])
    elif perfil == "ADMINISTRATIVO":
        opcoes.extend(["Administração", "Fornecedores"])
        icones.extend(["briefcase", "truck"])

with st.sidebar:
    menu_selecionado = option_menu(
        "Menu Principal",
        opcoes,
        icons=icones,
        menu_icon="cast",
        default_index=0
    )

# ==========================================
# PÁGINAS DO SISTEMA
# ==========================================
# FUNÇÃO AUXILIAR DE LIMPEZA (coloque no topo do arquivo ou antes da Visão Geral)
# ==========================================
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


# ==========================================
# 1. VISÃO GERAL
# ==========================================
if menu_selecionado == "Visão Geral":
    st.title("📊 Visão Geral e Relatórios")

    with get_connection() as conn:
        with conn.cursor() as cursor:
            query = """
                SELECT 
                    c.id AS id_carga,
                    COALESCE(c.numero_tp, c.numero_carga) AS "Nº TP",
                    c.numero_set AS "Nº SET",
                    c.data_coleta AS "Data Programada",
                    c.status AS "Status",
                    COALESCE(c.nome_motorista, '-') AS "Motorista",
                    COALESCE(c.tipo_motorista, '-') AS "Tipo Motorista",
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
            cursor.execute(query)
            dados = cursor.fetchall()
            colunas = [desc[0] for desc in cursor.description]
            df = pd.DataFrame(dados, columns=colunas)

    if df.empty:
        st.info("Nenhuma carga cadastrada até o momento.")
    else:
        def formatar_real(valor):
            return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        # Formatações de Rotas e Clientes
        df["Origem"] = df["origens_json"].apply(lambda x: formatar_origens_destinos(x))
        df["Destino"] = df["destinos_json"].apply(lambda x: formatar_origens_destinos(x))
        df["Cliente Origem"] = df["origens_json"].apply(lambda x: " | ".join([i.get("cliente","") for i in processar_json_lista(x) if isinstance(i, dict)]))
        df["Cliente Destino"] = df["destinos_json"].apply(lambda x: " | ".join([i.get("cliente","") for i in processar_json_lista(x) if isinstance(i, dict)]))

        # Cálculos Financeiros
        df["Receita Total (R$)"] = df["receita_frete"] + df["receita_pedagio"] + df["receita_taxa_descarga"]
        df["Custo Total (R$)"] = df["RPA"] + df["Pedágio Pago"] + df["Custo Descarga"]
        df["Margem (R$)"] = df["Receita Total (R$)"] - df["Custo Total (R$)"]
        df["Margem (%)"] = df.apply(lambda r: (r["Margem (R$)"] / r["Receita Total (R$)"] * 100) if r["Receita Total (R$)"] > 0 else 0.0, axis=1)

        # -------------------------------------------------------------
        # LIMPEZA DAS COLUNAS JSON (REMOVE ASPAS E COLCHETES)
        # -------------------------------------------------------------
        df["CT-e"] = df["CT-e"].apply(lambda x: limpar_formato_json_lista(x, separador=", "))
        df["Viagem"] = df["Viagem"].apply(lambda x: limpar_formato_json_lista(x, separador=", "))
        df["MDF-e"] = df["MDF-e"].apply(lambda x: limpar_formato_json_lista(x, separador=", "))
        df["Nota Fiscal"] = df["Nota Fiscal"].apply(lambda x: limpar_formato_json_lista(x, separador=" / "))
        # -------------------------------------------------------------

        # Exibição de Métricas
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

        # Botão Exportar Excel
        st.subheader("📥 Exportar Relatório")
        col_exp1, _ = st.columns([1, 3])
        with col_exp1:
            st.download_button(
                "📗 Exportar para Excel (.xlsx)", 
                data=gerar_excel_geral(df[cols_final]), 
                file_name="Relatorio_Geral_Transpes.xlsx", 
                use_container_width=True,
                key="btn_export_excel_visao_geral"
            )

        st.markdown("---")
        st.subheader("📋 Detalhamento das Cargas")

        # Tabela Formatada para Exibição
        df_exib = df.copy()
        cols_fin = ["Receita Total (R$)", "RPA", "Pedágio Pago", "Custo Descarga", "Custo Total (R$)", "Margem (R$)"]
        for c in cols_fin:
            df_exib[c] = df_exib[c].apply(formatar_real)
        
        # Margem inteira arredondada
        df_exib["Margem (%)"] = df_exib["Margem (%)"].apply(lambda v: f"{round(v)}%")

        cols_final = [
            "Nº TP", "Nº SET", "Data Programada", "Status", "Motorista", 
            "Tipo Motorista", "Origem", "Destino", "Cliente Origem", 
            "Cliente Destino", "Receita Total (R$)", "RPA", "Pedágio Pago", 
            "Custo Descarga", "Fornecedor Descarga", "Data Agendamento Descarga", 
            "CT-e", "Viagem", "MDF-e", "Contrato", "Nota Fiscal", 
            "Tipo de Pedágio", "Plataforma", "Vínculo",
            "Custo Total (R$)", "Margem (R$)", "Margem (%)", "Data Pagamento Saldo"
        ]

        st.dataframe(df_exib[cols_final], use_container_width=True, hide_index=True)

# 2. COMERCIAL
elif menu_selecionado == "Comercial":
    st.title("💼 Comercial - Novo Frete / Carga")

    if "com_form_version" not in st.session_state:
        st.session_state["com_form_version"] = 0

    v_com = st.session_state["com_form_version"]
    if st.session_state.get("exibir_modal_comercial", False):
        exibir_popup_comercial(st.session_state.get("carga_comercial_num", ""))

    c_q1, c_q2, c_q3 = st.columns(3)
    qtd_origens = c_q1.number_input("Qtd. Clientes / Origens*", 1, 10, 1, key=f"com_q_orig_{v_com}")
    qtd_destinos = c_q2.number_input("Qtd. Clientes / Destinos*", 1, 10, 1, key=f"com_q_dest_{v_com}")
    qtd_nfs = c_q3.number_input("Qtd. Notas Fiscais (Opcional)", 0, 10, 0, key=f"com_q_nfs_{v_com}")

    with st.form(f"form_comercial_{v_com}"):
        st.subheader("1. Identificação Geral")
        numero_set = st.text_input("Número do SET*", placeholder="Ex: SET-2026-001").strip().upper()

        st.subheader("2. Locais de Origem e Clientes")
        lista_origens = []
        for i in range(int(qtd_origens)):
            co1, co2, co3 = st.columns([2, 2, 1])
            cli = co1.text_input(f"Cliente Origem {i+1}*", key=f"com_cli_orig_{i}_{v_com}").upper()
            cid = co2.text_input(f"Cidade Origem {i+1}*", key=f"com_cid_orig_{i}_{v_com}").upper()
            uf = co3.text_input(f"UF Origem {i+1}*", key=f"com_uf_orig_{i}_{v_com}").upper()
            lista_origens.append({"cliente": cli, "cidade": cid, "estado": uf})

        st.subheader("3. Locais de Destino e Clientes")
        lista_destinos = []
        for j in range(int(qtd_destinos)):
            cd1, cd2, cd3 = st.columns([2, 2, 1])
            cli = cd1.text_input(f"Cliente Destino {j+1}*", key=f"com_cli_dest_{j}_{v_com}").upper()
            cid = cd2.text_input(f"Cidade Destino {j+1}*", key=f"com_cid_dest_{j}_{v_com}").upper()
            uf = cd3.text_input(f"UF Destino {j+1}*", key=f"com_uf_dest_{j}_{v_com}").upper()
            lista_destinos.append({"cliente": cli, "cidade": cid, "estado": uf})

        st.subheader("4. Notas Fiscais (Opcional)")
        lista_nfs_com = []
        if qtd_nfs > 0:
            cols_nf = st.columns(min(qtd_nfs, 4))
            for k in range(int(qtd_nfs)):
                val_nf = cols_nf[k % 4].text_input(f"Nota Fiscal {k+1}", key=f"com_nf_{k}_{v_com}").upper()
                if val_nf.strip():
                    lista_nfs_com.append(val_nf.strip())

        st.subheader("5. Valores Recebidos do Cliente")
        v1, v2, v3 = st.columns(3)
        receita_frete = v1.number_input("Valor do Frete (R$)*", min_value=0.0, value=0.0, step=100.0)
        receita_pedagio = v2.number_input("Valor do Pedágio (R$)", min_value=0.0, value=0.0, step=50.0)
        receita_taxa_descarga = v3.number_input("Taxa de Descarga (R$)", min_value=0.0, value=0.0, step=50.0)

        salvar_com = st.form_submit_button("💾 Salvar Lançamento Comercial", type="primary", use_container_width=True)

    if salvar_com:
        valid_orig = all(o["cliente"] and o["cidade"] and o["estado"] for o in lista_origens)
        valid_dest = all(d["cliente"] and d["cidade"] and d["estado"] for d in lista_destinos)

        if not numero_set:
            st.error("O campo SET é obrigatório.")
        elif not valid_orig:
            st.error("Preencha todos os dados de Cliente, Cidade e UF para as origens.")
        elif not valid_dest:
            st.error("Preencha todos os dados de Cliente, Cidade e UF para os destinos.")
        elif receita_frete <= 0:
            st.error("O Valor do Frete é obrigatório e deve ser maior que zero.")
        else:
            num_carga_novo = gerar_numero_carga_novo()
            json_orig = json.dumps(lista_origens)
            json_dest = json.dumps(lista_destinos)
            json_nfs = json.dumps(lista_nfs_com) if lista_nfs_com else ""

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
                           num_carga_novo, numero_set, json_orig, json_dest, json_nfs,
                           receita_frete, receita_pedagio, receita_taxa_descarga
                       ))
                        conn.commit()

                st.session_state["exibir_modal_comercial"] = True
                st.session_state["carga_comercial_num"] = num_carga_novo
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao salvar dados comerciais: {e}")

# 3. PROGRAMAÇÃO
elif menu_selecionado == "Programação":
    st.title("📌 Programação de Cargas")

    if "prog_form_version" not in st.session_state:
        st.session_state["prog_form_version"] = 0

    v_prog = st.session_state["prog_form_version"]
    if st.session_state.get("exibir_modal_programacao", False):
        exibir_popup_programacao(st.session_state.get("tp_programado_num", ""))

    query_prog = """
        SELECT id, numero_carga, numero_set, origens_json, destinos_json, notas_fiscais_comercial
        FROM cargas 
        WHERE status = 'PENDENTE_PROGRAMACAO'
    """
    with get_connection() as conn:
        cargas_prog_df = pd.read_sql_query(query_prog, conn)

    if cargas_prog_df.empty:
        st.info("Nenhuma carga pendente de programação no momento.")
    else:
        opcoes_cargas = {}
        for _, row in cargas_prog_df.iterrows():
            orig_str = formatar_origens_destinos(row['origens_json'], apenas_locais=True)
            dest_str = formatar_origens_destinos(row['destinos_json'], apenas_locais=True)
            nfs_list = processar_json_lista(row['notas_fiscais_comercial'])
            nf_str = ", ".join(nfs_list) if nfs_list else "LIVRE"

            label = f"SET: {row['numero_set']} - ORIGEM: {orig_str} - DESTINO: {dest_str} - NF: {nf_str}"
            opcoes_cargas[label] = row['id']

        selecionada = st.selectbox("Selecione a Carga Comercial*", list(opcoes_cargas.keys()), key=f"prog_sel_{v_prog}")
        carga_id = opcoes_cargas[selecionada]

        with st.form(f"form_programacao_{v_prog}"):
            st.subheader("1. Identificação do Transporte")
            c1, _, _ = st.columns([1, 1, 1])
            numero_tp = c1.text_input("Número de TP*", placeholder="Ex: TP-998877").strip().upper()

            st.subheader("2. Motorista e Veículo")
            m1, m2, m3, m4 = st.columns(4)
            nome_motorista = m1.text_input("Motorista*").upper()
            cpf_motorista = m2.text_input("CPF")
            telefone_motorista = m3.text_input("Telefone")
            tipo_motorista = m4.selectbox("Tipo de Motorista*", ["TERCEIRO", "FROTA", "AGREGADO"])

            v1, v2, v3, v4 = st.columns(4)
            tipo_veiculo = v1.selectbox("Tipo de Veículo*", ["TRUCK", "CARRETA 13m", "CARRETA 15m", "BITREM", "RODOTREM"])
            eixos = v2.number_input("Eixos*", min_value=2, max_value=9, value=6)
            placa_cavalo = v3.text_input("Placa Cavalo*").upper()
            placa_carreta = v4.text_input("Placa Carreta*").upper()

            st.subheader("3. RPA, Datas e Observações")
            d1, d2, d3 = st.columns(3)
            valor_rpa = d1.number_input("RPA (R$)*", min_value=0.0, value=0.0, step=100.0)
            data_coleta = d2.date_input("Data Coleta*", datetime.date.today(), format="DD/MM/YYYY")
            previsao_descarga = d3.date_input("Data Previsão Entrega*", datetime.date.today() + datetime.timedelta(days=2), format="DD/MM/YYYY")

            st.subheader("Dados do Pedágio e Operação")
            col_p1, col_p2, col_p3 = st.columns(3)

            tipo_pedagio = col_p1.text_input("Tipo de Pedágio", placeholder="Ex: Sem Parar, Veloe, ConectCar...").strip().upper()
            plataforma = col_p2.selectbox("Plataforma*", ["APP", "Card"])
            vinculo = col_p3.selectbox("Vínculo*", ["CPF", "CNPJ"])

            obs_prog = st.text_area("Observação").upper()

            salvar_prog = st.form_submit_button("💾 Salvar Programação", type="primary", use_container_width=True)

        if salvar_prog:
            if not numero_tp:
                st.error("O número de TP é obrigatório.")
            elif not nome_motorista or not placa_cavalo or not placa_carreta:
                st.error("Preencha os campos obrigatórios (Motorista, Placa Cavalo e Placa Carreta).")
            else:
                cpf_fmt = formatar_cpf(cpf_motorista)
                tel_fmt = formatar_telefone(telefone_motorista)
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
                                numero_tp, nome_motorista, cpf_fmt, tel_fmt, tipo_motorista,
                                tipo_veiculo, eixos, placa_cavalo, placa_carreta, valor_rpa,
                                data_coleta.strftime("%d/%m/%Y"), previsao_descarga.strftime("%d/%m/%Y"),
                                tipo_pedagio, plataforma, vinculo, obs_prog, carga_id
                            ))
                            conn.commit()

                    st.session_state["exibir_modal_programacao"] = True
                    st.session_state["tp_programado_num"] = numero_tp
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao salvar programação: {e}")

# 4. EXPEDIÇÃO
elif menu_selecionado == "Expedição":
    st.title("📦 Expedição")

    if "exp_form_version" not in st.session_state:
        st.session_state["exp_form_version"] = 0

    v_exp = st.session_state["exp_form_version"]
    if st.session_state.get("exibir_modal_expedicao", False):
        exibir_popup_expedicao(st.session_state.get("exp_tp_num", ""))

    query_exp = """
        SELECT id, numero_set, numero_tp, nome_motorista, origens_json, destinos_json, notas_fiscais_comercial, valor_rpa
        FROM cargas 
        WHERE status = 'PROGRAMADA'
    """
    with get_connection() as conn:
        cargas_exp_df = pd.read_sql_query(query_exp, conn)

    if cargas_exp_df.empty:
        st.info("Nenhuma carga aguardando expedição.")
    else:
        opcoes_cargas = {}
        rpa_map = {}
        nf_com_map = {}
        tp_map = {}
        for _, row in cargas_exp_df.iterrows():
            orig_str = formatar_origens_destinos(row['origens_json'], apenas_locais=True)
            dest_str = formatar_origens_destinos(row['destinos_json'], apenas_locais=True)
            nfs_list = processar_json_lista(row['notas_fiscais_comercial'])
            nf_str = ", ".join(nfs_list) if nfs_list else "LIVRE"

            label = f"SET: {row['numero_set']} - TP: {row['numero_tp']} - MOTORISTA: {row['nome_motorista']} - ORIGEM: {orig_str} - DESTINO: {dest_str} - NF: {nf_str}"
            cid = row['id']
            opcoes_cargas[label] = cid
            rpa_map[cid] = row['valor_rpa'] or 0.0
            nf_com_map[cid] = nfs_list
            tp_map[cid] = row['numero_tp']

        selecionada = st.selectbox("Selecione a Carga Programada*", list(opcoes_cargas.keys()), key=f"exp_sel_{v_exp}")
        carga_id = opcoes_cargas[selecionada]

        v_rpa = rpa_map[carga_id]
        adiantamento_val = v_rpa * 0.70
        nfs_comercial_existentes = nf_com_map[carga_id]

        e_c1, e_c2, e_c3, e_c4 = st.columns(4)
        qtd_ctes = e_c1.number_input("Qtd. CT-es*", 1, 10, 1, key=f"exp_q_cte_{v_exp}")
        qtd_viagens = e_c2.number_input("Qtd. Viagens", 1, 10, 1, key=f"exp_q_v_{v_exp}")
        qtd_mdfes = e_c3.number_input("Qtd. MDF-es*", 1, 10, 1, key=f"exp_q_mdfe_{v_exp}")
        qtd_nfs_exp = e_c4.number_input("Qtd. NFs (caso livre)", 0, 10, 0 if nfs_comercial_existentes else 1, key=f"exp_q_nf_{v_exp}")

        with st.form(f"form_expedicao_{v_exp}"):
            st.text_input("Adiantamento (70% do RPA - Consulta)", value=f"R$ {adiantamento_val:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), disabled=True)

            st.subheader("1. CT-e e Viagem")
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

            st.subheader("2. MDF-e e Contrato")
            c_m1, c_m2 = st.columns(2)
            lista_mdfes = []
            cols_mdfe = c_m1.columns(min(int(qtd_mdfes), 2))
            for k in range(int(qtd_mdfes)):
                v_m = cols_mdfe[k % 2].text_input(f"MDF-e {k+1}*", key=f"exp_mdfe_{k}_{v_exp}").upper()
                if v_m.strip(): lista_mdfes.append(v_m.strip())

            numero_contrato = c_m2.text_input("Número do Contrato*", placeholder="Ex: CTR-12345").strip().upper()

            st.subheader("3. Custos e Notas Fiscais")
            p1, p2 = st.columns(2)
            valor_pedagio_pago = p1.number_input("Valor Pedágio Pago ao Motorista (R$)", min_value=0.0, value=0.0, step=50.0)
            data_saida_filial = p2.date_input("Data Saída Filial*", datetime.date.today(), format="DD/MM/YYYY")

            lista_nfs_exp = []
            if nfs_comercial_existentes:
                st.text_input("Notas Fiscais (Preenchidas pelo Comercial - Somente Consulta)", value=", ".join(nfs_comercial_existentes), disabled=True)
            else:
                st.markdown("**Preenchimento de Notas Fiscais (Liberado pois Comercial deixou em branco):**")
                if qtd_nfs_exp > 0:
                    cols_nf = st.columns(min(int(qtd_nfs_exp), 4))
                    for n in range(int(qtd_nfs_exp)):
                        v_nf = cols_nf[n % 4].text_input(f"Nota Fiscal {n+1}", key=f"exp_nf_{n}_{v_exp}").upper()
                        if v_nf.strip(): lista_nfs_exp.append(v_nf.strip())

            obs_exp = st.text_area("Observação").upper()

            salvar_exp = st.form_submit_button("🚚 Confirmar Expedição / Saída", type="primary", use_container_width=True)

        if salvar_exp:
            if len(lista_ctes) < int(qtd_ctes):
                st.error("Preencha todos os campos obrigatórios de CT-e.")
            elif len(lista_mdfes) < int(qtd_mdfes):
                st.error("Preencha todos os campos obrigatórios de MDF-e.")
            else:
                json_ctes = json.dumps(lista_ctes)
                json_viagens = json.dumps(lista_viagens) if lista_viagens else ""
                json_mdfes = json.dumps(lista_mdfes)
                json_nfs_exp = json.dumps(lista_nfs_exp) if lista_nfs_exp else ""

                try:
                    with get_connection() as conn:
                        with conn.cursor() as cursor:
                            cursor.execute("""
                                INSERT INTO carga_expedicao (
                                    carga_id, numero_cte, numero_viagem, numero_mdfe, numero_contrato,
                                    valor_pedagio_pago, notas_fiscais_expedicao, data_saida_filial, observacoes_expedicao
                                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                            """, (
                                carga_id, json_ctes, json_viagens, json_mdfes, numero_contrato,
                                valor_pedagio_pago, json_nfs_exp, data_saida_filial.strftime("%d/%m/%Y"), obs_exp
                            ))
                            cursor.execute("UPDATE cargas SET status = 'EM TRÂNSITO' WHERE id = %s", (carga_id,))
                            conn.commit()

                    st.session_state["exibir_modal_expedicao"] = True
                    st.session_state["exp_tp_num"] = tp_map[carga_id]
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao salvar expedição: {e}")

# 5. OPERACIONAL
elif menu_selecionado == "Operacional":
    st.title("⚙️ Operacional - Gestão de Descarga")

    if "op_form_version" not in st.session_state:
        st.session_state["op_form_version"] = 0

    v_op = st.session_state["op_form_version"]
    if st.session_state.get("exibir_modal_operacional", False):
        exibir_popup_operacional(st.session_state.get("op_tp_num", ""))

    query_op = """
        SELECT 
            c.id, c.numero_tp, c.nome_motorista, c.origens_json, c.destinos_json,
            c.previsao_descarga, e.data_saida_filial
        FROM cargas c
        INNER JOIN carga_expedicao e ON c.id = e.carga_id
        WHERE c.status = 'EM TRÂNSITO'
    """
    with get_connection() as conn:
        cargas_op_df = pd.read_sql_query(query_op, conn)

    if cargas_op_df.empty:
        st.info("Nenhuma carga em trânsito aguardando operacional.")
    else:
        opcoes_cargas = {}
        dados_consulta = {}
        tp_map = {}
        for _, row in cargas_op_df.iterrows():
            orig_str = formatar_origens_destinos(row['origens_json'], apenas_locais=True)
            dest_str = formatar_origens_destinos(row['destinos_json'], apenas_locais=True)

            label = f"TP: {row['numero_tp']} - MOTORISTA: {row['nome_motorista']} - ORIGEM: {orig_str} - DESTINO: {dest_str} - SAÍDA: {row['data_saida_filial']} - PREV. DESCARGA: {row['previsao_descarga']}"
            cid = row['id']
            opcoes_cargas[label] = cid
            dados_consulta[cid] = row
            tp_map[cid] = row['numero_tp']

        selecionada = st.selectbox("Selecione a Carga em Trânsito*", list(opcoes_cargas.keys()), key=f"op_sel_{v_op}")
        carga_id = opcoes_cargas[selecionada]
        row_sel = dados_consulta[carga_id]

        with get_connection() as conn:
            forn_df = pd.read_sql_query("SELECT razao_social, nome_fantasia FROM fornecedores ORDER BY razao_social", conn)
        
        lista_fornecedores = [f"{r['razao_social']} ({r['nome_fantasia']})" if r['nome_fantasia'] else r['razao_social'] for _, r in forn_df.iterrows()]

        st.markdown("### 📋 Informações em Consulta (Somente Leitura)")
        inf1, inf2 = st.columns(2)
        cli_orig_full = formatar_origens_destinos(row_sel['origens_json'])
        cli_dest_full = formatar_origens_destinos(row_sel['destinos_json'])
        inf1.text_input("Cliente e Origem", value=cli_orig_full, disabled=True)
        inf2.text_input("Cliente e Destino", value=cli_dest_full, disabled=True)

        inf3, inf4 = st.columns(2)
        inf3.text_input("Data Saída Filial", value=row_sel['data_saida_filial'], disabled=True)
        inf4.text_input("Data Previsão Descarga", value=row_sel['previsao_descarga'], disabled=True)

        st.markdown("---")
        with st.form(f"form_operacional_{v_op}"):
            st.subheader("Preenchimento Operacional")
            f1, f2 = st.columns(2)
            if lista_fornecedores:
                fornecedor_sel = f1.selectbox("Fornecedor de Descarga*", lista_fornecedores)
            else:
                fornecedor_sel = f1.text_input("Fornecedor de Descarga*").upper()

            custo_descarga = f2.number_input("Custo Descarga (R$)*", min_value=0.0, value=0.0, step=50.0)

            p1, p2, p3 = st.columns(3)
            data_agendamento = p1.date_input("Data Agendamento*", datetime.date.today(), format="DD/MM/YYYY")
            forma_pagamento = p2.selectbox("Forma de Pagamento*", ["PIX", "BOLETO", "DEPÓSITO BANCÁRIO"])
            prazo_dias = p3.number_input("Prazo (dias)*", min_value=0, value=30, step=5)

            salvar_op = st.form_submit_button("✅ Concluir Registro Operacional", type="primary", use_container_width=True)

        if salvar_op:
            if not fornecedor_sel:
                st.error("Selecione um Fornecedor de Descarga.")
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
                    st.session_state["op_tp_num"] = tp_map[carga_id]
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao salvar operacional: {e}")

# 6. ADMINISTRAÇÃO
elif menu_selecionado == "Administração":
    st.title("💼 Administração e Acerto Financeiro")

    if "adm_form_version" not in st.session_state:
        st.session_state["adm_form_version"] = 0

    v_adm = st.session_state["adm_form_version"]
    if st.session_state.get("exibir_modal_administracao", False):
        exibir_popup_administracao(st.session_state.get("adm_tp_num", ""))

    query_adm = """
        SELECT c.id, c.numero_tp, c.nome_motorista, COALESCE(c.valor_rpa, 0.0) AS valor_rpa, e.numero_cte
        FROM cargas c
        LEFT JOIN carga_expedicao e ON c.id = e.carga_id
        WHERE c.status = 'ENTREGUE'
    """
    with get_connection() as conn:
        cargas_adm_df = pd.read_sql_query(query_adm, conn)

    if cargas_adm_df.empty:
        st.info("Nenhuma carga entregue aguardando acerto administrativo.")
    else:
        opcoes_cargas = {}
        rpa_map = {}
        tp_map = {}
        for _, row in cargas_adm_df.iterrows():
            ctes_list = processar_json_lista(row['numero_cte'])
            cte_str = ", ".join(ctes_list) if ctes_list else "N/A"
            label = f"TP: {row['numero_tp']} - MOTORISTA: {row['nome_motorista']} - CTe: {cte_str}"
            cid = row['id']
            opcoes_cargas[label] = cid
            rpa_map[cid] = row['valor_rpa']
            tp_map[cid] = row['numero_tp']

        selecionada = st.selectbox("Selecione a Carga Entregue*", list(opcoes_cargas.keys()), key=f"adm_sel_{v_adm}")
        carga_id = opcoes_cargas[selecionada]

        v_rpa = rpa_map[carga_id]
        adiantamento_calc = v_rpa * 0.70
        saldo_calc = v_rpa * 0.30

        with st.form(f"form_administracao_{v_adm}"):
            a1, a2 = st.columns(2)
            a1.text_input("Valor Adiantamento (70% RPA)", value=f"R$ {adiantamento_calc:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), disabled=True)
            a2.text_input("Valor Saldo (30% RPA)", value=f"R$ {saldo_calc:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), disabled=True)

            a3, a4 = st.columns(2)
            comprovante = a3.checkbox("Comprovante de Entregue Recebido")
            status_pag = a4.selectbox("Status Pagamento Saldo", ["PENDENTE", "PAGO", "CANCELADO"])

            data_lib = st.date_input("Data Liberação Saldo", datetime.date.today(), format="DD/MM/YYYY")

            salvar_adm = st.form_submit_button("💰 Finalizar Acerto e Liberar Saldo", type="primary", use_container_width=True)

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
                st.session_state["adm_tp_num"] = tp_map[carga_id]
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao salvar administração: {e}")

# 7. FORNECEDORES
elif menu_selecionado == "Fornecedores":
    st.title("🚚 Gestão de Fornecedores")

    col_btn, _ = st.columns([1, 3])
    with col_btn:
        if st.button("➕ Cadastrar Novo Fornecedor", type="primary", use_container_width=True):
            modal_cadastrar_fornecedor()

    st.markdown("---")
    query_forn = """
        SELECT nr_contrato AS "Nr. Contrato", razao_social AS "Razão Social", nome_fantasia AS "Nome Fantasia",
               local_atendimento AS "Local de Atendimento", uf AS "UF", cpf_cnpj AS "CPF/CNPJ", contato AS "Contato"
        FROM fornecedores ORDER BY id DESC
    """
    with get_connection() as conn:
        df_fornecedores = pd.read_sql_query(query_forn, conn)

    st.metric("Total de Fornecedores Cadastrados", len(df_fornecedores))
    st.dataframe(df_fornecedores, use_container_width=True, hide_index=True)

# 8. EXCLUIR CARGAS
elif menu_selecionado == "Excluir Cargas":
    st.title("🗑️ Excluir Cargas")
    
    with get_connection() as conn:
        cargas_df = pd.read_sql_query("SELECT id, numero_carga, COALESCE(numero_tp, '-') AS tp, COALESCE(nome_motorista, '-') AS mot, status FROM cargas", conn)
    
    if cargas_df.empty:
        st.info("Nenhuma carga cadastrada.")
    else:
        opcoes_cargas = {f"Carga: {row['numero_carga']} | TP: {row['tp']} | Motorista: {row['mot']} | Status: {row['status']}": row['id'] for _, row in cargas_df.iterrows()}
        selecionada = st.selectbox("Selecione a Carga a ser excluída*", list(opcoes_cargas.keys()))
        carga_id = opcoes_cargas[selecionada]

        if st.button("❌ Excluir Definitivamente", type="primary"):
            with get_connection() as conn:
                with conn.cursor() as cursor:
                    cursor.execute("DELETE FROM carga_administracao WHERE carga_id = %s", (carga_id,))
                    cursor.execute("DELETE FROM carga_operacional WHERE carga_id = %s", (carga_id,))
                    cursor.execute("DELETE FROM carga_expedicao WHERE carga_id = %s", (carga_id,))
                    cursor.execute("DELETE FROM cargas WHERE id = %s", (carga_id,))
                    conn.commit()
            st.success("Carga excluída com sucesso!")
            st.rerun()

# 9. USUÁRIOS
elif menu_selecionado == "Usuários":
    st.title("👥 Gestão de Usuários")
    
    with get_connection() as conn:
        users_df = pd.read_sql_query("SELECT id, usuario, nome, perfil FROM usuarios", conn)
    st.dataframe(users_df, use_container_width=True)

    st.subheader("➕ Adicionar Novo Usuário")
    with st.form("form_novo_usuario"):
        u1, u2 = st.columns(2)
        novo_usr = u1.text_input("Usuário*").strip().lower()
        novo_nome = u2.text_input("Nome Completo*").strip().upper()

        p1, p2 = st.columns(2)
        nova_senha = p1.text_input("Senha*", type="password")
        novo_perfil = p2.selectbox("Perfil*", ["ADMIN", "COMERCIAL", "PROGRAMACAO", "EXPEDICAO", "OPERACIONAL", "ADMINISTRATIVO"])

        salvar_usr = st.form_submit_button("💾 Salvar Usuário", use_container_width=True, type="primary")

        if salvar_usr:
            if not (novo_usr and novo_nome and nova_senha and novo_perfil):
                st.error("Preencha todos os campos obrigatórios.")
            else:
                try:
                    with get_connection() as conn:
                        with conn.cursor() as cursor:
                            cursor.execute(
                                "INSERT INTO usuarios (usuario, senha, nome, perfil) VALUES (%s, %s, %s, %s)",
                                (novo_usr, hash_senha(nova_senha), novo_nome, novo_perfil)
                            )
                            conn.commit()
                    st.success(f"Usuário {novo_usr} cadastrado com sucesso!")
                    st.rerun()
                except psycopg2.IntegrityError:
                    st.error("Nome de usuário já cadastrado.")
                except Exception as e:
                    st.error(f"Erro ao salvar usuário: {e}")
