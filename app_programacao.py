import io
import base64
import datetime
import hashlib
import json
import psycopg2
import pandas as pd
import streamlit as st
from streamlit_option_menu import option_menu

# Importações para geração de PDF e Excel
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

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
# Substitua SUA_SENHA_AQUI pela senha que você criou na conta do Supabase
DB_URL = "postgresql://postgres:SUA_SENHA_AQUI@db.ddfxntmohwaqfjfvvldd.supabase.co:5432/postgres"

def get_connection():
    return psycopg2.connect(DB_URL)

def hash_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

def init_db():
    conn = get_connection()
    cursor = conn.cursor()

    # Tabela de Usuários
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

    # 1. Programação
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cargas (
            id SERIAL PRIMARY KEY,
            numero_carga TEXT UNIQUE NOT NULL,
            cliente_origem TEXT DEFAULT '',
            cliente_destino TEXT DEFAULT '',
            nome_motorista TEXT NOT NULL,
            cpf_motorista TEXT NOT NULL,
            telefone_motorista TEXT,
            tipo_veiculo TEXT,
            placa_cavalo TEXT NOT NULL,
            placa_carreta TEXT,
            quantidade_eixos INTEGER NOT NULL,
            peso_total REAL NOT NULL,
            tipo_carga TEXT NOT NULL,
            medida_dn TEXT,
            valor_rpa REAL NOT NULL,
            tipo_motorista TEXT NOT NULL,
            cidade_origem TEXT NOT NULL,
            estado_origem TEXT NOT NULL,
            cidade_destino TEXT NOT NULL,
            estado_destino TEXT NOT NULL,
            data_carregamento TEXT NOT NULL,
            previsao_descarga TEXT NOT NULL,
            origens_json TEXT,
            destinos_json TEXT,
            tem_troca_nota INTEGER DEFAULT 0,
            cidade_troca_nota TEXT,
            estado_troca_nota TEXT,
            data_troca_nota TEXT,
            status TEXT DEFAULT 'PROGRAMADA',
            data_criacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # 2. Expedição
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS carga_expedicao (
            carga_id INTEGER PRIMARY KEY REFERENCES cargas(id) ON DELETE CASCADE,
            numero_set TEXT,
            numero_viagem TEXT,
            numero_cte TEXT,
            numero_mdfe TEXT,
            numero_nota_fiscal TEXT,
            valor_pedagio_pago REAL DEFAULT 0.00,
            data_saida_filial TEXT NOT NULL,
            observacoes_expedicao TEXT,
            data_emissao TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # 3. Operacional
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS carga_operacional (
            carga_id INTEGER PRIMARY KEY REFERENCES cargas(id) ON DELETE CASCADE,
            data_descarga TEXT NOT NULL,
            receita_frete REAL NOT NULL,
            receita_pedagio REAL DEFAULT 0.00,
            receita_taxa_descarga REAL DEFAULT 0.00,
            fornecedor_descarga TEXT,
            equipamento_descarga TEXT,
            custo_fornecedor_descarga REAL DEFAULT 0.00,
            peso_descarregado REAL,
            observacoes_descarga TEXT,
            receita_total REAL,
            custo_total REAL,
            margem_lucro_reais REAL,
            margem_lucro_pct REAL,
            data_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # 4. Administração
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

    conn.commit()
    conn.close()

init_db()

# ==========================================
# FUNÇÕES AUXILIARES DE EXPORTAÇÃO E IMAGEM
# ==========================================
def get_base64_image(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

def gerar_numero_carga_novo():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM cargas")
    qtd = cursor.fetchone()[0] + 1
    conn.close()
    ano = datetime.datetime.now().year
    return f"TRP-{ano}-{qtd:03d}"

def gerar_excel_geral(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Visao_Geral_Cargas')
        worksheet = writer.sheets['Visao_Geral_Cargas']
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
    output.seek(0)
    return output

def gerar_pdf_geral(df):
    output = io.BytesIO()
    doc = SimpleDocTemplate(
        output,
        pagesize=landscape(A4),
        rightMargin=20,
        leftMargin=20,
        topMargin=20,
        bottomMargin=20
    )
    elements = []
    styles = getSampleStyleSheet()
    
    titulo_style = ParagraphStyle(
        'TituloPDF',
        parent=styles['Heading1'],
        fontSize=16,
        leading=20,
        textColor=colors.HexColor('#0E2F56'),
        alignment=1,
        spaceAfter=15
    )
    
    elements.append(Paragraph("Relatório Geral de Cargas - Transpes", titulo_style))
    elements.append(Spacer(1, 10))
    
    colunas_pdf = [
        "Nº Carga", "Status", "Motorista", "Placa Cavalo", 
        "Origem", "Destino", "Receita (R$)", "Custo (R$)", "Margem (R$)", "Margem (%)"
    ]
    
    cols_existentes = [c for c in colunas_pdf if c in df.columns]
    df_pdf = df[cols_existentes].copy()
    
    table_data = [cols_existentes]
    for row in df_pdf.itertuples(index=False):
        table_data.append([str(val) if val is not None else "" for val in row])
        
    t = Table(table_data, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0E2F56')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8F9FA')])
    ]))
    
    elements.append(t)
    doc.build(elements)
    output.seek(0)
    return output

# ==========================================
# DIÁLOGOS / POP-UPS DE CONFIRMAÇÃO
# ==========================================
@st.dialog("📌 Carga Programada com Sucesso!")
def exibir_popup_programacao(num_carga):
    st.success(f"A carga **{num_carga}** foi cadastrada e seu status está como **PROGRAMADA**!")
    st.write("Os dados foram salvos no sistema e o formulário foi limpo para a próxima programação.")
    if st.button("OK / Próxima Programação", type="primary", use_container_width=True):
        st.session_state["exibir_modal_programacao"] = False
        st.session_state["prog_form_version"] += 1
        st.rerun()

@st.dialog("📦 Carga Expedida com Sucesso!")
def exibir_popup_expedicao(num_carga):
    st.success(f"A carga **{num_carga}** foi expedida e atualizada para **EM TRÂNSITO**!")
    st.write("Os dados foram salvos no sistema e os campos foram limpos para o próximo processo.")
    if st.button("OK / Próxima Carga", type="primary", use_container_width=True):
        st.session_state["exibir_modal_expedicao"] = False
        st.session_state["exp_form_version"] += 1
        st.rerun()

@st.dialog("⚙️ Carga Agendada para Descarga!")
def exibir_popup_operacional(num_carga):
    st.success(f"A carga **{num_carga}** teve o descarregamento registrado e seu status foi atualizado para **ENTREGUE**!")
    st.write("Informações financeiras operacionais gravadas com sucesso.")
    if st.button("OK / Próxima Carga", type="primary", use_container_width=True):
        st.session_state["exibir_modal_operacional"] = False
        st.session_state["op_form_version"] += 1
        st.rerun()

@st.dialog("💼 Processo Finalizado - Saldo Liberado com Sucesso!")
def exibir_popup_administracao(num_carga):
    st.success(f"A carga **{num_carga}** teve o acerto concluído e o status alterado para **FINALIZADA**!")
    st.write("A liberação de saldo foi registrada com sucesso.")
    if st.button("OK / Concluir", type="primary", use_container_width=True):
        st.session_state["exibir_modal_administracao"] = False
        st.session_state["adm_form_version"] += 1
        st.rerun()

# ==========================================
# AUTENTICAÇÃO E SESSÃO (TELA DE LOGIN)
# ==========================================
if "logado" not in st.session_state:
    st.session_state["logado"] = False

if not st.session_state["logado"]:
    try:
        bg_base64 = get_base64_image("fundo_transpes.png")
        bg_style = f"""
            background-image: url("data:image/jpeg;base64,{bg_base64}");
            background-size: cover;
            background-position: center;
            background-repeat: no-repeat;
            background-attachment: fixed;
        """
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
        [data-testid="stSidebar"], [data-testid="stHeader"] {{
            display: none;
        }}
        .stApp {{
            {bg_style}
        }}
        [data-testid="stMainBlockContainer"] {{
            max-width: 420px !important;
            padding-top: 5rem !important;
            padding-bottom: 2rem !important;
            margin: auto !important;
        }}
        .login-card {{
            background: rgba(255, 255, 255, 0.95);
            padding: 25px 20px 18px 20px;
            border-radius: 14px;
            box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.4);
            text-align: center;
            margin-bottom: 20px;
        }}
        .login-subtitle {{
            color: #444444;
            font-size: 13px;
            font-weight: 500;
        }}
        [data-testid="stWidgetLabel"] p {{
            color: #FFFFFF !important;
            font-weight: bold !important;
            font-size: 15px !important;
            text-shadow: 1px 1px 3px rgba(0,0,0,0.8);
        }}
        [data-testid="stNotification"] {{
            background-color: rgba(255, 235, 235, 0.95) !important;
            border-left: 5px solid #FF0000 !important;
        }}
        [data-testid="stNotification"] p {{
            color: #D32F2F !important;
            font-weight: bold !important;
            font-size: 15px !important;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

    st.markdown(f"""
        <div class="login-card">
            {logo_html}
            <div class="login-subtitle">Entre com suas credenciais de acesso</div>
        </div>
    """, unsafe_allow_html=True)

    with st.container():
        usuario_input = st.text_input("Usuário", placeholder="Digite seu usuário")
        senha_input = st.text_input("Senha", type="password", placeholder="Digite sua senha")
        
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Entrar", use_container_width=True, type="primary"):
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute(
    "SELECT usuario, nome, perfil FROM usuarios WHERE usuario = %s AND senha = %s",
    (usuario_input.strip().lower(), hash_senha(senha_input))

            )
            usr = cursor.fetchone()
            conn.close()
            
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
    opcoes.extend(["Programação", "Expedição", "Operacional", "Administração", "Excluir Cargas", "Usuários"])
    icones.extend(["clipboard-plus", "file-earmark-text", "tools", "briefcase", "trash", "people"])
else:
    if perfil == "PROGRAMACAO":
        opcoes.append("Programação")
        icones.append("clipboard-plus")
    elif perfil == "EXPEDICAO":
        opcoes.append("Expedição")
        icones.append("file-earmark-text")
    elif perfil == "OPERACIONAL":
        opcoes.append("Operacional")
        icones.append("tools")
    elif perfil == "ADMINISTRATIVO":
        opcoes.append("Administração")
        icones.append("briefcase")

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

# 1. VISÃO GERAL
if menu_selecionado == "Visão Geral":
    st.title("📊 Visão Geral e Relatórios")
    
    conn = get_connection()
    
    query = """
        SELECT 
            c.numero_carga AS "Nº Carga",
            c.status AS "Status",
            c.nome_motorista AS "Motorista",
            c.placa_cavalo AS "Placa Cavalo",
            c.cidade_origem || '/' || c.estado_origem AS "Origem",
            c.cidade_destino || '/' || c.estado_destino AS "Destino",
            
            -- Receita
            COALESCE(o.receita_frete, 0.0) AS "Receita Frete (R$)",
            COALESCE(o.receita_pedagio, 0.0) AS "Receita Pedágio (R$)",
            COALESCE(o.receita_taxa_descarga, 0.0) AS "Receita Taxa Descarga (R$)",
            (COALESCE(o.receita_frete, 0.0) + COALESCE(o.receita_pedagio, 0.0) + COALESCE(o.receita_taxa_descarga, 0.0)) AS "Receita Total (R$)",
            
            -- Custos
            COALESCE(c.valor_rpa, 0.0) AS "RPA (R$)",
            COALESCE(e.valor_pedagio_pago, 0.0) AS "Pedágio Pago (R$)",
            COALESCE(o.custo_fornecedor_descarga, 0.0) AS "Custo Descarga (R$)",
            (COALESCE(c.valor_rpa, 0.0) + COALESCE(e.valor_pedagio_pago, 0.0) + COALESCE(o.custo_fornecedor_descarga, 0.0)) AS "Custo Total (R$)"
            
        FROM cargas c
        LEFT JOIN carga_expedicao e ON c.id = e.carga_id
        LEFT JOIN carga_operacional o ON c.id = o.carga_id
        LEFT JOIN carga_administracao a ON c.id = a.carga_id
        ORDER BY c.id DESC
    """
    df = pd.read_sql_query(query, conn)
    conn.close()

    if df.empty:
        st.info("Nenhuma carga cadastrada até o momento.")
    else:
        def formatar_real(valor):
            return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        df["Margem (R$)"] = df["Receita Total (R$)"] - df["Custo Total (R$)"]
        df["Margem (%)"] = df.apply(
            lambda r: (r["Margem (R$)"] / r["Receita Total (R$)"] * 100) if r["Receita Total (R$)"] > 0 else 0.0, 
            axis=1
        )

        rec_tot = df["Receita Total (R$)"].sum()
        custo_tot = df["Custo Total (R$)"].sum()
        margem_tot = df["Margem (R$)"].sum()

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Total Cargas", len(df))
        m2.metric("Receita Total", formatar_real(rec_tot))
        m3.metric("Custos Totais", formatar_real(custo_tot))
        m4.metric("Margem Total", formatar_real(margem_tot))
        m5.metric("Margem Média", f"{(margem_tot / rec_tot * 100) if rec_tot > 0 else 0.0:.1f}%")

        st.markdown("##### ⏳ Pendências por Setor")
        
        pend_expedicao = len(df[df["Status"] == "PROGRAMADA"])
        pend_operacional = len(df[df["Status"] == "EM TRÂNSITO"])
        pend_administracao = len(df[df["Status"] == "ENTREGUE"])

        p1, p2, p3 = st.columns(3)
        p1.metric("📦 Pendências EXPEDIÇÃO", f"{pend_expedicao} carga(s)", delta="Aguardando CT-e / Saída", delta_color="off")
        p2.metric("⚙️ Pendências OPERACIONAL", f"{pend_operacional} carga(s)", delta="Em trânsito / A descarregar", delta_color="off")
        p3.metric("💼 Pendências ADMINISTRATIVO", f"{pend_administracao} carga(s)", delta="Aguardando acerto financeiro", delta_color="off")

        st.markdown("---")
        
        st.subheader("📥 Exportar Relatório")
        col_exp1, col_exp2, _ = st.columns([1, 1, 2])

        with col_exp1:
            excel_data = gerar_excel_geral(df)
            st.download_button(
                label="📗 Exportar para Excel (.xlsx)",
                data=excel_data,
                file_name="Relatorio_Geral_Transpes.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        with col_exp2:
            pdf_data = gerar_pdf_geral(df)
            st.download_button(
                label="📕 Exportar para PDF",
                data=pdf_data,
                file_name="Relatorio_Geral_Transpes.pdf",
                mime="application/pdf",
                use_container_width=True
            )

        st.markdown("---")
        st.subheader("📋 Detalhamento das Cargas")
        
        df_exibicao = df.copy()
        colunas_financeiras = [
            "Receita Total (R$)", "RPA (R$)", "Pedágio Pago (R$)", 
            "Custo Descarga (R$)", "Custo Total (R$)", "Margem (R$)"
        ]
        
        for col in colunas_financeiras:
            df_exibicao[col] = df_exibicao[col].apply(formatar_real)

        df_exibicao["Margem (%)"] = df_exibicao["Margem (%)"].apply(lambda v: f"{v:.1f}%")

        colunas_ordenadas = [
            "Nº Carga", "Status", "Motorista", "Origem", "Destino",
            "Receita Total (R$)", "RPA (R$)", "Pedágio Pago (R$)", 
            "Custo Descarga (R$)", "Custo Total (R$)", "Margem (R$)", "Margem (%)"
        ]
        
        st.dataframe(df_exibicao[colunas_ordenadas], use_container_width=True)

# 2. PROGRAMAÇÃO
elif menu_selecionado == "Programação":
    st.title("📌 Programação de Cargas")

    if "prog_form_version" not in st.session_state:
        st.session_state["prog_form_version"] = 0

    v = st.session_state["prog_form_version"]
    
    if st.session_state.get("exibir_modal_programacao", False):
        exibir_popup_programacao(st.session_state.get("carga_programada_num", ""))
    
    num_carga_sugerido = gerar_numero_carga_novo()

    st.subheader("1. Identificação da Carga")
    c1, _, _ = st.columns([1, 1, 1])
    numero_carga = c1.text_input("Número da Carga*", value=num_carga_sugerido, key=f"prog_num_carga_{v}").upper()

    st.subheader("2. Dados do Motorista e Veículo")
    m1, m2, m3 = st.columns(3)
    nome_motorista = m1.text_input("Nome do Motorista*", key=f"prog_nome_mot_{v}").upper()
    cpf_motorista = m2.text_input("CPF do Motorista*", key=f"prog_cpf_mot_{v}")
    telefone_motorista = m3.text_input("Telefone", key=f"prog_tel_mot_{v}")

    v1, v2, v3, v4, v5 = st.columns(5)
    tipo_motorista = v1.selectbox("Tipo de Motorista*", ["TERCEIRO", "FROTA", "AGREGADO"], key=f"prog_tipo_mot_{v}")
    tipo_veiculo = v2.selectbox("Tipo de Veículo*", ["CARRETA", "BITREM", "RODOTREM", "VANDERLÉIA", "TOCO", "TRUCK"], key=f"prog_tipo_veic_{v}")
    placa_cavalo = v3.text_input("Placa Cavalo*", key=f"prog_placa_cavalo_{v}").upper()
    placa_carreta = v4.text_input("Placa Carreta", key=f"prog_placa_carreta_{v}").upper()
    qtd_eixos = v5.number_input("Qtd. Eixos*", min_value=2, max_value=9, value=6, key=f"prog_eixos_{v}")

    st.subheader("3. Especificações da Carga")
    e1, e2, e3, e4 = st.columns(4)
    peso_total = e1.number_input("Peso Total (ton)*", min_value=0.1, value=30.0, step=0.5, key=f"prog_peso_{v}")
    tipo_carga = e2.text_input("Tipo da Carga*", value="GERAL", key=f"prog_tipo_carga_{v}").upper()
    medida_dn = e3.text_input("Medida DN", key=f"prog_medida_dn_{v}").upper()
    valor_rpa = e4.number_input("Valor RPA (R$)*", min_value=0.0, value=0.0, step=100.0, key=f"prog_rpa_{v}")

    st.subheader("4. Rotas e Clientes (Múltiplas Origens e Destinos)")
    
    col_qtd_orig, col_qtd_dest = st.columns(2)
    qtd_origens = col_qtd_orig.number_input("Qtd. de Clientes / Locais de Origem*", min_value=1, max_value=10, value=1, key=f"prog_qtd_origens_{v}")
    qtd_destinos = col_qtd_dest.number_input("Qtd. de Clientes / Locais de Destino*", min_value=1, max_value=10, value=1, key=f"prog_qtd_destinos_{v}")

    lista_origens = []
    st.markdown("**📍 Locais e Clientes de Origem:**")
    for i in range(int(qtd_origens)):
        st.caption(f"Origem {i+1}")
        co1, co2, co3 = st.columns([2, 2, 1])
        cli_orig = co1.text_input(f"Cliente Origem {i+1}*", key=f"cli_orig_{i}_{v}").upper()
        cid_orig = co2.text_input(f"Cidade Origem {i+1}*", value="CONTAGEM" if i == 0 else "", key=f"cid_orig_{i}_{v}").upper()
        est_orig = co3.text_input(f"UF Origem {i+1}*", value="MG" if i == 0 else "", key=f"est_orig_{i}_{v}").upper()
        lista_origens.append({"cliente": cli_orig, "cidade": cid_orig, "estado": est_orig})

    st.markdown("<br>", unsafe_allow_html=True)
    lista_destinos = []
    st.markdown("**🏁 Locais e Clientes de Destino:**")
    for j in range(int(qtd_destinos)):
        st.caption(f"Destino {j+1}")
        cd1, cd2, cd3 = st.columns([2, 2, 1])
        cli_dest = cd1.text_input(f"Cliente Destino {j+1}*", key=f"cli_dest_{j}_{v}").upper()
        cid_dest = cd2.text_input(f"Cidade Destino {j+1}*", key=f"cid_dest_{j}_{v}").upper()
        est_dest = cd3.text_input(f"UF Destino {j+1}*", key=f"est_dest_{j}_{v}").upper()
        lista_destinos.append({"cliente": cli_dest, "cidade": cid_dest, "estado": est_dest})

    st.subheader("5. Datas e Troca de Nota")
    d1, d2 = st.columns(2)
    data_carregamento = d1.date_input("Data Carregamento*", datetime.date.today(), format="DD/MM/YYYY", key=f"prog_dt_carreg_{v}")
    previsao_descarga = d2.date_input("Previsão Descarga*", datetime.date.today() + datetime.timedelta(days=2), format="DD/MM/YYYY", key=f"prog_prev_desc_{v}")

    st.markdown("---")
    
    tem_troca_nota = st.checkbox("Houve Troca de Nota?", key=f"chk_troca_nota_{v}")
    
    cidade_troca = ""
    estado_troca = ""
    data_troca_str = ""

    if tem_troca_nota:
        st.markdown("**📋 Dados da Troca de Nota:**")
        tn1, tn2, tn3 = st.columns(3)
        cidade_troca = tn1.text_input("Cidade da Troca de Nota*", key=f"cid_troca_input_{v}").upper()
        estado_troca = tn2.text_input("Estado da Troca de Nota*", key=f"est_troca_input_{v}").upper()
        data_troca = tn3.date_input("Data da Troca de Nota*", datetime.date.today(), format="DD/MM/YYYY", key=f"dt_troca_input_{v}")
        data_troca_str = data_troca.strftime("%d/%m/%Y")

    st.markdown("<br>", unsafe_allow_html=True)
    salvar = st.button("💾 Salvar Programação", type="primary", use_container_width=True)

    if salvar:
        origem_valida = all(o["cliente"] and o["cidade"] and o["estado"] for o in lista_origens)
        destino_valido = all(d["cliente"] and d["cidade"] and d["estado"] for d in lista_destinos)
        troca_valida = not tem_troca_nota or (cidade_troca and estado_troca and data_troca_str)

        if not (numero_carga and nome_motorista and cpf_motorista and placa_cavalo and origem_valida and destino_valido and troca_valida):
            st.error("Por favor, preencha todos os campos obrigatórios (*), incluindo as origens, destinos e troca de nota se marcada.")
        else:
            try:
                conn = get_connection()
                cursor = conn.cursor()
                
                primeira_origem = lista_origens[0]
                primeiro_destino = lista_destinos[0]

                cursor.execute("""
    INSERT INTO cargas (
        numero_carga, cliente_origem, cliente_destino, nome_motorista, cpf_motorista,
        telefone_motorista, tipo_veiculo, placa_cavalo, placa_carreta, quantidade_eixos,
        peso_total, tipo_carga, medida_dn, valor_rpa, tipo_motorista,
        cidade_origem, estado_origem, cidade_destino, estado_destino,
        data_carregamento, previsao_descarga, origens_json, destinos_json,
        tem_troca_nota, cidade_troca_nota, estado_troca_nota, data_troca_nota, status
    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 'PROGRAMADA')
""", (...)) 
                  (
                    numero_carga, primeira_origem["cliente"], primeiro_destino["cliente"], 
                    nome_motorista, cpf_motorista, telefone_motorista, tipo_veiculo, 
                    placa_cavalo, placa_carreta, qtd_eixos, peso_total, tipo_carga, 
                    medida_dn, valor_rpa, tipo_motorista, primeira_origem["cidade"], 
                    primeira_origem["estado"], primeiro_destino["cidade"], primeiro_destino["estado"],
                    data_carregamento.strftime("%d/%m/%Y"), previsao_descarga.strftime("%d/%m/%Y"),
                    json.dumps(lista_origens), json.dumps(lista_destinos),
                    1 if tem_troca_nota else 0, cidade_troca, estado_troca, data_troca_str
                ))
                conn.commit()
                conn.close()

                st.session_state["exibir_modal_programacao"] = True
                st.session_state["carga_programada_num"] = numero_carga
                st.rerun()

            except sqlite3.IntegrityError:
                st.error(f"Erro: O número de carga '{numero_carga}' já existe.")

# 3. EXPEDIÇÃO
elif menu_selecionado == "Expedição":
    st.title("📦 Expedição e Emissão de Documentos")
    
    if "exp_form_version" not in st.session_state:
        st.session_state["exp_form_version"] = 0

    v_exp = st.session_state["exp_form_version"]
    
    if st.session_state.get("exibir_modal_expedicao", False):
        exibir_popup_expedicao(st.session_state.get("carga_expedicao_num", ""))

    conn = get_connection()
    
    query_exp = """
        SELECT 
            c.id, 
            c.numero_carga, 
            c.nome_motorista,
            c.cliente_origem,
            c.cidade_origem,
            c.estado_origem,
            c.cliente_destino,
            c.cidade_destino,
            c.estado_destino,
            COALESCE(c.valor_rpa, 0.0) AS valor_rpa
        FROM cargas c
        WHERE c.status = 'PROGRAMADA'
    """
    cargas_df = pd.read_sql_query(query_exp, conn)
    
    if cargas_df.empty:
        st.info("Nenhuma carga aguardando expedição no momento.")
    else:
        opcoes_cargas = {}
        dados_rpa = {}
        for _, row in cargas_df.iterrows():
            label = (
                f"{row['numero_carga']} - {row['nome_motorista']} | "
                f"{row['cliente_origem']} ({row['cidade_origem']}/{row['estado_origem']}) ➔ "
                f"{row['cliente_destino']} ({row['cidade_destino']}/{row['estado_destino']})"
            )
            opcoes_cargas[label] = (row['id'], row['numero_carga'])
            dados_rpa[row['id']] = row['valor_rpa']

        selecionada = st.selectbox("Selecione a Carga*", list(opcoes_cargas.keys()), key=f"exp_sel_carga_{v_exp}")
        carga_id, num_carga_sel = opcoes_cargas[selecionada]
        
        v_rpa = dados_rpa[carga_id]
        v_adiantamento_calc = v_rpa * 0.70

        st.markdown("---")
        
        st.text_input("Valor Adiantamento (70% do RPA)", value=f"R$ {v_adiantamento_calc:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), disabled=True, key=f"exp_adiant_{v_exp}")

        numero_set = st.text_input("Número do SET*", placeholder="Ex: 55421", key=f"exp_set_{v_exp}").upper()

        st.markdown("---")
        
        # 1. VIAGENS
        st.subheader("🚩 Viagens")
        qtd_viagens = st.selectbox("Quantidade de Viagens", list(range(1, 11)), index=0, key=f"qtd_viagens_{v_exp}")
        lista_viagens = []
        cols_v = st.columns(min(qtd_viagens, 4))
        for i in range(qtd_viagens):
            val_v = cols_v[i % 4].text_input(f"Viagem {i+1}*", key=f"exp_viagem_{i}_{v_exp}").upper()
            if val_v.strip():
                lista_viagens.append(val_v.strip())

        # 2. NOTAS FISCAIS
        st.subheader("📄 Notas Fiscais (NF)")
        qtd_nfs = st.selectbox("Quantidade de Notas Fiscais", list(range(1, 11)), index=0, key=f"qtd_nfs_{v_exp}")
        lista_nfs = []
        cols_nf = st.columns(min(qtd_nfs, 4))
        for i in range(qtd_nfs):
            val_nf = cols_nf[i % 4].text_input(f"Nº Nota Fiscal {i+1}*", key=f"exp_nf_{i}_{v_exp}").upper()
            if val_nf.strip():
                lista_nfs.append(val_nf.strip())

        # 3. CT-e
        st.subheader("📑 Conhecimentos de Transporte (CT-e)")
        qtd_ctes = st.selectbox("Quantidade de CT-es", list(range(1, 11)), index=0, key=f"qtd_ctes_{v_exp}")
        lista_ctes = []
        cols_cte = st.columns(min(qtd_ctes, 4))
        for i in range(qtd_ctes):
            val_cte = cols_cte[i % 4].text_input(f"Nº CT-e {i+1}*", key=f"exp_cte_{i}_{v_exp}").upper()
            if val_cte.strip():
                lista_ctes.append(val_cte.strip())

        # 4. MDF-e
        st.subheader("📋 Manifestos (MDF-e)")
        qtd_mdfes = st.selectbox("Quantidade de MDF-es", list(range(1, 11)), index=0, key=f"qtd_mdfes_{v_exp}")
        lista_mdfes = []
        cols_mdfe = st.columns(min(qtd_mdfes, 4))
        for i in range(qtd_mdfes):
            val_mdfe = cols_mdfe[i % 4].text_input(f"Nº MDF-e {i+1}*", key=f"exp_mdfe_{i}_{v_exp}").upper()
            if val_mdfe.strip():
                lista_mdfes.append(val_mdfe.strip())

        st.markdown("---")

        col_outros1, col_outros2 = st.columns(2)
        valor_pedagio_pago = col_outros1.number_input("Valor Pedágio Pago ao Motorista (R$)", min_value=0.0, value=0.0, key=f"exp_pedagio_{v_exp}")
        data_saida = col_outros2.date_input("Data de Saída / Início Viagem*", datetime.date.today(), format="DD/MM/YYYY", key=f"exp_dt_saida_{v_exp}")

        obs_exp = st.text_area("Observações da Expedição", key=f"exp_obs_{v_exp}").upper()

        st.markdown("<br>", unsafe_allow_html=True)
        salvar_exp = st.button("🚚 Confirmar Saída / Enviar para Trânsito", use_container_width=True, type="primary")

        if salvar_exp:
            if not numero_set.strip():
                st.error("Por favor, preencha o Número do SET.")
            elif len(lista_viagens) < qtd_viagens:
                st.error("Por favor, preencha todos os campos de Viagem selecionados.")
            elif len(lista_nfs) < qtd_nfs:
                st.error("Por favor, preencha todos os campos de Nota Fiscal selecionados.")
            elif len(lista_ctes) < qtd_ctes:
                st.error("Por favor, preencha todos os campos de CT-e selecionados.")
            elif len(lista_mdfes) < qtd_mdfes:
                st.error("Por favor, preencha todos os campos de MDF-e selecionados.")
            else:
                json_viagens = json.dumps(lista_viagens)
                json_nfs = json.dumps(lista_nfs)
                json_ctes = json.dumps(lista_ctes)
                json_mdfes = json.dumps(lista_mdfes)

                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO carga_expedicao (
                        carga_id, numero_set, numero_viagem, numero_cte, numero_mdfe, numero_nota_fiscal,
                        valor_pedagio_pago, data_saida_filial, observacoes_expedicao
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    carga_id, 
                    numero_set.strip(), 
                    json_viagens,
                    json_ctes, 
                    json_mdfes, 
                    json_nfs, 
                    valor_pedagio_pago, 
                    data_saida.strftime("%d/%m/%Y"), 
                    obs_exp
                ))
                
                cursor.execute("UPDATE cargas SET status = 'EM TRÂNSITO' WHERE id = %s", (carga_id,))
                conn.commit()

                st.session_state["exibir_modal_expedicao"] = True
                st.session_state["carga_expedicao_num"] = num_carga_sel
                st.rerun()

    conn.close()

# 4. OPERACIONAL
elif menu_selecionado == "Operacional":
    st.title("⚙️ Operacional e Descarga")
    
    if "op_form_version" not in st.session_state:
        st.session_state["op_form_version"] = 0

    v_op = st.session_state["op_form_version"]

    if st.session_state.get("exibir_modal_operacional", False):
        exibir_popup_operacional(st.session_state.get("carga_operacional_num", ""))

    conn = get_connection()
    
    query_op = """
        SELECT 
            c.id, 
            c.numero_carga, 
            c.nome_motorista,
            c.cliente_origem,
            c.cidade_origem,
            c.estado_origem,
            c.cliente_destino,
            c.cidade_destino,
            c.estado_destino,
            c.previsao_descarga,
            e.numero_cte
        FROM cargas c
        LEFT JOIN carga_expedicao e ON c.id = e.carga_id
        WHERE c.status = 'EM TRÂNSITO'
    """
    cargas_df = pd.read_sql_query(query_op, conn)
    
    if cargas_df.empty:
        st.info("Nenhuma carga em trânsito no momento.")
    else:
        opcoes_cargas = {}
        dados_cargas = {}
        for _, row in cargas_df.iterrows():
            cte_str = "N/A"
            if row['numero_cte']:
                try:
                    ctes = json.loads(row['numero_cte'])
                    cte_str = ", ".join(ctes) if isinstance(ctes, list) else str(ctes)
                except Exception:
                    cte_str = str(row['numero_cte'])

            label = (
                f"{row['numero_carga']} - {row['nome_motorista']} | CTe: {cte_str} | "
                f"{row['cliente_origem']} ({row['cidade_origem']}/{row['estado_origem']}) ➔ "
                f"{row['cliente_destino']} ({row['cidade_destino']}/{row['estado_destino']})"
            )
            opcoes_cargas[label] = (row['id'], row['numero_carga'])
            dados_cargas[row['id']] = row['previsao_descarga']

        selecionada = st.selectbox("Selecione a Carga*", list(opcoes_cargas.keys()), key=f"op_sel_carga_{v_op}")
        carga_id, num_carga_sel = opcoes_cargas[selecionada]
        prev_descarga_val = dados_cargas[carga_id]

        st.markdown("---")
        o1, o2, o3 = st.columns(3)
        o1.text_input("Previsão de Descarga", value=prev_descarga_val, disabled=True, key=f"op_prev_{v_op}")
        receita_frete = o2.number_input("Receita Frete (R$)*", min_value=0.0, value=0.0, key=f"op_rec_frete_{v_op}")
        receita_pedagio = o3.number_input("Receita Pedágio (R$)", min_value=0.0, value=0.0, key=f"op_rec_pedagio_{v_op}")

        o4, o5, o6 = st.columns(3)
        receita_taxa = o4.number_input("Receita Taxa Descarga (R$)", min_value=0.0, value=0.0, key=f"op_rec_taxa_{v_op}")
        fornecedor = o5.text_input("Fornecedor Descarga", key=f"op_fornecedor_{v_op}").upper()
        equipamento = o6.text_input("Equipamento Descarga", key=f"op_equipamento_{v_op}").upper()

        custo_fornecedor = st.number_input("Custo Fornecedor (R$)", min_value=0.0, value=0.0, key=f"op_custo_fornec_{v_op}")
        peso_descarregado = st.number_input("Peso Descarregado (ton)", min_value=0.0, value=0.0, key=f"op_peso_desc_{v_op}")
        obs = st.text_area("Observações", key=f"op_obs_{v_op}").upper()

        st.markdown("<br>", unsafe_allow_html=True)
        salvar_op = st.button("✅ Finalizar Operacional", use_container_width=True, type="primary")

        if salvar_op:
            receita_total = receita_frete + receita_pedagio + receita_taxa
            custo_total = custo_fornecedor
            margem_reais = receita_total - custo_total
            margem_pct = (margem_reais / receita_total * 100) if receita_total > 0 else 0.0

            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO carga_operacional (
                    carga_id, data_descarga, receita_frete, receita_pedagio, receita_taxa_descarga,
                    fornecedor_descarga, equipamento_descarga, custo_fornecedor_descarga,
                    peso_descarregado, observacoes_descarga, receita_total, custo_total,
                    margem_lucro_reais, margem_lucro_pct
                ) VALUES (?, '', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                carga_id, receita_frete, receita_pedagio, receita_taxa,
                fornecedor, equipamento, custo_fornecedor, peso_descarregado, obs,
                receita_total, custo_total, margem_reais, margem_pct
            ))
            
            cursor.execute("UPDATE cargas SET status = 'ENTREGUE' WHERE id = ?", (carga_id,))
            conn.commit()

            st.session_state["exibir_modal_operacional"] = True
            st.session_state["carga_operacional_num"] = num_carga_sel
            st.rerun()

    conn.close()

# 5. ADMINISTRAÇÃO
elif menu_selecionado == "Administração":
    st.title("💼 Administração e Acerto Financeiro")
    
    if "adm_form_version" not in st.session_state:
        st.session_state["adm_form_version"] = 0

    v_adm = st.session_state["adm_form_version"]

    if st.session_state.get("exibir_modal_administracao", False):
        exibir_popup_administracao(st.session_state.get("carga_admin_num", ""))

    conn = get_connection()

    query_adm = """
        SELECT 
            c.id, 
            c.numero_carga, 
            c.nome_motorista,
            c.cliente_origem,
            c.cidade_origem,
            c.estado_origem,
            c.cliente_destino,
            c.cidade_destino,
            c.estado_destino,
            COALESCE(c.valor_rpa, 0.0) AS valor_rpa,
            e.numero_cte
        FROM cargas c
        LEFT JOIN carga_expedicao e ON c.id = e.carga_id
        WHERE c.status = 'ENTREGUE'
    """
    cargas_df = pd.read_sql_query(query_adm, conn)
    
    if cargas_df.empty:
        st.info("Nenhuma carga entregue aguardando acerto administrativo.")
    else:
        opcoes_cargas = {}
        dados_rpa = {}
        for _, row in cargas_df.iterrows():
            cte_str = "N/A"
            if row['numero_cte']:
                try:
                    ctes = json.loads(row['numero_cte'])
                    cte_str = ", ".join(ctes) if isinstance(ctes, list) else str(ctes)
                except Exception:
                    cte_str = str(row['numero_cte'])

            label = (
                f"{row['numero_carga']} - {row['nome_motorista']} | CTe: {cte_str} | "
                f"{row['cliente_origem']} ({row['cidade_origem']}/{row['estado_origem']}) ➔ "
                f"{row['cliente_destino']} ({row['cidade_destino']}/{row['estado_destino']})"
            )
            opcoes_cargas[label] = (row['id'], row['numero_carga'])
            dados_rpa[row['id']] = row['valor_rpa']

        selecionada = st.selectbox("Selecione a Carga*", list(opcoes_cargas.keys()), key=f"adm_sel_carga_{v_adm}")
        carga_id, num_carga_sel = opcoes_cargas[selecionada]
        
        v_rpa = dados_rpa[carga_id]
        adiantamento_calc = v_rpa * 0.70
        saldo_calc = v_rpa * 0.30

        st.markdown("---")
        data_descarga_real = st.date_input("Data Efetiva da Descarga*", datetime.date.today(), format="DD/MM/YYYY", key=f"adm_dt_descarga_{v_adm}")

        a1, a2 = st.columns(2)
        a1.text_input("Valor Adiantamento (70% RPA)", value=f"R$ {adiantamento_calc:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), disabled=True, key=f"adm_val_adiant_{v_adm}")
        a2.text_input("Valor Saldo (30% RPA)", value=f"R$ {saldo_calc:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), disabled=True, key=f"adm_val_saldo_{v_adm}")

        a3, a4 = st.columns(2)
        comprovante = a3.checkbox("Comprovante de Entregue Recebido", key=f"adm_comprovante_{v_adm}")
        status_pag = a4.selectbox("Status Pagamento Saldo", ["PENDENTE", "PAGO", "CANCELADO"], key=f"adm_status_pag_{v_adm}")

        data_lib = st.date_input("Data Liberação Saldo", datetime.date.today(), format="DD/MM/YYYY", key=f"adm_dt_lib_{v_adm}")

        st.markdown("<br>", unsafe_allow_html=True)
        salvar_adm = st.button("💰 Finalizar Acerto e Liberar Saldo", use_container_width=True, type="primary")

        if salvar_adm:
            cursor = conn.cursor()
            
            cursor.execute("""
                UPDATE carga_operacional 
                SET data_descarga = ? 
                WHERE carga_id = ?
            """, (data_descarga_real.strftime("%d/%m/%Y"), carga_id))

            cursor.execute("""
                INSERT INTO carga_administracao (
                    carga_id, valor_adiantamento, valor_saldo, comprovante_entregue,
                    data_liberacao_saldo, status_pagamento_saldo
                ) VALUES (?, ?, ?, ?, ?, ?)
            """, (carga_id, adiantamento_calc, saldo_calc, 1 if comprovante else 0, data_lib.strftime("%d/%m/%Y"), status_pag))
            
            cursor.execute("UPDATE cargas SET status = 'FINALIZADA' WHERE id = ?", (carga_id,))
            conn.commit()

            st.session_state["exibir_modal_administracao"] = True
            st.session_state["carga_admin_num"] = num_carga_sel
            st.rerun()

    conn.close()

# 6. EXCLUIR CARGAS
elif menu_selecionado == "Excluir Cargas":
    st.title("🗑️ Excluir Cargas")
    
    conn = get_connection()
    cargas_df = pd.read_sql_query("SELECT id, numero_carga, nome_motorista, status FROM cargas", conn)
    
    if cargas_df.empty:
        st.info("Nenhuma carga cadastrada.")
    else:
        opcoes_cargas = {f"{row['numero_carga']} - {row['nome_motorista']} ({row['status']})": row['id'] for _, row in cargas_df.iterrows()}
        selecionada = st.selectbox("Selecione a Carga a ser excluída*", list(opcoes_cargas.keys()))
        carga_id = opcoes_cargas[selecionada]

        if st.button("❌ Excluir Definitivamente", type="primary"):
            cursor = conn.cursor()
            cursor.execute("DELETE FROM carga_administracao WHERE carga_id = ?", (carga_id,))
            cursor.execute("DELETE FROM carga_operacional WHERE carga_id = ?", (carga_id,))
            cursor.execute("DELETE FROM carga_expedicao WHERE carga_id = ?", (carga_id,))
            cursor.execute("DELETE FROM cargas WHERE id = %s", (carga_id,))
            conn.commit()
            st.success("Carga excluída com sucesso!")
            st.rerun()
    conn.close()

# 7. USUÁRIOS
elif menu_selecionado == "Usuários":
    st.title("👥 Gestão de Usuários")
    
    conn = get_connection()
    users_df = pd.read_sql_query("SELECT id, usuario, nome, perfil FROM usuarios", conn)
    st.dataframe(users_df, use_container_width=True)

    st.subheader("➕ Adicionar Novo Usuário")
    with st.form("form_novo_usuario"):
        u1, u2 = st.columns(2)
        novo_usr = u1.text_input("Usuário*").strip().lower()
        novo_nome = u2.text_input("Nome Completo*").strip().upper()

        p1, p2 = st.columns(2)
        nova_senha = p1.text_input("Senha*", type="password")
        novo_perfil = p2.selectbox("Perfil*", ["ADMIN", "PROGRAMACAO", "EXPEDICAO", "OPERACIONAL", "ADMINISTRATIVO"])

        salvar_usr = st.form_submit_button("💾 Salvar Usuário", use_container_width=True, type="primary")

        if salvar_usr:
            if not (novo_usr and novo_nome and nova_senha and novo_perfil):
                st.error("Preencha todos os campos obrigatórios.")
            else:
                try:
                    cursor = conn.cursor()
                    cursor.execute(
                        "INSERT INTO usuarios (usuario, senha, nome, perfil) VALUES (?, ?, ?, ?)",
                        (novo_usr, hash_senha(nova_senha), novo_nome, novo_perfil)
                    )
                    conn.commit()
                    st.success(f"Usuário {novo_usr} cadastrado com sucesso!")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("Nome de usuário já cadastrado.")
    conn.close()