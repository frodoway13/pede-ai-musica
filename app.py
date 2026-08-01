import base64
import datetime
import sqlite3
from contextlib import contextmanager
from io import BytesIO

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image, ImageOps

# =========================================================
# CONFIGURAÇÃO GERAL
# =========================================================
st.set_page_config(
    page_title="Pede Aí - Repertório & Pedidos",
    page_icon="🎵",
    layout="centered",
)

EXCEL_FILE = "PedeAi.xlsx"          # só o repertório vive aqui (somente leitura)
DB_FILE = "pedidos.db"              # pedidos vivem num banco leve (rápido de gravar/atualizar)
CHAVE_PIX = "11977150185"
INSTAGRAM_LINK = "https://www.instagram.com/willllopes?igsh=ZXlkOHhrZXRlYWpu"

VALORES_CAIXINHA = [0.0, 5.0, 10.0, 20.0, 50.0]

# =========================================================
# CSS — visual mais moderno, cartões, cores, responsivo
# =========================================================
st.markdown(
    """
    <style>
        #MainMenu, footer, header {visibility: hidden;}

        .stApp {
            background: linear-gradient(180deg, #0f0c29 0%, #302b63 45%, #24243e 100%);
        }

        .block-container {
            padding-top: 1.5rem;
            max-width: 720px;
        }

        h1, h2, h3, h4, p, label, span, div {
            color: #f4f2ff;
        }

        .hero-card {
            background: rgba(255,255,255,0.06);
            border: 1px solid rgba(255,255,255,0.12);
            border-radius: 20px;
            padding: 1.4rem 1.2rem;
            text-align: center;
            backdrop-filter: blur(6px);
            margin-bottom: 1.2rem;
        }

        .song-card {
            background: rgba(255,255,255,0.07);
            border: 1px solid rgba(255,255,255,0.10);
            border-radius: 16px;
            padding: 1.1rem 1.2rem;
            margin-bottom: 0.9rem;
        }

        .pix-card {
            background: linear-gradient(135deg, #ff8c42 0%, #ff3c78 100%);
            border-radius: 20px;
            padding: 1.4rem;
            text-align: center;
            margin-top: 1rem;
        }

        .badge {
            display: inline-block;
            padding: 2px 10px;
            border-radius: 999px;
            font-size: 0.75rem;
            font-weight: 600;
            margin-left: 6px;
        }
        .badge-pendente { background: #ffd166; color: #3a2e00; }
        .badge-tocando  { background: #06d6a0; color: #003321; }

        div.stButton > button {
            border-radius: 999px;
            font-weight: 600;
            border: none;
            padding: 0.6rem 1.2rem;
            background: linear-gradient(135deg, #ff8c42, #ff3c78);
            color: white;
        }
        div.stButton > button:hover {
            opacity: 0.9;
            color: white;
        }

        [data-testid="stMetricValue"] { color: #ffd166; }

        /* Campo de seleção (fechado) */
        [data-baseweb="select"] > div,
        [data-testid="stSelectbox"] div[data-baseweb="select"] > div {
            background-color: #24243e !important;
            border-color: rgba(255,255,255,0.25) !important;
            color: #f4f2ff !important;
        }
        [data-baseweb="select"] > div * {
            color: #f4f2ff !important;
        }

        /* Menu suspenso (lista de opções) que abre ao clicar
           -> cobre TODAS as variações de nome que o componente
              da Streamlit pode usar dependendo da versão */
        div[data-baseweb="popover"],
        div[data-baseweb="popover"] div,
        div[data-baseweb="menu"],
        ul[data-baseweb="menu"],
        [data-testid="stSelectboxVirtualDropdown"],
        [data-testid="stVirtualDropdown"] {
            background-color: #24243e !important;
        }
        div[data-baseweb="popover"] *,
        div[data-baseweb="menu"] *,
        [data-testid="stSelectboxVirtualDropdown"] *,
        [data-testid="stVirtualDropdown"] * {
            color: #f4f2ff !important;
        }
        div[data-baseweb="popover"] li:hover,
        div[data-baseweb="popover"] li[aria-selected="true"],
        div[data-baseweb="popover"] [role="option"]:hover,
        div[data-baseweb="popover"] [aria-selected="true"] {
            background-color: #ff3c78 !important;
            color: #ffffff !important;
        }

        /* Campo de busca dentro do selectbox */
        [data-baseweb="select"] input {
            color: #f4f2ff !important;
        }

        /* Botão grande do Instagram com foto de fundo */
        .insta-button {
            display: flex;
            align-items: flex-end;
            justify-content: center;
            width: 100%;
            height: 220px;
            border-radius: 20px;
            background-size: cover;
            background-position: center;
            text-decoration: none;
            overflow: hidden;
            position: relative;
            margin-top: 0.6rem;
            box-shadow: 0 8px 24px rgba(0,0,0,0.35);
            transition: transform 0.15s ease;
        }
        .insta-button:hover {
            transform: scale(1.015);
        }
        .insta-button::before {
            content: "";
            position: absolute;
            inset: 0;
            background: linear-gradient(180deg, rgba(0,0,0,0) 40%, rgba(0,0,0,0.75) 100%);
        }
        .insta-button span {
            position: relative;
            z-index: 1;
            width: 100%;
            padding: 1.1rem;
            text-align: center;
            font-size: 1.15rem;
            font-weight: 700;
            color: #ffffff;
        }

        /* Campos de texto (nome, mensagem/dedicatória, chave Pix)
           -> usa data-testid (mais estável entre versões do Streamlit)
              + uma regra coringa em cima de input/textarea como garantia */
        [data-testid="stTextInput"] input,
        [data-testid="stTextArea"] textarea,
        [data-testid="stNumberInput"] input,
        [data-testid="stTextInput"] > div,
        [data-testid="stTextArea"] > div,
        input[type="text"],
        textarea {
            background-color: #24243e !important;
            color: #f4f2ff !important;
            border: 1px solid rgba(255,255,255,0.25) !important;
            -webkit-text-fill-color: #f4f2ff !important;
        }
        [data-testid="stTextInput"] input::placeholder,
        [data-testid="stTextArea"] textarea::placeholder,
        input::placeholder,
        textarea::placeholder {
            color: rgba(244,242,255,0.5) !important;
            -webkit-text-fill-color: rgba(244,242,255,0.5) !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# BANCO DE DADOS (pedidos) — SQLite
# ---------------------------------------------------------
# Por que trocar o Excel por SQLite para os pedidos?
#  - Antes: cada novo pedido reescrevia o arquivo .xlsx INTEIRO
#    (as duas abas, todas as linhas) a cada clique. Isso fica
#    cada vez mais lento conforme a fila cresce e tem risco de
#    corromper o arquivo se dois pedidos chegarem ao mesmo tempo.
#  - Agora: INSERT/UPDATE tocam só a linha necessária, é rápido
#    e seguro mesmo com vários clientes pedindo ao mesmo tempo.
# =========================================================
@contextmanager
def get_conn():
    conn = sqlite3.connect(DB_FILE, timeout=5)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
    finally:
        conn.close()


def init_db():
    with get_conn() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS pedidos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                horario TEXT NOT NULL,
                musica TEXT NOT NULL,
                cliente TEXT NOT NULL,
                mensagem TEXT,
                caixinha REAL DEFAULT 0,
                status TEXT DEFAULT 'Pendente'
            )
            """
        )
        conn.commit()


def inserir_pedido(musica, cliente, mensagem, caixinha):
    with get_conn() as conn:
        conn.execute(
            """INSERT INTO pedidos (horario, musica, cliente, mensagem, caixinha, status)
               VALUES (?, ?, ?, ?, ?, 'Pendente')""",
            (
                datetime.datetime.now().strftime("%H:%M"),
                musica,
                cliente,
                mensagem,
                caixinha,
            ),
        )
        conn.commit()


def marcar_concluido(pedido_id):
    with get_conn() as conn:
        conn.execute("UPDATE pedidos SET status = 'Concluído' WHERE id = ?", (pedido_id,))
        conn.commit()


def carregar_pedidos():
    with get_conn() as conn:
        return pd.read_sql_query("SELECT * FROM pedidos ORDER BY id DESC", conn)


init_db()


def carregar_imagem_corrigida(caminho):
    """Abre a imagem e aplica a rotação correta lendo o EXIF do celular.
    Fotos tiradas com celular guardam a orientação real nos metadados,
    e às vezes isso é ignorado, deixando a foto de cabeça para baixo
    ou de lado."""
    imagem = Image.open(caminho)
    return ImageOps.exif_transpose(imagem)


def imagem_para_base64(caminho):
    """Converte uma imagem local em uma string base64, para poder
    usá-la como background-image dentro de um botão em HTML/CSS
    (o Streamlit não serve arquivos locais por URL diretamente)."""
    imagem = carregar_imagem_corrigida(caminho).convert("RGB")
    buffer = BytesIO()
    imagem.save(buffer, format="JPEG", quality=85)
    return base64.b64encode(buffer.getvalue()).decode()


# =========================================================
# REPERTÓRIO — cache mais generoso (o repertório muda pouco,
# não precisa reler o Excel do disco a cada segundo)
# =========================================================
@st.cache_data(ttl=300, show_spinner=False)
def carregar_repertorio():
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name="Repertorio")
    except Exception as e:
        st.error(f"Erro ao ler a aba 'Repertorio' do arquivo Excel: {e}")
        return pd.DataFrame()


repertorio_df = carregar_repertorio()

query_params = st.query_params
modo_admin = query_params.get("admin") == "1"

# =========================================================
# PAINEL DO CANTOR
# =========================================================
if modo_admin:
    st.title("🎤 Painel de Controle (Palco)")
    st.warning("⚠️ Modo Administrador Ativado")

    col_a, col_b = st.columns([3, 1])
    with col_b:
        if st.button("🔄 Atualizar repertório"):
            carregar_repertorio.clear()
            st.rerun()

    pedidos_atuais = carregar_pedidos()
    pendentes = pedidos_atuais[pedidos_atuais["status"] != "Concluído"]
    concluidos = pedidos_atuais[pedidos_atuais["status"] == "Concluído"]

    m1, m2, m3 = st.columns(3)
    m1.metric("Na fila", len(pendentes))
    m2.metric("Tocadas", len(concluidos))
    m3.metric("Caixinha total (R$)", f"{pedidos_atuais['caixinha'].sum():.2f}")

    st.divider()
    tab_fila, tab_historico = st.tabs(["🎶 Fila de pedidos", "📁 Histórico"])

    with tab_fila:
        if pendentes.empty:
            st.info("Nenhum pedido na fila no momento.")
        else:
            for _, row in pendentes.iterrows():
                val_caixinha = row["caixinha"] or 0.0
                str_caixinha = f" · 💰 R$ {val_caixinha:.2f}" if val_caixinha > 0 else ""
                with st.expander(f"🟡 {row['musica']} — {row['cliente']}{str_caixinha}"):
                    st.write(f"**Horário:** {row['horario']}")
                    st.write(f"**Mensagem:** {row['mensagem'] or '—'}")
                    if st.button("✅ Marcar como Tocada", key=f"btn_{row['id']}"):
                        marcar_concluido(row["id"])
                        st.toast("Pedido concluído!", icon="✅")
                        st.rerun()

    with tab_historico:
        if concluidos.empty:
            st.info("Ainda não há pedidos concluídos.")
        else:
            st.dataframe(
                concluidos[["horario", "musica", "cliente", "caixinha", "status"]],
                use_container_width=True,
                hide_index=True,
            )

# =========================================================
# TELA DO CLIENTE
# =========================================================
else:
    col_img1, col_img2, col_img3 = st.columns([1, 2, 1])
    with col_img2:
        try:
            st.image(carregar_imagem_corrigida("foto_perfil.jpg"), width=150, use_container_width=True)
        except Exception:
            pass

    st.markdown(
        """
        <div class="hero-card">
            <h2>🎵 Pedidos de Música</h2>
            <p>Escolha uma música do repertório, mande um recado e apoie o artista!</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if repertorio_df.empty:
        st.error("A tabela de repertório está vazia ou não foi carregada.")
    else:
        # Resolve nomes de coluna uma única vez
        col_cat = "Categoria" if "Categoria" in repertorio_df.columns else repertorio_df.columns[3]
        col_musica = "Música" if "Música" in repertorio_df.columns else repertorio_df.columns[1]
        col_artista = "Artista_Original" if "Artista_Original" in repertorio_df.columns else repertorio_df.columns[2]

        col_f1, col_f2 = st.columns([1, 1.4])
        with col_f1:
            categorias = ["Todas"] + sorted(repertorio_df[col_cat].dropna().unique().tolist())
            filtro_cat = st.selectbox("Estilo:", categorias)
        with col_f2:
            busca = st.text_input("🔍 Buscar música ou artista:")

        musicas_filtradas = repertorio_df
        if filtro_cat != "Todas":
            musicas_filtradas = musicas_filtradas[musicas_filtradas[col_cat] == filtro_cat]
        if busca:
            musicas_filtradas = musicas_filtradas[
                musicas_filtradas[col_musica].str.contains(busca, case=False, na=False)
                | musicas_filtradas[col_artista].str.contains(busca, case=False, na=False)
            ]

        lista_musicas = musicas_filtradas[col_musica].tolist()

        st.caption(f"{len(lista_musicas)} música(s) encontrada(s)")

        if not lista_musicas:
            st.warning("Nenhuma música encontrada com esse filtro/busca.")
            musica_escolhida = None
        else:
            musica_escolhida = st.selectbox("Selecione a Música:", lista_musicas)

        st.divider()

        if musica_escolhida:
            st.markdown(f'<div class="song-card"><h4>🎤 Pedir: {musica_escolhida}</h4></div>', unsafe_allow_html=True)

            with st.form("form_pedido", clear_on_submit=True):
                nome_cliente = st.text_input("Seu Nome:")
                mensagem = st.text_area("Mensagem / Dedicatória:")
                caixinha = st.selectbox("Apoie o Artista com uma Caixinha (Pix):", VALORES_CAIXINHA)
                enviar = st.form_submit_button("Enviar Pedido ao Palco 🚀", use_container_width=True)

            if enviar:
                if not nome_cliente.strip():
                    st.warning("Por favor, preencha o seu nome antes de enviar!")
                else:
                    inserir_pedido(musica_escolhida, nome_cliente.strip(), mensagem, caixinha)
                    st.session_state["pedido_enviado"] = True
                    st.session_state["valor_caixinha_atual"] = caixinha
                    st.toast("Pedido enviado para o palco! 🎸", icon="🎶")
                    st.rerun()

        if st.session_state.get("pedido_enviado", False):
            st.markdown("<div id='secao_pix'></div>", unsafe_allow_html=True)
            st.markdown(
                """
                <div class="pix-card">
                    <h3>☕ Apoie o Artista</h3>
                    <p>Seu pedido já está no palco! Escaneie o QR Code ou copie a chave Pix abaixo:</p>
                </div>
                """,
                unsafe_allow_html=True,
            )

            try:
                col_qr1, col_qr2, col_qr3 = st.columns([1, 1, 1])
                with col_qr2:
                    st.image(carregar_imagem_corrigida("qrcode_pix.png"), width=230)
            except Exception:
                st.info("Dica: salve o QR Code como 'qrcode_pix.png' na pasta do projeto.")

            st.text_input("Chave Pix:", CHAVE_PIX, key="copia_cola_pix")

            st.markdown("### 📱 Curtiu o som e o atendimento?")
            st.write("Siga o perfil oficial para acompanhar os próximos shows e bastidores:")

            try:
                foto_b64 = imagem_para_base64("foto_perfil.jpg")
                st.markdown(
                    f"""
                    <a href="{INSTAGRAM_LINK}" target="_blank" class="insta-button"
                       style="background-image: url('data:image/jpeg;base64,{foto_b64}');">
                        <span>📸 Seguir no Instagram</span>
                    </a>
                    """,
                    unsafe_allow_html=True,
                )
            except Exception:
                st.link_button("📸 Seguir no Instagram", INSTAGRAM_LINK, use_container_width=True)

            components.html(
                """
                <script>
                    const element = window.parent.document.getElementById('secao_pix');
                    if (element) { element.scrollIntoView({ behavior: 'smooth' }); }
                </script>
                """,
                height=0,
            )
