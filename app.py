import datetime
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# Configuração da página do App
st.set_page_config(
    page_title="Pede Aí - Repertório & Pedidos",
    page_icon="🎵",
    layout="centered",
)

EXCEL_FILE = "PedeAi.xlsx"
CHAVE_PIX = "12345678900"  # Substitua pela sua chave Pix real (somente números se for CPF/Telefone)
INSTAGRAM_LINK = (
    "https://www.instagram.com/willllopes?igsh=ZXlkOHhrZXRlYWpu"  # Coloque o link do seu Instagram
)


# Função para carregar os dados das abas do Excel
@st.cache_data(ttl=1)
def load_data():
  try:
    df_repertorio = pd.read_excel(EXCEL_FILE, sheet_name="Repertorio")
  except Exception as e:
    st.error(
        f"Erro ao ler a aba 'Repertorio' do arquivo Excel: {e}. Verifique se o"
        " arquivo e o nome da aba estão corretos."
    )
    return pd.DataFrame(), pd.DataFrame()

  try:
    df_pedidos = pd.read_excel(EXCEL_FILE, sheet_name="Pedidos")
  except Exception:
    df_pedidos = pd.DataFrame(
        columns=[
            "Horário",
            "Música Pedida",
            "Nome do Cliente",
            "Mesa / Contato",
            "Mensagem/Recado",
            "Valor Caixinha (R$)",
            "Status do Pedido",
        ]
    )
  return df_repertorio, df_pedidos


repertorio_df, pedidos_df = load_data()

# --- SEGURANÇA E SEPARAÇÃO DE TELAS ---
query_params = st.query_params
modo_admin = query_params.get("admin") == "1"

if modo_admin:
  # --- PAINEL DO CANTOR (Apenas com o link secreto ?admin=1) ---
  st.title("🎤 Painel de Controle (Palco)")
  st.warning("⚠️ Modo Administrador Ativado")
  st.write("Gerencie os pedidos recebidos em tempo real:")

  _, pedidos_atuais = load_data()

  if not pedidos_atuais.empty:
    for index, row in pedidos_atuais.iterrows():
      status = row.get("Status do Pedido", "Pendente")
      if status != "Concluído":
        status_cor = "🟢" if status == "Pendente" else "🟡"
        val_caixinha = row.get("Valor Caixinha (R$)", 0.0)
        str_caixinha = (
            f" - 💰 R$ {val_caixinha:.2f}"
            if pd.notna(val_caixinha) and val_caixinha > 0
            else ""
        )

        with st.expander(
            f"{status_cor} {row['Música Pedida']} — Mesa: {row['Mesa / Contato']}"
            f" ({row['Nome do Cliente']}){str_caixinha}"
        ):
          st.write(f"**Horário:** {row['Horário']}")
          st.write(f"**Mensagem:** {row['Mensagem/Recado']}")
          st.write(
              f"**Caixinha:** R$ {val_caixinha:.2f}"
              if pd.notna(val_caixinha)
              else "**Caixinha:** R$ 0,00"
          )

          if st.button(
              f"Marcar como Tocada / Concluído", key=f"btn_{index}"
          ):
            pedidos_atuais.at[index, "Status do Pedido"] = "Concluído"
            try:
              with pd.ExcelWriter(
                  EXCEL_FILE,
                  engine="openpyxl",
                  mode="a",
                  if_sheet_exists="replace",
              ) as writer:
                repertorio_df.to_excel(
                    writer, sheet_name="Repertorio", index=False
                )
                pedidos_atuais.to_excel(
                    writer, sheet_name="Pedidos", index=False
                )
              st.success("Pedido concluído com sucesso!")
              st.rerun()
            except Exception as e:
              st.error(f"Erro ao atualizar status: {e}")

    with st.expander("📁 Ver Histórico de Pedidos Concluídos"):
      st.dataframe(
          pedidos_atuais[pedidos_atuais["Status do Pedido"] == "Concluído"]
      )
  else:
    st.info("Nenhum pedido na fila no momento.")

else:
  # --- TELA DO CLIENTE (O que o público acessa normalmente) ---

  # 1. Foto do Cantor no Topo (Opcional: se tiver a foto salva como 'foto_perfil.jpg')
  col_img1, col_img2, col_img3 = st.columns([1, 2, 1])
  with col_img2:
    try:
      st.image(
          "foto_perfil.jpg",
          width=150,
          caption="Cantor / Artista",
          use_container_width=True,
      )
    except Exception:
      # Se não tiver a foto ainda, exibe um avatar padrão elegante
      pass

  st.title("🎵 Pedidos de Música & Caixinha")
  st.write(
      "Escolha uma música do repertório, mande um recado e apoie o artista!"
  )

  if not repertorio_df.empty:
    col_cat = (
        "Categoria"
        if "Categoria" in repertorio_df.columns
        else repertorio_df.columns[3]
    )
    categorias = ["Todas"] + list(repertorio_df[col_cat].dropna().unique())
    filtro_cat = st.selectbox("Filtrar por Estilo:", categorias)

    if filtro_cat != "Todas":
      musicas_filtradas = repertorio_df[repertorio_df[col_cat] == filtro_cat]
    else:
      musicas_filtradas = repertorio_df

    busca = st.text_input("🔍 Buscar música ou artista:")
    col_musica = (
        "Música"
        if "Música" in repertorio_df.columns
        else repertorio_df.columns[1]
    )
    col_artista = (
        "Artista_Original"
        if "Artista_Original" in repertorio_df.columns
        else repertorio_df.columns[2]
    )

    if busca:
      musicas_filtradas = musicas_filtradas[
          musicas_filtradas[col_musica].str.contains(
              busca, case=False, na=False
          )
          | musicas_filtradas[col_artista].str.contains(
              busca, case=False, na=False
          )
      ]

    lista_musicas = musicas_filtradas[col_musica].values
    if len(lista_musicas) > 0:
      musica_escolhida = st.selectbox("Selecione a Música:", lista_musicas)
    else:
      musica_escolhida = None
      st.warning("Nenhuma música encontrada com esse filtro/busca.")

    st.divider()

    if musica_escolhida:
      with st.form("form_pedido", clear_on_submit=True):
        st.subheader(f"Pedir: {musica_escolhida}")
        nome_cliente = st.text_input("Seu Nome:")
        mesa = st.text_input("Sua Mesa / Local (Ex: Mesa 04):")
        mensagem = st.text_area("Mensagem / Dedicatória:")
        caixinha = st.selectbox(
            "Apoie o Artista com uma Caixinha (Pix):",
            [0.0, 5.0, 10.0, 20.0, 50.0],
        )

        enviar = st.form_submit_button("Enviar Pedido ao Palco 🚀")

        if enviar:
          if not nome_cliente.strip():
            st.warning("Por favor, preencha o seu nome antes de enviar!")
          else:
            novo_pedido = {
                "Horário": datetime.datetime.now().strftime("%H:%M"),
                "Música Pedida": musica_escolhida,
                "Nome do Cliente": nome_cliente,
                "Mesa / Contato": mesa,
                "Mensagem/Recado": mensagem,
                "Valor Caixinha (R$)": caixinha,
                "Status do Pedido": "Pendente",
            }

            novo_df = pd.DataFrame([novo_pedido])
            atualizado_df = pd.concat(
                [pedidos_df, novo_df], ignore_index=True
            )

            try:
              with pd.ExcelWriter(
                  EXCEL_FILE,
                  engine="openpyxl",
                  mode="a",
                  if_sheet_exists="replace",
              ) as writer:
                repertorio_df.to_excel(
                    writer, sheet_name="Repertorio", index=False
                )
                atualizado_df.to_excel(
                    writer, sheet_name="Pedidos", index=False
                )

              st.success(
                  "Pedido enviado com sucesso para o palco! Fique atento. 🎸"
              )
              st.session_tag_pedido_enviado = True

            except Exception as e:
              st.error(f"Erro ao salvar o pedido na planilha: {e}")

      # Se o pedido foi enviado, exibimos a âncora do Pix e rolando a tela automaticamente para ela
      if st.session_get("session_tag_pedido_enviado", False):
        st.markdown("<div id='secao_pix'></div>", unsafe_allow_html=True)

        st.divider()
        st.subheader("☕ Apoie o Artista - Pagamento Pix")
        st.info(
            "Seu pedido já está no palco! Para dar aquela força, escaneie o"
            " QR Code abaixo ou copie a chave Pix:"
        )

        try:
          st.image("qrcode_pix.png", width=230)
        except Exception:
          st.warning(
              "Dica: Salve o QR Code como 'qrcode_pix.png' na pasta do projeto."
          )

        # Rótulo amigável pedido
        st.text_input(
            "Chave Pix (Copia e Cola):", CHAVE_PIX, key="input_copia_cola"
        )

        # Bloco amigável para o Instagram (sem atrapalhar o Pix)
        st.markdown("---")
        st.write(
            "📸 **Curtiu o som?** Aproveite para seguir o artista nas redes"
            " sociais e acompanhar a agenda de shows!"
        )
        st.link_button("✨ Seguir no Instagram", INSTAGRAM_LINK)

        # Script em JavaScript puro para fazer o navegador rolar suavemente até a seção do Pix
        components.html(
            """
                <script>
                    const element = window.parent.document.getElementById('secao_pix');
                    if (element) {
                        element.scrollIntoView({ behavior: 'smooth' });
                    }
                </script>
                """,
            height=0,
        )
  else:
    st.error("A tabela de repertório está vazia ou não foi carregada.")
