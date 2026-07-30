import datetime
import pandas as pd
import streamlit as st

# Configuração da página do App
st.set_page_config(
    page_title="Pede Aí - Repertório & Pedidos",
    page_icon="🎵",
    layout="centered",
)

EXCEL_FILE = "PedeAi.xlsx"
CHAVE_PIX = "00020126330014br.gov.bcb.pix0111314090568805204000053039865802BR5911LOWI81421046009Sao Paulo610901227-20062230519daqr36314241445470963044984"  # Coloque sua chave Pix aqui (somente números se for CPF/Telefone)


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
# O cliente só vê a escolha de música. Para ver o painel do cantor,
# o link no navegador precisa terminar com: ?admin=1
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
            f" ({row['Name'] if 'Name' in row else row['Nome do Cliente']}){str_caixinha}"
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

              if caixinha > 0:
                st.divider()
                st.subheader("☕ Apoie o Artista - Pagamento Pix")
                st.info(
                    f"Você escolheu uma caixinha de **R$ {caixinha:.2f}**."
                    " Escaneie o QR Code ou clique no botão abaixo para copiar a"
                    " chave Pix:"
                )

                try:
                  st.image("qrcode_pix.png", width=230)
                except Exception:
                  pass

                # Componente nativo do Streamlit com botão de copiar rápido
                st.code(CHAVE_PIX, language="text")

            except Exception as e:
              st.error(f"Erro ao salvar o pedido na planilha: {e}")
  else:
    st.error("A tabela de repertório está vazia ou não foi carregada.")