import streamlit as st
import pandas as pd
import io

# Função para compatibilidade entre versões
def rerun():
    if hasattr(st, "rerun"):
        st.rerun()
    elif hasattr(st, "experimental_rerun"):
        st.experimental_rerun()

# Função para detectar automaticamente a linha do cabeçalho
def detectar_linha_cabecalho(arquivo, max_linhas=10):
    df_preview = pd.read_excel(arquivo, header=None, nrows=max_linhas)
    melhor_linha = 0
    melhor_score = -float('inf')

    for i in range(len(df_preview)):
        linha = df_preview.iloc[i]
        score = 0
        for valor in linha:
            if isinstance(valor, str) and len(valor.strip()) > 0:
                score += 1
            elif isinstance(valor, (int, float)):
                score -= 1
        if score > melhor_score:
            melhor_score = score
            melhor_linha = i
    return melhor_linha

st.set_page_config(page_title="Editor de Excel com Filtros Dinâmicos", layout="wide")
st.title("📑 Editor de Planilhas com Filtros Dinâmicos e Edição")

# Upload do arquivo
arquivo = st.file_uploader("📁 Envie sua planilha Excel", type=["xlsx"])

# Inicializar número de filtros
if "num_filtros" not in st.session_state:
    st.session_state.num_filtros = 1

if arquivo:
    st.write("🔍 Detectando automaticamente a linha do cabeçalho...")
    linha_detectada = detectar_linha_cabecalho(arquivo)
    st.success(f"✅ Linha detectada como cabeçalho: {linha_detectada}")

    df = pd.read_excel(arquivo, header=linha_detectada)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

    st.write("👀 Pré-visualização dos dados:")
    st.dataframe(df.head())

    st.write("### 📄 Planilha Original", df)

    st.subheader("🔍 Filtros Dinâmicos")
    filtros = []

    # Criar campos de filtro dinamicamente com botão de remover
    for i in range(st.session_state.num_filtros):
        col1, col2, col3 = st.columns([4, 4, 1])
        with col1:
            coluna = st.selectbox(f"Coluna do filtro {i+1}", df.columns, key=f"coluna_{i}")
        with col2:
            valores_unicos = df[coluna].dropna().unique().tolist()
            valores_unicos = [str(v) for v in valores_unicos]
            valor = st.selectbox(f"Valor do filtro {i+1}", valores_unicos, key=f"valor_{i}")
        with col3:
            if st.button("🗑️", key=f"remover_{i}"):
                st.session_state.num_filtros -= 1
                for j in range(i, st.session_state.num_filtros):
                    st.session_state[f"coluna_{j}"] = st.session_state.get(f"coluna_{j+1}", "")
                    st.session_state[f"valor_{j}"] = st.session_state.get(f"valor_{j+1}", "")
                st.session_state.pop(f"coluna_{st.session_state.num_filtros}", None)
                st.session_state.pop(f"valor_{st.session_state.num_filtros}", None)
                rerun()
        filtros.append((coluna, valor))

    if st.button("➕ Adicionar Filtro"):
        st.session_state.num_filtros += 1
        rerun()

    # Aplicar filtros
    df_filtrado = df.copy()
    for coluna, valor in filtros:
        df_filtrado = df_filtrado[df_filtrado[coluna].astype(str) == valor]

    st.write("### 📌 Linhas Filtradas", df_filtrado)

    # Edição avançada — dentro do bloco if arquivo
    if not df_filtrado.empty:
        st.subheader("✏️ Edição Avançada")

        colunas_editar = st.multiselect("Selecione as colunas que deseja alterar", df.columns)
        novos_valores = {}
        for col in colunas_editar:
            novos_valores[col] = st.text_input(f"Novo valor para '{col}'", key=f"novo_{col}")

        aplicar_em_todas = st.checkbox("Aplicar em todas as linhas filtradas", value=True)

        if st.button("✅ Aplicar alteração"):
            condicao = pd.Series(True, index=df.index)
            for coluna, valor in filtros:
                condicao &= df[coluna].astype(str) == valor

            if not aplicar_em_todas:
                condicao = condicao & (condicao.cumsum() == 1)

            for col, novo_valor in novos_valores.items():
                df.loc[condicao, col] = novo_valor

            st.success("✅ Alterações aplicadas com sucesso!")
            st.write("### 📝 Planilha Alterada", df)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            buffer.seek(0)

            st.download_button(
                label="📥 Baixar Planilha Alterada",
                data=buffer,
                file_name="planilha_editada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
