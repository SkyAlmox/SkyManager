# pip install streamlit pandas openpyxl
# streamlit run app.py

import time
import pandas as pd
import streamlit as st

EXCEL_PATH = "data.xlsx"
SHEET_NAME = "Planilha1"
REFRESH_SECONDS = 30

@st.cache_data(ttl=REFRESH_SECONDS)
def load_data():
    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
    df.columns = [str(c).strip().upper() for c in df.columns]
    if "DATA" in df.columns:
        df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce", dayfirst=False)
    for col in ["QUANTIDADE", "PRECO UNIT.", "PRECO TOTAL",
                "DIAMETRO CABECA (MM)", "DIAMETRO ROSCA (MM)",
                "COMPRIMENTO ROSCA (MM)"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df

def sidebar_filters(df):
    st.sidebar.header("Filtros")
    for col in ["DESCRICAO", "CATEGORIA / TIPO", "MATERIAL", "FORNECEDOR", "UNIDADE", "NF", "LOCALIZACAO"]:
        if col in df.columns:
            opts = sorted([x for x in df[col].dropna().astype(str).unique()])
            sel = st.sidebar.multiselect(f"{col}", opts)
            if sel:
                df = df[df[col].astype(str).isin(sel)]
    if "DATA" in df.columns and df["DATA"].notna().any():
        min_d, max_d = df["DATA"].min(), df["DATA"].max()
        start, end = st.sidebar.date_input("Período", [min_d, max_d])
        if isinstance(start, list):
            start, end = start
        df = df[(df["DATA"] >= pd.Timestamp(start)) & (df["DATA"] <= pd.Timestamp(end))]
    return df

def main():
    st.set_page_config(page_title="Dashboard Almoxarifado", layout="wide")
    st.title("Dashboard de Almoxarifado — Planilha1")
    st.caption("Os dados são lidos diretamente do Excel e atualizados periodicamente.")

    placeholder = st.empty()
    with placeholder.container():
        df = load_data()
        df = sidebar_filters(df)

        st.subheader("Tabela")
        cols = st.multiselect("Escolha as colunas para exibir", list(df.columns),
                              default=[c for c in df.columns if c in ["DATA","DESCRICAO","CATEGORIA / TIPO","MATERIAL","UNIDADE","QUANTIDADE","PRECO UNIT.","PRECO TOTAL","FORNECEDOR"]])
        st.dataframe(df[cols] if cols else df, use_container_width=True, height=400)

        st.subheader("Indicadores")
        kpi_cols = st.columns(3)
        total_itens = int(df["QUANTIDADE"].sum()) if "QUANTIDADE" in df.columns else len(df)
        valor_total = float(df["PRECO TOTAL"].sum()) if "PRECO TOTAL" in df.columns else 0.0
        distintos = df["DESCRICAO"].nunique() if "DESCRICAO" in df.columns else len(df.columns)
        kpi_cols[0].metric("Quantidade total", f"{total_itens:,}".replace(",", "."))
        kpi_cols[1].metric("Valor total (R$)", f"{valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        kpi_cols[2].metric("Itens distintos", f"{distintos}")

        st.subheader("Gráficos")
        chart_dim = st.selectbox("Dimensão (eixo X)", [c for c in df.columns if df[c].dtype == "object"] + ["DESCRICAO","CATEGORIA / TIPO","MATERIAL"])
        metric = st.selectbox("Métrica (eixo Y)", [c for c in ["QUANTIDADE","PRECO TOTAL","PRECO UNIT."] if c in df.columns])
        top_n = st.slider("Top N", 5, 50, 15)

        grp = df.groupby(chart_dim, dropna=False)[metric].sum().sort_values(ascending=False).head(top_n)
        st.bar_chart(grp)

        st.subheader("Exportar")
        st.download_button("Baixar dados filtrados (CSV)", df.to_csv(index=False).encode("utf-8"), file_name="planilha1_filtrada.csv", mime="text/csv")

    st.sidebar.info(f"Atualiza a cada {REFRESH_SECONDS}s. Basta salvar o Excel para refletir aqui.")

if __name__ == "__main__":
    main()
