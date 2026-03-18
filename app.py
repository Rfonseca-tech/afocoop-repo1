import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import datetime

st.set_page_config(page_title="AFOOCOP Pricing Simulator", layout="wide")

# ---------------------------------------------------------------------------
# 1. Data Loading & Preprocessing
# ---------------------------------------------------------------------------
@st.cache_data(ttl=5)
def load_and_preprocess_data(file_path):
    try:
        if str(file_path).endswith(".csv"):
            df_raw = pd.read_csv(file_path)
        else:
            df_raw = pd.read_excel(file_path, sheet_name="MASTER_DATA")
    except Exception as e:
        return None, f"Error loading file: {e}"

    # Normalize column names for safe access
    # We want the text label 'Janeiro/2025' to be our 'MONTH' column in the app
    col_map = {
        "MONTH_LABEL": "MONTH",     # The text one becomes the main MONTH
        "MONTH": "MONTH_DATE",      # The datetime one gets moved out of the way
        "PLACA": "LICENSE_PLATE",
        "Valor Equipamento": "EQUIPMENT_VALUE",
        "Cavalo/Carreta": "EQUIPMENT_TYPE",
        "Marca": "EQUIPMENT_BRAND",
        "Modelo": "EQUIPMENT_MODEL",
        "Ano Fabricação": "MANUFACTURE_YEAR",
        "Ano Modelo": "MODEL_YEAR",
        "TIPO_LANCAMENTO": "ENTRY_TYPE",
        "VALOR": "TRANSACTION_AMOUNT"
    }
    
    # Rename columns that exist. Do it carefully to avoid collision
    if "MONTH" in df_raw.columns:
        df_raw.rename(columns={"MONTH": "MONTH_DATE"}, inplace=True)
    if "MONTH_LABEL" in df_raw.columns:
        df_raw.rename(columns={"MONTH_LABEL": "MONTH"}, inplace=True)
        
    for excel_col, app_col in col_map.items():
        if excel_col in ["MONTH", "MONTH_LABEL"]: continue # Already handled
        if excel_col in df_raw.columns:
            df_raw.rename(columns={excel_col: app_col}, inplace=True)
            
    # Ensure required columns exist, fill if missing
    req_cols = [
        "MONTH", "LICENSE_PLATE", "TRANSACTION_AMOUNT", "EQUIPMENT_VALUE", "EQUIPMENT_TYPE",
        "EQUIPMENT_BRAND", "EQUIPMENT_MODEL", "MANUFACTURE_YEAR", "MODEL_YEAR", "ENTRY_TYPE"
    ]
    for c in req_cols:
        if c not in df_raw.columns:
            df_raw[c] = np.nan if c in ["EQUIPMENT_VALUE", "TRANSACTION_AMOUNT", "MANUFACTURE_YEAR", "MODEL_YEAR"] else "Unknown"

    # Convert numeric columns
    df_raw["EQUIPMENT_VALUE"] = pd.to_numeric(df_raw["EQUIPMENT_VALUE"], errors="coerce").fillna(0)
    df_raw["TRANSACTION_AMOUNT"] = pd.to_numeric(df_raw["TRANSACTION_AMOUNT"], errors="coerce").fillna(0)
    df_raw["MANUFACTURE_YEAR"] = pd.to_numeric(df_raw["MANUFACTURE_YEAR"], errors="coerce")
    df_raw["MODEL_YEAR"] = pd.to_numeric(df_raw["MODEL_YEAR"], errors="coerce")
    
    # Fill categorical
    df_raw["EQUIPMENT_TYPE"] = df_raw["EQUIPMENT_TYPE"].fillna("Desconhecido")
    df_raw["EQUIPMENT_BRAND"] = df_raw["EQUIPMENT_BRAND"].fillna("Desconhecido")
    df_raw["EQUIPMENT_MODEL"] = df_raw["EQUIPMENT_MODEL"].fillna("Desconhecido")
    df_raw["ENTRY_TYPE"] = df_raw["ENTRY_TYPE"].fillna("Desconhecido")

    # Group by Month and Plate to get Monthly totals per truck
    agg_funcs = {
        "TRANSACTION_AMOUNT": "sum",
        "EQUIPMENT_VALUE": "first",
        "EQUIPMENT_TYPE": "first",
        "EQUIPMENT_BRAND": "first",
        "EQUIPMENT_MODEL": "first",
        "MANUFACTURE_YEAR": "first",
        "MODEL_YEAR": "first",
        "ENTRY_TYPE": "first"
    }
    
    # If there are other columns like FUNDO, we can keep the first one
    if "FUNDO" in df_raw.columns:
        agg_funcs["FUNDO"] = "first"

    grouped = df_raw.groupby(["MONTH", "LICENSE_PLATE"], as_index=False).agg(agg_funcs)
    grouped.rename(columns={"TRANSACTION_AMOUNT": "CURRENT_PAYMENT"}, inplace=True)
    
    # Create Price Range Buckets
    def get_price_bucket(val):
        if pd.isna(val) or val <= 0: return "Sem Valor Definido"
        if val <= 200000: return "Até R$ 200k"
        if val <= 300000: return "R$ 200k a R$ 300k"
        if val <= 450000: return "R$ 300k a R$ 450k"
        if val <= 600000: return "R$ 450k a R$ 600k"
        return "Acima de R$ 600k"
        
    grouped["EQUIP_VAL_BUCKET"] = grouped["EQUIPMENT_VALUE"].apply(get_price_bucket)

    current_year = datetime.now().year
    grouped["FLEET_YEAR_BASE"] = grouped["MODEL_YEAR"].where(
        grouped["MODEL_YEAR"].notna() & (grouped["MODEL_YEAR"] > 0),
        grouped["MANUFACTURE_YEAR"]
    )
    grouped["FLEET_AGE"] = current_year - grouped["FLEET_YEAR_BASE"]
    grouped.loc[(grouped["FLEET_AGE"] < 0) | (grouped["FLEET_AGE"] > 80), "FLEET_AGE"] = np.nan

    def get_age_bucket(age):
        if pd.isna(age): return "Não informado"
        if age <= 2: return "0 a 2 anos"
        if age <= 5: return "3 a 5 anos"
        if age <= 8: return "6 a 8 anos"
        if age <= 12: return "9 a 12 anos"
        return "13+ anos"

    grouped["AGE_BUCKET"] = grouped["FLEET_AGE"].apply(get_age_bucket)

    return grouped, None


# ---------------------------------------------------------------------------
# App Layout
# ---------------------------------------------------------------------------
st.title("🚚 AFOOCOP: Relatório e Simulador de Rateio")
st.markdown("Analise o custo compartilhado atual e simule novos cenários de cobrança com base em faixas de valor do equipamento.")

import os

# Sidebar - Data Loading
st.sidebar.header("1. Base de Dados")

# Get absolute path for reliability
current_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(current_dir, "AFOOCOP_Rateios_Consolidado.xlsx")

# Read directly from local path
try:
    if os.path.exists(file_path):
        df, error = load_and_preprocess_data(file_path)
        if error:
            st.sidebar.error(f"Erro no processamento: {error}")
            df = None
        else:
            st.sidebar.success("✅ Rateios Consolidados carregados!")
    else:
        df = None
        error = f"Não encontrei o arquivo em: {file_path}"
        st.sidebar.error(error)
except Exception as e:
    st.sidebar.error(f"Erro sistêmico ao ler arquivo: {e}")
    df = None
    error = str(e)

if error or df is None:
    st.info("👈 O arquivo consolidado não foi encontrado na pasta correta.")
    st.stop()

# Helpers para ordenação cronológica dos meses em PT-BR
_MONTH_PT = {"Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4,
             "Maio": 5, "Junho": 6, "Julho": 7, "Agosto": 8,
             "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12}

def _month_sort_key(label):
    try:
        nome, ano = str(label).split("/")
        return (int(ano), _MONTH_PT.get(nome.strip(), 0))
    except Exception:
        return (9999, 0)

all_months_sorted = sorted(df["MONTH"].unique().tolist(), key=_month_sort_key)

# Sidebar - Filters
st.sidebar.header("2. Filtros")

all_months = df["MONTH"].unique().tolist()
selected_months = st.sidebar.multiselect("Filtrar por Mês", all_months_sorted, default=all_months_sorted)

all_equip_types = df["EQUIPMENT_TYPE"].unique().tolist()
selected_types = st.sidebar.multiselect("Filtrar por Tipo de Equip.", all_equip_types, default=all_equip_types)

all_brands = df["EQUIPMENT_BRAND"].unique().tolist()
selected_brands = st.sidebar.multiselect("Filtrar por Marca", all_brands, default=all_brands)

all_buckets = ["Até R$ 200k", "R$ 200k a R$ 300k", "R$ 300k a R$ 450k", "R$ 450k a R$ 600k", "Acima de R$ 600k", "Sem Valor Definido"]
default_buckets = [b for b in all_buckets if b != "Sem Valor Definido"]
selected_buckets = st.sidebar.multiselect("Filtrar por Faixa de Valor do Equipamento", all_buckets, default=default_buckets)

search_plate = st.sidebar.text_input("Buscar por Placa")

# Apply Filters
filtered_df = df[
    (df["MONTH"].isin(selected_months)) &
    (df["EQUIPMENT_TYPE"].isin(selected_types)) &
    (df["EQUIPMENT_BRAND"].isin(selected_brands)) &
    (df["EQUIP_VAL_BUCKET"].isin(selected_buckets))
].copy()

if search_plate.strip():
    filtered_df = filtered_df[filtered_df["LICENSE_PLATE"].str.contains(search_plate.strip().upper(), na=False)]

if filtered_df.empty:
    st.warning("Nenhum dado encontrado para os filtros selecionados.")
    st.stop()

# ---------------------------------------------------------------------------
# 3. Value-Range Bracket Configuration
# ---------------------------------------------------------------------------
st.subheader("Configuração das Faixas de Cobrança (Brackets)")
st.caption("Edite a tabela abaixo para definir as faixas de valor do veículo e a mensalidade proposta para cada uma.")

# Default Brackets using the data-driven distribution
default_brackets = pd.DataFrame([
    {"Nome da Faixa": "Até R$ 200k", "Valor Mínimo (R$)": 0, "Valor Máximo (R$)": 200000, "Mensalidade Simulada (R$)": 80.0},
    {"Nome da Faixa": "R$ 200k a R$ 300k", "Valor Mínimo (R$)": 200000.01, "Valor Máximo (R$)": 300000, "Mensalidade Simulada (R$)": 120.0},
    {"Nome da Faixa": "R$ 300k a R$ 450k", "Valor Mínimo (R$)": 300000.01, "Valor Máximo (R$)": 450000, "Mensalidade Simulada (R$)": 180.0},
    {"Nome da Faixa": "R$ 450k a R$ 600k", "Valor Mínimo (R$)": 450000.01, "Valor Máximo (R$)": 600000, "Mensalidade Simulada (R$)": 250.0},
    {"Nome da Faixa": "Acima de R$ 600k", "Valor Mínimo (R$)": 600000.01, "Valor Máximo (R$)": 99999999, "Mensalidade Simulada (R$)": 350.0},
])

# Editable dataframe for brackets
edited_brackets = st.data_editor(
    default_brackets, 
    num_rows="dynamic", 
    use_container_width=True,
    hide_index=True
)

# Apply bracket logic
def assign_bracket(val, brackets_df):
    if pd.isna(val) or val <= 0:
        return "Sem Valor Definido", 0.0
        
    for _, row in brackets_df.iterrows():
        try:
            if row["Valor Mínimo (R$)"] <= val <= row["Valor Máximo (R$)"]:
                return row["Nome da Faixa"], float(row["Mensalidade Simulada (R$)"])
        except:
            continue
    return "Fora da Faixa", 0.0

# Calculate Simulated Values
filtered_df[["BRACKET_NAME", "SIMULATED_PAYMENT"]] = filtered_df.apply(
    lambda row: pd.Series(assign_bracket(row["EQUIPMENT_VALUE"], edited_brackets)), axis=1
)

filtered_df["DIFFERENCE"] = filtered_df["SIMULATED_PAYMENT"] - filtered_df["CURRENT_PAYMENT"]

# ---------------------------------------------------------------------------
# 4. Main Metrics
# ---------------------------------------------------------------------------
st.markdown("---")
st.subheader("Indicadores Principais")

col1, col2 = st.columns(2)

total_trucks = filtered_df["LICENSE_PLATE"].nunique()
total_with_fipe = len(filtered_df[filtered_df["EQUIPMENT_VALUE"] > 0])
avg_equip_val = filtered_df[filtered_df["EQUIPMENT_VALUE"] > 0]["EQUIPMENT_VALUE"].mean() if total_with_fipe > 0 else 0

total_current = filtered_df["CURRENT_PAYMENT"].sum()
total_simulated = filtered_df["SIMULATED_PAYMENT"].sum()
diff_total = total_simulated - total_current

col1.metric("Veículos Únicos", f"{total_trucks:,}")
col2.metric("Média Valor FIPE", f"R$ {avg_equip_val:,.2f}")


# ---------------------------------------------------------------------------
# 5. Visual Analysis
# ---------------------------------------------------------------------------
st.markdown("---")
st.subheader("Análise Visual")

# Chart 2: Output by Bracket
bracket_totals = filtered_df.groupby("BRACKET_NAME").agg(
    COUNT=("LICENSE_PLATE", "count"),
    SIMULATED_REVENUE=("SIMULATED_PAYMENT", "sum")
).reset_index()

fig2 = px.pie(
    bracket_totals, values="COUNT", names="BRACKET_NAME", hole=0.4,
    title="Distribuição de Veículos por Faixa de Valor"
)
fig2.update_traces(
    texttemplate="%{label}<br>%{value} veic. (%{percent})",
    textposition="inside"
)
fig2.update_layout(hiddenlabels=["Sem Valor Definido"])
st.plotly_chart(fig2, use_container_width=True)

# Chart FAP — Faixas por Mês (contagem absoluta + % por faixa)
BUCKET_ORDER = ["Até R$ 200k", "R$ 200k a R$ 300k", "R$ 300k a R$ 450k", "R$ 450k a R$ 600k", "Acima de R$ 600k", "Sem Valor Definido"]

def hide_sem_valor_default(fig):
    for tr in fig.data:
        if getattr(tr, "name", None) == "Sem Valor Definido":
            tr.visible = "legendonly"
    return fig

if "FUNDO" in filtered_df.columns:
    st.markdown("---")
    st.subheader("FAP — Faixas de Valor por Mês")

    df_fap = filtered_df[filtered_df["FUNDO"] == "FAP"]
    if not df_fap.empty:
        fap_faixa = df_fap.groupby(["MONTH", "EQUIP_VAL_BUCKET"])["LICENSE_PLATE"].count().reset_index()
        fap_faixa.columns = ["MONTH", "EQUIP_VAL_BUCKET", "COUNT"]
        month_tot = fap_faixa.groupby("MONTH")["COUNT"].transform("sum")
        fap_faixa["PCT"] = (fap_faixa["COUNT"] / month_tot * 100).round(1)
        fap_faixa["LABEL"] = fap_faixa["COUNT"].apply(lambda v: f"{int(v)} veíc.")

        fig_fap_faixas = px.bar(
            fap_faixa, x="MONTH", y="COUNT", color="EQUIP_VAL_BUCKET",
            barmode="stack",
            category_orders={"EQUIP_VAL_BUCKET": BUCKET_ORDER, "MONTH": all_months_sorted},
            title="FAP — Quantidade de Veículos por Faixa e Mês",
            labels={"COUNT": "Nº de Veículos", "MONTH": "Mês", "EQUIP_VAL_BUCKET": "Faixa de Valor"},
            text="LABEL",
        )
        fig_fap_faixas.update_traces(textposition="inside", insidetextanchor="middle")
        fig_fap_faixas.update_layout(legend_title_text="Faixa", yaxis_title="Nº de Veículos")
        hide_sem_valor_default(fig_fap_faixas)
        st.plotly_chart(fig_fap_faixas, use_container_width=True)

# Charts 3–6: Composição + Top 10 por Fundo — FAP primeiro, depois DPA

if "FUNDO" in filtered_df.columns:
    filtered_df["MARCA_MODELO"] = filtered_df["EQUIPMENT_BRAND"].str.strip() + " — " + filtered_df["EQUIPMENT_MODEL"].str.strip()

    for fundo_name in ["FAP", "DPA"]:
        st.markdown("---")
        st.subheader(f"📊 {fundo_name}")

        df_fundo = filtered_df[filtered_df["FUNDO"] == fundo_name]

        if df_fundo.empty:
            st.warning(f"Sem dados para {fundo_name}.")
            continue

        # Composição por faixa de valor — largura total
        comp = df_fundo.groupby(["MONTH", "EQUIP_VAL_BUCKET"])["LICENSE_PLATE"].count().reset_index()
        comp.columns = ["MONTH", "EQUIP_VAL_BUCKET", "COUNT"]
        month_totals = comp.groupby("MONTH")["COUNT"].transform("sum")
        comp["PCT"] = (comp["COUNT"] / month_totals * 100).round(1)
        comp["LABEL"] = comp["PCT"].apply(lambda v: f"{v:.1f}%")

        fig_comp = px.bar(
            comp, x="MONTH", y="PCT", color="EQUIP_VAL_BUCKET",
            barmode="stack",
            category_orders={"EQUIP_VAL_BUCKET": BUCKET_ORDER, "MONTH": all_months_sorted},
            title=f"Composição por Faixa de Valor — {fundo_name}",
            labels={"PCT": "% de Veículos", "MONTH": "Mês", "EQUIP_VAL_BUCKET": "Faixa de Valor"},
            text="LABEL",
        )
        fig_comp.update_traces(textposition="inside", insidetextanchor="middle")
        fig_comp.update_layout(yaxis_ticksuffix="%", legend_title_text="Faixa")
        hide_sem_valor_default(fig_comp)
        st.plotly_chart(fig_comp, use_container_width=True)

        if fundo_name == "DPA":
            comp["COUNT_LABEL"] = comp["COUNT"].apply(lambda v: f"{int(v)} veíc.")
            fig_dpa_qtd = px.bar(
                comp, x="MONTH", y="COUNT", color="EQUIP_VAL_BUCKET",
                barmode="stack",
                category_orders={"EQUIP_VAL_BUCKET": BUCKET_ORDER, "MONTH": all_months_sorted},
                title="DPA — Quantidade de Veículos por Faixa e Mês",
                labels={"COUNT": "Nº de Veículos", "MONTH": "Mês", "EQUIP_VAL_BUCKET": "Faixa de Valor"},
                text="COUNT_LABEL",
            )
            fig_dpa_qtd.update_traces(textposition="inside", insidetextanchor="middle")
            fig_dpa_qtd.update_layout(legend_title_text="Faixa", yaxis_title="Nº de Veículos")
            hide_sem_valor_default(fig_dpa_qtd)
            st.plotly_chart(fig_dpa_qtd, use_container_width=True)

        # Top 10 por faixa — filtros de faixa e mês
        available_buckets = [b for b in BUCKET_ORDER if b in df_fundo["EQUIP_VAL_BUCKET"].unique()]
        available_months = ["Acumulado"] + sorted(df_fundo["MONTH"].unique().tolist(), key=_month_sort_key)

        fc1, fc2 = st.columns(2)
        selected_bucket = fc1.selectbox(
            f"Faixa de Valor — {fundo_name}",
            options=available_buckets,
            key=f"bucket_select_{fundo_name}"
        )
        selected_month_top = fc2.selectbox(
            f"Mês — {fundo_name}",
            options=available_months,
            key=f"month_select_{fundo_name}"
        )

        df_bucket = df_fundo[df_fundo["EQUIP_VAL_BUCKET"] == selected_bucket]
        if selected_month_top != "Acumulado":
            df_bucket = df_bucket[df_bucket["MONTH"] == selected_month_top]

        total_bucket = df_bucket["LICENSE_PLATE"].nunique()

        if total_bucket == 0:
            st.warning(f"Sem dados para {fundo_name} na faixa/mês selecionados.")
            continue

        top10 = (
            df_bucket.groupby("MARCA_MODELO")["LICENSE_PLATE"]
            .nunique()
            .reset_index()
            .rename(columns={"LICENSE_PLATE": "COUNT"})
            .sort_values("COUNT", ascending=False)
            .head(10)
        )
        top10["PCT"] = (top10["COUNT"] / total_bucket * 100).round(1)
        top10 = top10.sort_values("PCT", ascending=True)
        top10["LABEL"] = top10.apply(lambda r: f"{r['PCT']:.1f}% | {int(r['COUNT'])} veíc.", axis=1)

        periodo = selected_month_top if selected_month_top != "Acumulado" else "Acumulado"
        fig_top = px.bar(
            top10, x="PCT", y="MARCA_MODELO", orientation="h",
            title=f"Top 10 Marcas/Equipamentos — {fundo_name} / {selected_bucket} / {periodo}",
            labels={"PCT": "% de Veículos", "MARCA_MODELO": "Marca / Modelo"},
            text="LABEL",
        )
        fig_top.update_traces(textposition="inside", insidetextanchor="middle")
        fig_top.update_layout(xaxis_ticksuffix="%", yaxis_title="", margin={"l": 10})
        st.plotly_chart(fig_top, use_container_width=True)

else:
    st.warning("Coluna FUNDO não encontrada — gráficos de composição DPA/FAP indisponíveis.")

# ---------------------------------------------------------------------------
# 7. Detailed Analysis Table & Export
# ---------------------------------------------------------------------------
st.markdown("---")
st.subheader("Tabela Analítica (Detalhada)")

display_cols = [
    "MONTH", "LICENSE_PLATE", "EQUIPMENT_TYPE", "EQUIPMENT_BRAND", "EQUIPMENT_MODEL", "EQUIPMENT_VALUE", 
    "CURRENT_PAYMENT", "BRACKET_NAME", "SIMULATED_PAYMENT", "DIFFERENCE"
]
st.dataframe(
    filtered_df[display_cols].style.format({
        "EQUIPMENT_VALUE": "R$ {:.2f}",
        "CURRENT_PAYMENT": "R$ {:.2f}",
        "SIMULATED_PAYMENT": "R$ {:.2f}",
        "DIFFERENCE": "R$ {:.2f}"
    }), 
    use_container_width=True,
    height=400
)

# Export button
csv_export = filtered_df.to_csv(index=False).encode('utf-8')
st.download_button(
    label="⬇️ Baixar Tabela de Simulação (CSV)",
    data=csv_export,
    file_name="afoocop_simulacao.csv",
    mime="text/csv",
)

# ---------------------------------------------------------------------------
# 7. Insights Summary
# ---------------------------------------------------------------------------
st.markdown("---")
st.subheader("Resumo de Insights")

benefited = len(filtered_df[filtered_df["DIFFERENCE"] < 0])
penalized = len(filtered_df[filtered_df["DIFFERENCE"] > 0])

st.info(f"""
💡 **Resumo do Cenário Simulado:**
- Sob a sua nova regra de faixas, a arrecadação total passaria de **R$ {total_current:,.2f}** para **R$ {total_simulated:,.2f}**.
- Isso representa uma variação de **{((total_simulated - total_current) / total_current) * 100:.1f}%** no caixa do agrupamento.
- **{benefited}** cobranças mensais ficariam **mais baratas** (motoristas que economizam).
- **{penalized}** cobranças mensais ficariam **mais caras** (veículos de maior valor que passam a pagar mais).
- Existem **{len(filtered_df[filtered_df['BRACKET_NAME'] == 'Sem Valor Definido'])}** veículos sem o `Valor Equipamento` cadastrado na base, portanto, foram cobrados R$ 0,00 na simulação.
""")

# ---------------------------------------------------------------------------
# 8. Fleet Age Analysis
# ---------------------------------------------------------------------------
st.markdown("---")
st.subheader("Idade da Frota e Ativadores do Seguro")

age_base = (
    filtered_df.sort_values("MONTH")
    .groupby("LICENSE_PLATE", as_index=False)
    .agg(
        FLEET_AGE=("FLEET_AGE", "first"),
        AGE_BUCKET=("AGE_BUCKET", "first"),
        CURRENT_PAYMENT_TOTAL=("CURRENT_PAYMENT", "sum"),
        ENTRY_TYPE=("ENTRY_TYPE", "first"),
        EQUIPMENT_BRAND=("EQUIPMENT_BRAND", "first"),
        EQUIPMENT_MODEL=("EQUIPMENT_MODEL", "first"),
    )
)

age_valid = age_base[age_base["FLEET_AGE"].notna()].copy()

if age_valid.empty:
    st.warning("Sem dados de ano para calcular a idade da frota.")
else:
    age_order = ["0 a 2 anos", "3 a 5 anos", "6 a 8 anos", "9 a 12 anos", "13+ anos", "Não informado"]
    avg_fleet_age = age_valid["FLEET_AGE"].mean()

    m1, m2 = st.columns(2)
    m1.metric("Idade média da frota", f"{avg_fleet_age:.1f} anos")
    m2.metric("Veículos com idade calculada", f"{age_valid['LICENSE_PLATE'].nunique():,}")

    dist_age = age_base.groupby("AGE_BUCKET", as_index=False)["LICENSE_PLATE"].count().rename(
        columns={"LICENSE_PLATE": "COUNT"}
    )
    fig_age_dist = px.bar(
        dist_age,
        x="AGE_BUCKET",
        y="COUNT",
        category_orders={"AGE_BUCKET": age_order},
        title="Distribuição da Frota por Faixa de Idade",
        labels={"AGE_BUCKET": "Faixa de Idade", "COUNT": "Nº de Veículos"},
        text="COUNT",
    )
    fig_age_dist.update_traces(textposition="outside")
    fig_age_dist.update_layout(xaxis_title="Faixa de Idade", yaxis_title="Nº de Veículos")
    st.plotly_chart(fig_age_dist, use_container_width=True)

    ativadores = (
        filtered_df[filtered_df["FLEET_AGE"].notna()]
        .groupby(["AGE_BUCKET", "ENTRY_TYPE"], as_index=False)["CURRENT_PAYMENT"]
        .sum()
    )
    top_ativadores = (
        ativadores.sort_values(["AGE_BUCKET", "CURRENT_PAYMENT"], ascending=[True, False])
        .groupby("AGE_BUCKET", as_index=False)
        .head(3)
    )
    entry_totals = top_ativadores.groupby("ENTRY_TYPE")["CURRENT_PAYMENT"].transform("sum")
    top_ativadores["PCT"] = (top_ativadores["CURRENT_PAYMENT"] / entry_totals * 100).round(1)
    top_ativadores["LABEL"] = top_ativadores["PCT"].apply(lambda v: f"{v:.1f}%")

    fig_ativadores = px.bar(
        top_ativadores,
        x="PCT",
        y="ENTRY_TYPE",
        color="AGE_BUCKET",
        orientation="h",
        category_orders={"AGE_BUCKET": age_order},
        title="Maiores Ativadores do Seguro por Faixa de Idade da Frota (Top 3)",
        labels={
            "PCT": "% do Total",
            "ENTRY_TYPE": "Ativador",
            "AGE_BUCKET": "Faixa de Idade",
        },
        text="LABEL",
    )
    fig_ativadores.update_traces(textposition="inside", insidetextanchor="middle")
    fig_ativadores.update_layout(yaxis_title="Ativador", xaxis_ticksuffix="%", xaxis_title="% do Total")
    st.plotly_chart(fig_ativadores, use_container_width=True)


