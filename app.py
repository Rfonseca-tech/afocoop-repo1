import os
import uuid
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

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

    col_map = {
        "MONTH_LABEL": "MONTH",
        "MONTH": "MONTH_DATE",
        "PLACA": "LICENSE_PLATE",
        "Valor Equipamento": "EQUIPMENT_VALUE",
        "Cavalo/Carreta": "EQUIPMENT_TYPE",
        "Marca": "EQUIPMENT_BRAND",
        "Modelo": "EQUIPMENT_MODEL",
        "Ano Fabricação": "MANUFACTURE_YEAR",
        "Ano Modelo": "MODEL_YEAR",
        "TIPO_LANCAMENTO": "ENTRY_TYPE",
        "VALOR": "TRANSACTION_AMOUNT",
    }

    if "MONTH" in df_raw.columns:
        df_raw.rename(columns={"MONTH": "MONTH_DATE"}, inplace=True)
    if "MONTH_LABEL" in df_raw.columns:
        df_raw.rename(columns={"MONTH_LABEL": "MONTH"}, inplace=True)

    for excel_col, app_col in col_map.items():
        if excel_col in ["MONTH", "MONTH_LABEL"]:
            continue
        if excel_col in df_raw.columns:
            df_raw.rename(columns={excel_col: app_col}, inplace=True)

    req_cols = [
        "MONTH",
        "LICENSE_PLATE",
        "TRANSACTION_AMOUNT",
        "EQUIPMENT_VALUE",
        "EQUIPMENT_TYPE",
        "EQUIPMENT_BRAND",
        "EQUIPMENT_MODEL",
        "MANUFACTURE_YEAR",
        "MODEL_YEAR",
        "ENTRY_TYPE",
    ]
    for c in req_cols:
        if c not in df_raw.columns:
            df_raw[c] = np.nan if c in ["EQUIPMENT_VALUE", "TRANSACTION_AMOUNT", "MANUFACTURE_YEAR", "MODEL_YEAR"] else "Unknown"

    df_raw["EQUIPMENT_VALUE"] = pd.to_numeric(df_raw["EQUIPMENT_VALUE"], errors="coerce").fillna(0)
    df_raw["TRANSACTION_AMOUNT"] = pd.to_numeric(df_raw["TRANSACTION_AMOUNT"], errors="coerce").fillna(0)
    df_raw["MANUFACTURE_YEAR"] = pd.to_numeric(df_raw["MANUFACTURE_YEAR"], errors="coerce")
    df_raw["MODEL_YEAR"] = pd.to_numeric(df_raw["MODEL_YEAR"], errors="coerce")

    df_raw["EQUIPMENT_TYPE"] = df_raw["EQUIPMENT_TYPE"].fillna("Desconhecido")
    df_raw["EQUIPMENT_BRAND"] = df_raw["EQUIPMENT_BRAND"].fillna("Desconhecido")
    df_raw["EQUIPMENT_MODEL"] = df_raw["EQUIPMENT_MODEL"].fillna("Desconhecido")
    df_raw["ENTRY_TYPE"] = df_raw["ENTRY_TYPE"].fillna("Desconhecido")

    agg_funcs = {
        "TRANSACTION_AMOUNT": "sum",
        "EQUIPMENT_VALUE": "first",
        "EQUIPMENT_TYPE": "first",
        "EQUIPMENT_BRAND": "first",
        "EQUIPMENT_MODEL": "first",
        "MANUFACTURE_YEAR": "first",
        "MODEL_YEAR": "first",
        "ENTRY_TYPE": "first",
    }

    if "FUNDO" in df_raw.columns:
        agg_funcs["FUNDO"] = "first"

    grouped = df_raw.groupby(["MONTH", "LICENSE_PLATE"], as_index=False).agg(agg_funcs)
    grouped.rename(columns={"TRANSACTION_AMOUNT": "CURRENT_PAYMENT"}, inplace=True)

    current_year = datetime.now().year
    grouped["FLEET_YEAR_BASE"] = grouped["MODEL_YEAR"].where(
        grouped["MODEL_YEAR"].notna() & (grouped["MODEL_YEAR"] > 0),
        grouped["MANUFACTURE_YEAR"],
    )
    grouped["FLEET_AGE"] = current_year - grouped["FLEET_YEAR_BASE"]
    grouped.loc[(grouped["FLEET_AGE"] < 0) | (grouped["FLEET_AGE"] > 80), "FLEET_AGE"] = np.nan

    def get_age_bucket(age):
        if pd.isna(age):
            return "Não informado"
        if age <= 2:
            return "0 a 2 anos"
        if age <= 5:
            return "3 a 5 anos"
        if age <= 8:
            return "6 a 8 anos"
        if age <= 12:
            return "9 a 12 anos"
        return "13+ anos"

    grouped["AGE_BUCKET"] = grouped["FLEET_AGE"].apply(get_age_bucket)

    return grouped, None


# ---------------------------------------------------------------------------
# 2. Helpers for Config-Driven Ranges/Filters
# ---------------------------------------------------------------------------
_MONTH_PT = {
    "Janeiro": 1,
    "Fevereiro": 2,
    "Março": 3,
    "Abril": 4,
    "Maio": 5,
    "Junho": 6,
    "Julho": 7,
    "Agosto": 8,
    "Setembro": 9,
    "Outubro": 10,
    "Novembro": 11,
    "Dezembro": 12,
}


def _month_sort_key(label):
    try:
        nome, ano = str(label).split("/")
        return (int(ano), _MONTH_PT.get(nome.strip(), 0))
    except Exception:
        return (9999, 0)


def _new_id():
    return str(uuid.uuid4())


def _parse_optional_float(value):
    if value is None:
        return None
    txt = str(value).strip()
    if txt == "":
        return None
    txt = txt.replace("R$", "").replace(" ", "")
    if "," in txt and "." in txt:
        txt = txt.replace(".", "").replace(",", ".")
    elif "," in txt:
        txt = txt.replace(",", ".")
    try:
        return float(txt)
    except Exception:
        return None


def _format_optional_float(value):
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    return f"{float(value):g}"


def _range_row(label, start, end, monthly):
    return {
        "id": _new_id(),
        "label": label,
        "start": start,
        "end": end,
        "monthly": monthly,
    }


def _assign_range(value, active_ranges):
    if pd.isna(value) or value <= 0:
        return "Sem Valor Definido", 0.0
    for r in active_ranges:
        start = r["start"]
        end = r["end"]
        if start is None:
            continue
        if end is None:
            if value >= start:
                return r["label"], float(r["monthly"])
        elif start <= value <= end:
            return r["label"], float(r["monthly"])
    return "Fora da Faixa", 0.0


def _new_filter_row(field="MONTH", operator="in", selected=None, query=""):
    return {
        "id": _new_id(),
        "field": field,
        "operator": operator,
        "selected": selected,
        "query": query,
    }


DEFAULT_RANGE_CONFIG = [
    _range_row("Até R$ 200k", 0.0, 200000.0, 80.0),
    _range_row("R$ 200k a R$ 300k", 200000.01, 300000.0, 120.0),
    _range_row("R$ 300k a R$ 450k", 300000.01, 450000.0, 180.0),
    _range_row("R$ 450k a R$ 600k", 450000.01, 600000.0, 250.0),
    _range_row("Acima de R$ 600k", 600000.01, None, 350.0),
]

DEFAULT_FILTER_TEMPLATES = [
    {"field": "MONTH", "operator": "in"},
    {"field": "EQUIPMENT_TYPE", "operator": "in"},
    {"field": "EQUIPMENT_BRAND", "operator": "in"},
    {"field": "EQUIP_VAL_BUCKET", "operator": "in_default_bucket"},
    {"field": "LICENSE_PLATE", "operator": "contains"},
]


# ---------------------------------------------------------------------------
# 3. App Layout + Data Load
# ---------------------------------------------------------------------------
st.title("🚚 AFOOCOP: Relatório e Simulador de Rateio")
st.markdown("Analise o custo compartilhado atual e simule novos cenários de cobrança com base em faixas de valor do equipamento.")

st.sidebar.header("1. Base de Dados")
current_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(current_dir, "AFOOCOP_Rateios_Consolidado.xlsx")

error = None
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

all_months_sorted = sorted(df["MONTH"].unique().tolist(), key=_month_sort_key)


# ---------------------------------------------------------------------------
# 3.5 Boleto Inputs (Manual + Formula)
# ---------------------------------------------------------------------------
st.markdown("---")
tab_boleto_fap = st.tabs(["Boleto - FAP"])[0]

if "boleto_fap_mensal" not in st.session_state:
    st.session_state.boleto_fap_mensal = pd.DataFrame(
        {
            "Mês": all_months_sorted,
            "Quantidade de Cavalos": [0 for _ in all_months_sorted],
            "Valor Total do Rateio": [0.0 for _ in all_months_sorted],
            "Participação": [0.0 for _ in all_months_sorted],
        }
    )
else:
    previous_boleto = st.session_state.boleto_fap_mensal.copy()
    if "Mês" in previous_boleto.columns:
        previous_map = previous_boleto.set_index("Mês")
        synced_rows = []
        for month in all_months_sorted:
            if month in previous_map.index:
                row = previous_map.loc[month]
                synced_rows.append(
                    {
                        "Mês": month,
                        "Quantidade de Cavalos": int(pd.to_numeric(row.get("Quantidade de Cavalos", 0), errors="coerce") or 0),
                        "Valor Total do Rateio": float(pd.to_numeric(row.get("Valor Total do Rateio", 0.0), errors="coerce") or 0.0),
                        "Participação": float(pd.to_numeric(row.get("Participação", 0.0), errors="coerce") or 0.0),
                    }
                )
            else:
                synced_rows.append(
                    {
                        "Mês": month,
                        "Quantidade de Cavalos": 0,
                        "Valor Total do Rateio": 0.0,
                        "Participação": 0.0,
                    }
                )
        st.session_state.boleto_fap_mensal = pd.DataFrame(synced_rows)

with tab_boleto_fap:
    st.subheader("Boleto - FAP — Por Mês")
    st.caption("Preencha manualmente as linhas 3, 4 e 5 por mês. A linha 6 é calculada automaticamente em cada mês.")

    boleto_df = st.session_state.boleto_fap_mensal.copy()
    boleto_df["Valor do Boleto"] = np.where(
        boleto_df["Quantidade de Cavalos"] > 0,
        (boleto_df["Valor Total do Rateio"] + boleto_df["Participação"]) / boleto_df["Quantidade de Cavalos"],
        0.0,
    )

    edited_boleto_df = st.data_editor(
        boleto_df,
        hide_index=True,
        use_container_width=True,
        disabled=["Mês", "Valor do Boleto"],
        column_config={
            "Mês": st.column_config.TextColumn("Mês"),
            "Quantidade de Cavalos": st.column_config.NumberColumn("3 - Quantidade de Cavalos", min_value=0, step=1, format="%d"),
            "Valor Total do Rateio": st.column_config.NumberColumn("4 - Valor Total do Rateio", format="R$ %.2f", step=100.0),
            "Participação": st.column_config.NumberColumn("5 - Participação", format="R$ %.2f", step=100.0),
            "Valor do Boleto": st.column_config.NumberColumn("6 - Valor do Boleto (Automático)", format="R$ %.2f"),
        },
        key="boleto_fap_editor",
    )

    st.session_state.boleto_fap_mensal = edited_boleto_df[
        ["Mês", "Quantidade de Cavalos", "Valor Total do Rateio", "Participação"]
    ].copy()


# ---------------------------------------------------------------------------
# 4. Sidebar: Configurable Ranges
# ---------------------------------------------------------------------------
st.sidebar.header("2. Faixas de Cobrança")
st.sidebar.caption("Cada linha possui início/fim editáveis, '+' para inserir nova faixa e '🗑️' para remover.")

if "range_config" not in st.session_state:
    st.session_state.range_config = [r.copy() for r in DEFAULT_RANGE_CONFIG]

if not st.session_state.range_config:
    st.session_state.range_config = [DEFAULT_RANGE_CONFIG[0].copy()]

range_rows = st.session_state.range_config
pending_range_add_after = None
pending_range_delete_idx = None

for idx, row in enumerate(range_rows):
    rid = row["id"]
    top1, top2 = st.sidebar.columns([2.2, 1.1])
    row["label"] = top1.text_input(
        "Nome",
        value=row.get("label", f"Faixa {idx + 1}"),
        key=f"range_label_{rid}",
        label_visibility="collapsed",
    )
    monthly_txt = top2.text_input(
        "Mensalidade",
        value=_format_optional_float(row.get("monthly")),
        key=f"range_monthly_{rid}",
        label_visibility="collapsed",
    )

    bot1, bot2, bot3, bot4 = st.sidebar.columns([1.2, 1.2, 0.6, 0.6])
    start_txt = bot1.text_input(
        "Início",
        value=_format_optional_float(row.get("start")),
        key=f"range_start_{rid}",
        label_visibility="collapsed",
    )
    end_txt = bot2.text_input(
        "Fim",
        value=_format_optional_float(row.get("end")),
        key=f"range_end_{rid}",
        placeholder="vazio",
        label_visibility="collapsed",
    )

    if bot3.button("+", key=f"range_add_{rid}", help="Inserir faixa abaixo"):
        pending_range_add_after = idx
    if bot4.button("🗑️", key=f"range_del_{rid}", help="Remover faixa"):
        pending_range_delete_idx = idx

    st.sidebar.markdown("<div style='height: 6px;'></div>", unsafe_allow_html=True)

    row["start"] = _parse_optional_float(start_txt)
    row["end"] = _parse_optional_float(end_txt)
    parsed_monthly = _parse_optional_float(monthly_txt)
    row["monthly"] = 0.0 if parsed_monthly is None else parsed_monthly

if pending_range_add_after is not None:
    prev_row = st.session_state.range_config[pending_range_add_after]
    st.session_state.range_config.insert(
        pending_range_add_after + 1,
        _range_row(
            label=f"Faixa {pending_range_add_after + 2}",
            start=prev_row.get("end"),
            end=None,
            monthly=prev_row.get("monthly", 0.0),
        ),
    )
    st.rerun()

if pending_range_delete_idx is not None:
    if len(st.session_state.range_config) > 1:
        st.session_state.range_config.pop(pending_range_delete_idx)
        st.rerun()
    else:
        st.sidebar.warning("É necessário manter pelo menos uma faixa.")

normalized_ranges = sorted(
    [
        {
            "id": r.get("id", _new_id()),
            "label": (r.get("label") or "").strip() or "Faixa sem nome",
            "start": _parse_optional_float(r.get("start")),
            "end": _parse_optional_float(r.get("end")),
            "monthly": _parse_optional_float(r.get("monthly")) or 0.0,
        }
        for r in st.session_state.range_config
    ],
    key=lambda x: (float("inf") if x["start"] is None else x["start"]),
)

active_ranges = []
range_errors = []
for i, r in enumerate(normalized_ranges):
    if r["start"] is None:
        range_errors.append(f"Faixa {i + 1}: início inválido ou vazio.")
        continue
    if r["end"] is not None and r["end"] < r["start"]:
        range_errors.append(f"{r['label']}: fim menor que início.")
        continue
    active_ranges.append(r)

if range_errors:
    st.sidebar.warning("Configuração de faixas com inconsistências:\n- " + "\n- ".join(range_errors))

# Apply configured ranges globally to keep all app outputs in sync
df = df.copy()
df[["EQUIP_VAL_BUCKET", "SIMULATED_PAYMENT_BASE"]] = df.apply(
    lambda row: pd.Series(_assign_range(row["EQUIPMENT_VALUE"], active_ranges)),
    axis=1,
)


# ---------------------------------------------------------------------------
# 5. Sidebar: Configurable Filters
# ---------------------------------------------------------------------------
st.sidebar.header("3. Filtros")
st.sidebar.caption("Filtros configuráveis por linha. Você pode adicionar quantos filtros quiser.")

filterable_fields = []
for c in df.columns:
    if c in [
        "MONTH",
        "EQUIPMENT_TYPE",
        "EQUIPMENT_BRAND",
        "EQUIP_VAL_BUCKET",
        "LICENSE_PLATE",
        "FUNDO",
        "ENTRY_TYPE",
        "EQUIPMENT_MODEL",
    ]:
        filterable_fields.append(c)
filterable_fields = list(dict.fromkeys(filterable_fields))

if "filter_config" not in st.session_state:
    initial_filters = []
    for tpl in DEFAULT_FILTER_TEMPLATES:
        field = tpl["field"]
        op = tpl["operator"]
        if op == "in_default_bucket":
            options = sorted(df["EQUIP_VAL_BUCKET"].dropna().unique().tolist())
            selected = [b for b in options if b != "Sem Valor Definido"]
            initial_filters.append(_new_filter_row(field="EQUIP_VAL_BUCKET", operator="in", selected=selected))
        elif op == "in":
            options = sorted(df[field].dropna().unique().tolist(), key=_month_sort_key if field == "MONTH" else None)
            initial_filters.append(_new_filter_row(field=field, operator="in", selected=options))
        else:
            initial_filters.append(_new_filter_row(field=field, operator="contains", query=""))
    st.session_state.filter_config = initial_filters

if not st.session_state.filter_config:
    st.session_state.filter_config = [_new_filter_row()]

pending_filter_add_after = None
pending_filter_delete_idx = None

for idx, row in enumerate(st.session_state.filter_config):
    fid = row["id"]
    fx1, fx2, fx3, fx4 = st.sidebar.columns([1.2, 0.9, 0.45, 0.45])

    field = fx1.selectbox(
        "Campo",
        options=filterable_fields,
        index=filterable_fields.index(row.get("field")) if row.get("field") in filterable_fields else 0,
        key=f"filter_field_{fid}",
        label_visibility="collapsed",
    )

    op_options = ["in", "contains"] if field == "LICENSE_PLATE" else ["in", "not in"]
    operator = fx2.selectbox(
        "Operador",
        options=op_options,
        index=op_options.index(row.get("operator")) if row.get("operator") in op_options else 0,
        key=f"filter_op_{fid}",
        label_visibility="collapsed",
    )

    if fx3.button("+", key=f"filter_add_{fid}", help="Inserir filtro abaixo"):
        pending_filter_add_after = idx
    if fx4.button("🗑️", key=f"filter_del_{fid}", help="Remover filtro"):
        pending_filter_delete_idx = idx

    row["field"] = field
    row["operator"] = operator

    if operator == "contains":
        row["query"] = st.sidebar.text_input(
            f"Texto ({field})",
            value=row.get("query", ""),
            key=f"filter_query_{fid}",
            label_visibility="collapsed",
        )
        row["selected"] = None
    else:
        options = sorted(df[field].dropna().unique().tolist(), key=_month_sort_key if field == "MONTH" else None)
        previous_selected = row.get("selected")
        widget_key = f"filter_values_{fid}"
        force_widget_sync = False
        if previous_selected is None:
            current_selected = options
        else:
            current_selected = [x for x in previous_selected if x in options]

            if field == "EQUIP_VAL_BUCKET":
                # If a faixa label was renamed, previous selection contains stale labels.
                # In this case, auto-include newly available labels to avoid data disappearing.
                stale_labels = [x for x in previous_selected if x not in options]
                if stale_labels:
                    default_bucket_options = [b for b in options if b != "Sem Valor Definido"]
                    current_selected = sorted(set(current_selected + default_bucket_options))
                    force_widget_sync = True

        if field == "EQUIP_VAL_BUCKET" and not current_selected:
            # Keep legacy default behavior when selection is empty.
            current_selected = [b for b in options if b != "Sem Valor Definido"]
            force_widget_sync = True

        # Streamlit keeps widget state by key and may ignore changed defaults.
        # When faixa labels change, we must sync the widget state explicitly.
        if force_widget_sync:
            st.session_state[widget_key] = current_selected

        row["selected"] = st.sidebar.multiselect(
            f"Valores ({field})",
            options=options,
            default=current_selected,
            key=widget_key,
            label_visibility="collapsed",
        )
        row["query"] = ""

if pending_filter_add_after is not None:
    st.session_state.filter_config.insert(pending_filter_add_after + 1, _new_filter_row(field="MONTH", operator="in"))
    st.rerun()

if pending_filter_delete_idx is not None:
    if len(st.session_state.filter_config) > 1:
        st.session_state.filter_config.pop(pending_filter_delete_idx)
        st.rerun()
    else:
        st.sidebar.warning("É necessário manter pelo menos um filtro.")

filtered_df = df.copy()
for f in st.session_state.filter_config:
    field = f.get("field")
    operator = f.get("operator")
    if field not in filtered_df.columns:
        continue

    if operator == "contains":
        query = (f.get("query") or "").strip()
        if query:
            filtered_df = filtered_df[filtered_df[field].astype(str).str.contains(query, case=False, na=False)]
    elif operator == "in":
        selected = f.get("selected")
        if selected is not None:
            filtered_df = filtered_df[filtered_df[field].isin(selected)]
    elif operator == "not in":
        selected = f.get("selected")
        if selected is not None:
            filtered_df = filtered_df[~filtered_df[field].isin(selected)]

if filtered_df.empty:
    st.warning("Nenhum dado encontrado para os filtros selecionados.")
    st.stop()


# ---------------------------------------------------------------------------
# 6. Simulated Values (from configured ranges)
# ---------------------------------------------------------------------------
filtered_df[["BRACKET_NAME", "SIMULATED_PAYMENT"]] = filtered_df.apply(
    lambda row: pd.Series(_assign_range(row["EQUIPMENT_VALUE"], active_ranges)),
    axis=1,
)
filtered_df["DIFFERENCE"] = filtered_df["SIMULATED_PAYMENT"] - filtered_df["CURRENT_PAYMENT"]


# ---------------------------------------------------------------------------
# 7. Main Metrics
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
# 8. Visual Analysis
# ---------------------------------------------------------------------------
st.markdown("---")
st.subheader("Análise Visual")

bracket_totals = filtered_df.groupby("BRACKET_NAME").agg(
    COUNT=("LICENSE_PLATE", "count"),
    SIMULATED_REVENUE=("SIMULATED_PAYMENT", "sum"),
).reset_index()

fig2 = px.pie(
    bracket_totals,
    values="COUNT",
    names="BRACKET_NAME",
    hole=0.4,
    title="Distribuição de Veículos por Faixa de Valor",
)
fig2.update_traces(texttemplate="%{label}<br>%{value} veic. (%{percent})", textposition="inside")
fig2.update_layout(hiddenlabels=["Sem Valor Definido"])
st.plotly_chart(fig2, use_container_width=True)

BUCKET_ORDER = [r["label"] for r in active_ranges] + ["Sem Valor Definido", "Fora da Faixa"]


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
            fap_faixa,
            x="MONTH",
            y="COUNT",
            color="EQUIP_VAL_BUCKET",
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

if "FUNDO" in filtered_df.columns:
    filtered_df["MARCA_MODELO"] = filtered_df["EQUIPMENT_BRAND"].str.strip() + " — " + filtered_df["EQUIPMENT_MODEL"].str.strip()

    for fundo_name in ["FAP", "DPA"]:
        st.markdown("---")
        st.subheader(f"📊 {fundo_name}")

        df_fundo = filtered_df[filtered_df["FUNDO"] == fundo_name]

        if df_fundo.empty:
            st.warning(f"Sem dados para {fundo_name}.")
            continue

        comp = df_fundo.groupby(["MONTH", "EQUIP_VAL_BUCKET"])["LICENSE_PLATE"].count().reset_index()
        comp.columns = ["MONTH", "EQUIP_VAL_BUCKET", "COUNT"]
        month_totals = comp.groupby("MONTH")["COUNT"].transform("sum")
        comp["PCT"] = (comp["COUNT"] / month_totals * 100).round(1)
        comp["LABEL"] = comp["PCT"].apply(lambda v: f"{v:.1f}%")

        fig_comp = px.bar(
            comp,
            x="MONTH",
            y="PCT",
            color="EQUIP_VAL_BUCKET",
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
                comp,
                x="MONTH",
                y="COUNT",
                color="EQUIP_VAL_BUCKET",
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

        available_buckets = [b for b in BUCKET_ORDER if b in df_fundo["EQUIP_VAL_BUCKET"].unique()]
        available_months = ["Acumulado"] + sorted(df_fundo["MONTH"].unique().tolist(), key=_month_sort_key)

        fc1, fc2 = st.columns(2)
        selected_bucket = fc1.selectbox(
            f"Faixa de Valor — {fundo_name}",
            options=available_buckets,
            key=f"bucket_select_{fundo_name}",
        )
        selected_month_top = fc2.selectbox(
            f"Mês — {fundo_name}",
            options=available_months,
            key=f"month_select_{fundo_name}",
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
            top10,
            x="PCT",
            y="MARCA_MODELO",
            orientation="h",
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
# 9. Detailed Analysis Table & Export
# ---------------------------------------------------------------------------
st.markdown("---")
st.subheader("Tabela Analítica (Detalhada)")

display_cols = [
    "MONTH",
    "LICENSE_PLATE",
    "EQUIPMENT_TYPE",
    "EQUIPMENT_BRAND",
    "EQUIPMENT_MODEL",
    "EQUIPMENT_VALUE",
    "CURRENT_PAYMENT",
    "BRACKET_NAME",
    "SIMULATED_PAYMENT",
    "DIFFERENCE",
]
st.dataframe(
    filtered_df[display_cols].style.format(
        {
            "EQUIPMENT_VALUE": "R$ {:.2f}",
            "CURRENT_PAYMENT": "R$ {:.2f}",
            "SIMULATED_PAYMENT": "R$ {:.2f}",
            "DIFFERENCE": "R$ {:.2f}",
        }
    ),
    use_container_width=True,
    height=400,
)

csv_export = filtered_df.to_csv(index=False).encode("utf-8")
st.download_button(
    label="⬇️ Baixar Tabela de Simulação (CSV)",
    data=csv_export,
    file_name="afoocop_simulacao.csv",
    mime="text/csv",
)


# ---------------------------------------------------------------------------
# 10. Insights Summary
# ---------------------------------------------------------------------------
st.markdown("---")
st.subheader("Resumo de Insights")

benefited = len(filtered_df[filtered_df["DIFFERENCE"] < 0])
penalized = len(filtered_df[filtered_df["DIFFERENCE"] > 0])

pct_change = ((total_simulated - total_current) / total_current * 100) if total_current else 0

st.info(
    f"""
💡 **Resumo do Cenário Simulado:**
- Sob a sua nova regra de faixas, a arrecadação total passaria de **R$ {total_current:,.2f}** para **R$ {total_simulated:,.2f}**.
- Isso representa uma variação de **{pct_change:.1f}%** no caixa do agrupamento.
- **{benefited}** cobranças mensais ficariam **mais baratas** (motoristas que economizam).
- **{penalized}** cobranças mensais ficariam **mais caras** (veículos de maior valor que passam a pagar mais).
- Existem **{len(filtered_df[filtered_df['BRACKET_NAME'] == 'Sem Valor Definido'])}** veículos sem o `Valor Equipamento` cadastrado na base, portanto, foram cobrados R$ 0,00 na simulação.
"""
)


# ---------------------------------------------------------------------------
# 11. Fleet Age Analysis
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

    dist_age = age_base.groupby("AGE_BUCKET", as_index=False)["LICENSE_PLATE"].count().rename(columns={"LICENSE_PLATE": "COUNT"})
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
        labels={"PCT": "% do Total", "ENTRY_TYPE": "Ativador", "AGE_BUCKET": "Faixa de Idade"},
        text="LABEL",
    )
    fig_ativadores.update_traces(textposition="inside", insidetextanchor="middle")
    fig_ativadores.update_layout(yaxis_title="Ativador", xaxis_ticksuffix="%", xaxis_title="% do Total")
    st.plotly_chart(fig_ativadores, use_container_width=True)
