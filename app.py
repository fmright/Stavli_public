"""
Streamlit-дашборд для анализа эффективности ставок.

Главные вопросы:
1) Какие каналы прогнозов реально зарабатывают?
2) Какие типы ставок сливают?

Принимает Excel с листом 'Ставки' стандартной структуры (см. шаблон template_stavki.xlsx).
"""

import io
import pandas as pd
import plotly.express as px
import streamlit as st


# ============================================================
# НАСТРОЙКИ
# ============================================================
st.set_page_config(
    page_title="Ставки — аналитика ROI",
    page_icon="🎯",
    layout="wide",
)


# ============================================================
# ПАРОЛЬ
# ============================================================
def check_password() -> bool:
    """Простая защита паролем. Пароль хранится в st.secrets['app_password']."""
    def password_entered():
        if st.session_state.get("password") == st.secrets.get("app_password"):
            st.session_state["password_correct"] = True
            # Не храним сам пароль в session_state после проверки
            if "password" in st.session_state:
                del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if st.session_state.get("password_correct", False):
        return True

    st.text_input(
        "🔐 Пароль",
        type="password",
        on_change=password_entered,
        key="password",
    )
    if (
        "password_correct" in st.session_state
        and not st.session_state["password_correct"]
    ):
        st.error("Неверный пароль 😕")
    return False


if not check_password():
    st.stop()


# Минимальное число ставок для «надёжного» ROI — ниже подсвечиваем как шум
MIN_N_FOR_RELIABLE = 30

REQUIRED_COLUMNS = [
    "дата", "канал", "спорт", "лига", "матч", "тип_ставки",
    "направление", "ставка", "кэф", "результат", "профит",
]
OPTIONAL_COLUMNS = ["линия", "кэф_прогноза", "комиссия", "формат", "тип_записи", "заметки"]


# ============================================================
# ЗАГРУЗКА И ПОДГОТОВКА
# ============================================================
@st.cache_data
def load_data(file_bytes: bytes) -> pd.DataFrame:
    """Читает Excel, оставляет только строки с заполненной датой."""
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Ставки")
    df = df.dropna(subset=["дата"]).reset_index(drop=True)
    df["дата"] = pd.to_datetime(df["дата"])

    # Считаем профит сами (на случай если формула не пересчиталась)
    def calc_profit(row):
        if row["результат"] == "WIN":
            return row["ставка"] * (row["кэф"] - 1)
        if row["результат"] == "LOSE":
            return -row["ставка"]
        return 0  # RETURN, PENDING

    df["профит"] = df.apply(calc_profit, axis=1)
    return df


def validate_columns(df: pd.DataFrame):
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    return (len(missing) == 0), missing


def calc_roi(df: pd.DataFrame) -> float:
    stake = df["ставка"].sum()
    return df["профит"].sum() / stake * 100 if stake else 0.0


def calc_winrate(df: pd.DataFrame) -> float:
    decided = df[df["результат"].isin(["WIN", "LOSE"])]
    if len(decided) == 0:
        return 0.0
    return (decided["результат"] == "WIN").sum() / len(decided) * 100


def aggregate(df: pd.DataFrame, by) -> pd.DataFrame:
    by_list = [by] if isinstance(by, str) else by
    agg = df.groupby(by_list).agg(
        ставок=("профит", "count"),
        оборот=("ставка", "sum"),
        профит=("профит", "sum"),
    )
    agg["ROI%"] = (agg["профит"] / agg["оборот"] * 100).round(1)
    return agg.sort_values("ROI%", ascending=False).reset_index()


def reliability_marker(n: int) -> str:
    return "✅" if n >= MIN_N_FOR_RELIABLE else "⚠️"


# ============================================================
# UI: HEADER + UPLOAD
# ============================================================
st.title("🎯 Аналитика ставок")
st.caption(
    "Два главных вопроса: **какие каналы реально зарабатывают** "
    "и **какие типы ставок сливают**. Всё остальное — детали."
)

uploaded = st.file_uploader(
    "Файл со ставками (.xlsx по шаблону)",
    type=["xlsx"],
    help="Excel с листом 'Ставки' стандартной структуры. Шаблон — template_stavki.xlsx",
)

if not uploaded:
    st.info(
        "👆 Загрузи файл, чтобы увидеть аналитику. "
        "Используй шаблон `template_stavki.xlsx` — он гарантирует, что всё заработает."
    )
    with st.expander("📋 Что должно быть в файле"):
        st.markdown(
            f"**Лист:** `Ставки`\n\n"
            f"**Обязательные колонки:** {', '.join(f'`{c}`' for c in REQUIRED_COLUMNS)}\n\n"
            f"**Опциональные:** {', '.join(f'`{c}`' for c in OPTIONAL_COLUMNS)}"
        )
    st.stop()

df_all = load_data(uploaded.getvalue())
ok, missing = validate_columns(df_all)
if not ok:
    st.error(f"В файле не хватает обязательных колонок: {', '.join(missing)}")
    st.stop()

# ============================================================
# ФИЛЬТРЫ
# ============================================================
st.sidebar.header("🔍 Фильтры")

# Главный переключатель: показывать реальные ставки или сигналы или всё.
# Если в данных нет колонки тип_записи (старые файлы) — фильтр не показываем,
# считаем всё как BET.
if "тип_записи" in df_all.columns:
    rec_type_filter = st.sidebar.radio(
        "Что анализируем",
        options=["💰 Реальные ставки", "📡 Сигналы (без ставки)", "Всё вместе"],
        index=0,
        help=(
            "Реальные ставки — то, на что были поставлены деньги. "
            "Сигналы — отслеженные прогнозы, на которые не успели поставить."
        ),
    )
    if rec_type_filter == "💰 Реальные ставки":
        df_all = df_all[df_all["тип_записи"].fillna("BET") == "BET"]
    elif rec_type_filter == "📡 Сигналы (без ставки)":
        df_all = df_all[df_all["тип_записи"] == "SIGNAL"]
    # "Всё вместе" — фильтр не применяем
    st.sidebar.divider()

if len(df_all) == 0:
    st.warning("По выбранному типу записей данных нет. Переключи фильтр.")
    st.stop()

date_min, date_max = df_all["дата"].min().date(), df_all["дата"].max().date()
date_range = st.sidebar.date_input(
    "Период",
    value=(date_min, date_max),
    min_value=date_min,
    max_value=date_max,
)


def multi(label, col):
    options = sorted(df_all[col].dropna().unique().tolist())
    return st.sidebar.multiselect(label, options, default=[])


f_channels = multi("Канал", "канал")
f_sports = multi("Спорт", "спорт")
f_leagues = multi("Лига", "лига")
f_types = multi("Тип ставки", "тип_ставки")
f_results = st.sidebar.multiselect(
    "Результат",
    options=["WIN", "LOSE", "RETURN", "PENDING"],
    default=[],
)

df = df_all.copy()
if isinstance(date_range, tuple) and len(date_range) == 2:
    start, end = date_range
    df = df[(df["дата"].dt.date >= start) & (df["дата"].dt.date <= end)]
for col, selected in [
    ("канал", f_channels), ("спорт", f_sports), ("лига", f_leagues),
    ("тип_ставки", f_types), ("результат", f_results),
]:
    if selected:
        df = df[df[col].isin(selected)]

if len(df) == 0:
    st.warning("По текущим фильтрам ставок нет. Сбрось часть фильтров в сайдбаре.")
    st.stop()

# ============================================================
# ВЕРХНИЕ МЕТРИКИ
# ============================================================
total_stake = df["ставка"].sum()
total_profit = df["профит"].sum()
roi = calc_roi(df)
winrate = calc_winrate(df)
wins = (df["результат"] == "WIN").sum()
loses = (df["результат"] == "LOSE").sum()
returns = (df["результат"] == "RETURN").sum()
pending = (df["результат"] == "PENDING").sum()

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Ставок", len(df))
c2.metric("Оборот", f"{total_stake:,.0f} ₽".replace(",", " "))
c3.metric(
    "Прибыль",
    f"{total_profit:+,.0f} ₽".replace(",", " "),
    delta_color="normal" if total_profit >= 0 else "inverse",
)
c4.metric("ROI", f"{roi:+.2f}%")
c5.metric("Винрейт", f"{winrate:.1f}%",
          help="Доля WIN среди WIN+LOSE (RETURN и PENDING не учитываются)")

caption_bits = [f"🟢 WIN: {wins}", f"🔴 LOSE: {loses}", f"⚪ RETURN: {returns}"]
if pending:
    caption_bits.append(f"⏳ PENDING: {pending}")
st.caption("  •  ".join(caption_bits))

if len(df) < MIN_N_FOR_RELIABLE:
    st.warning(
        f"⚠️ Маленькая выборка ({len(df)} ставок). "
        f"Для надёжных выводов нужно хотя бы {MIN_N_FOR_RELIABLE}+. "
        "Сейчас многое может быть статистическим шумом, а не реальным сигналом."
    )

st.divider()

# ============================================================
# ТАБЫ
# ============================================================
tab_channels, tab_types, tab_heat, tab_sport, tab_combo, tab_table = st.tabs([
    "📺 Каналы",
    "🎲 Типы ставок",
    "🔥 Канал × тип (тепло)",
    "🏀 Спорт и лиги",
    "🏆 Топы комбинаций",
    "📋 Все ставки",
])


def render_bar(df_in: pd.DataFrame, by: str, *, title: str = "", caption: str = ""):
    if title:
        st.subheader(title)
    if caption:
        st.caption(caption)

    agg = aggregate(df_in, by)
    if len(agg) == 0:
        st.info("Нет данных для отображения.")
        return

    fig = px.bar(
        agg, x="ROI%", y=by, orientation="h",
        color="ROI%",
        color_continuous_scale=["#d62728", "#ffffff", "#2ca02c"],
        color_continuous_midpoint=0,
        text="ROI%",
        hover_data={"ставок": True, "оборот": ":,.0f", "профит": ":+,.0f"},
    )
    fig.update_traces(texttemplate="%{text}%", textposition="outside")
    fig.update_layout(
        height=max(280, 50 + 35 * len(agg)),
        yaxis={"categoryorder": "total ascending"},
        showlegend=False,
        margin=dict(l=0, r=0, t=10, b=0),
        xaxis_title="ROI, %",
    )
    st.plotly_chart(fig, use_container_width=True)

    display = agg.copy()
    display["надёжность"] = display["ставок"].apply(reliability_marker)
    display["оборот"] = display["оборот"].map(lambda x: f"{x:,.0f}".replace(",", " "))
    display["профит"] = display["профит"].map(lambda x: f"{x:+,.0f}".replace(",", " "))
    display["ROI%"] = display["ROI%"].map(lambda x: f"{x:+.1f}%")
    st.dataframe(display, use_container_width=True, hide_index=True)


with tab_channels:
    render_bar(df, "канал",
               title="ROI по каналу прогноза",
               caption="Какие источники прогнозов реально зарабатывают, а какие тянут вниз")

with tab_types:
    render_bar(df, "тип_ставки",
               title="ROI по типу ставки",
               caption="Какие категории ставок работают, а какие сливают")
    st.divider()
    render_bar(df, "направление",
               title="ROI по направлению",
               caption="Больше/меньше, плюс/минус, П1/П2 и т.д.")

with tab_heat:
    st.subheader("Тепловая карта: канал × тип ставки")
    st.caption(
        "**Главный экран.** Зелёные клетки = сочетания, которые зарабатывают. "
        "Красные = сливают. Серые = такой комбинации нет в данных. "
        "Внутри клетки — ROI %."
    )
    heat = df.groupby(["канал", "тип_ставки"]).agg(
        ставок=("профит", "count"),
        оборот=("ставка", "sum"),
        профит=("профит", "sum"),
    ).reset_index()
    heat["ROI%"] = (heat["профит"] / heat["оборот"] * 100).round(1)
    pivot_roi = heat.pivot(index="канал", columns="тип_ставки", values="ROI%")
    pivot_n = heat.pivot(index="канал", columns="тип_ставки", values="ставок")
    fig = px.imshow(
        pivot_roi,
        text_auto=".0f",
        aspect="auto",
        color_continuous_scale=["#d62728", "#ffffff", "#2ca02c"],
        color_continuous_midpoint=0,
        labels=dict(color="ROI %"),
    )
    fig.update_traces(
        customdata=pivot_n.values,
        hovertemplate=("Канал: %{y}<br>Тип: %{x}<br>"
                       "ROI: %{z}%<br>Ставок: %{customdata}<extra></extra>"),
    )
    fig.update_layout(
        height=max(400, 80 + 40 * pivot_roi.shape[0]),
        margin=dict(l=0, r=0, t=20, b=0),
        xaxis={"side": "bottom", "tickangle": -30},
    )
    st.plotly_chart(fig, use_container_width=True)
    st.caption(
        f"💡 ROI считается даже на 1 ставке. Если в клетке мало ставок — "
        f"цифре пока доверять рано. Для надёжности нужно ~{MIN_N_FOR_RELIABLE}+."
    )

with tab_sport:
    render_bar(df, "спорт", title="ROI по спорту")
    st.divider()
    render_bar(df, "лига", title="ROI по лиге")

with tab_combo:
    st.subheader("Топ комбинаций: канал + тип ставки")
    st.caption("Самые прибыльные и убыточные конкретные сочетания.")
    combo = df.groupby(["канал", "тип_ставки"]).agg(
        ставок=("профит", "count"),
        оборот=("ставка", "sum"),
        профит=("профит", "sum"),
    ).reset_index()
    combo["ROI%"] = (combo["профит"] / combo["оборот"] * 100).round(1)

    max_n = int(combo["ставок"].max()) if len(combo) else 1
    min_n = st.slider(
        "Минимум ставок в комбинации",
        min_value=1, max_value=max(max_n, 1), value=min(2, max_n),
        help="Группы с меньшим числом ставок не показываем — это шум",
    )
    combo_f = combo[combo["ставок"] >= min_n].copy()

    if len(combo_f) == 0:
        st.info("Нет комбинаций с таким числом ставок. Сдвинь слайдер влево.")
    else:
        ca, cb = st.columns(2)
        with ca:
            st.markdown("### 🟢 Топ прибыльных")
            top = combo_f.sort_values("ROI%", ascending=False).head(10).copy()
            top["оборот"] = top["оборот"].map(lambda x: f"{x:,.0f}".replace(",", " "))
            top["профит"] = top["профит"].map(lambda x: f"{x:+,.0f}".replace(",", " "))
            top["ROI%"] = top["ROI%"].map(lambda x: f"{x:+.1f}%")
            st.dataframe(top, use_container_width=True, hide_index=True)
        with cb:
            st.markdown("### 🔴 Топ убыточных")
            bot = combo_f.sort_values("ROI%", ascending=True).head(10).copy()
            bot["оборот"] = bot["оборот"].map(lambda x: f"{x:,.0f}".replace(",", " "))
            bot["профит"] = bot["профит"].map(lambda x: f"{x:+,.0f}".replace(",", " "))
            bot["ROI%"] = bot["ROI%"].map(lambda x: f"{x:+.1f}%")
            st.dataframe(bot, use_container_width=True, hide_index=True)

with tab_table:
    st.subheader("Все ставки (с учётом фильтров)")
    st.caption("Кликни на колонку, чтобы отсортировать.")
    show_cols = [c for c in (REQUIRED_COLUMNS + OPTIONAL_COLUMNS) if c in df.columns]
    table = df[show_cols].sort_values("дата", ascending=False).copy()
    table["дата"] = table["дата"].dt.strftime("%Y-%m-%d")
    st.dataframe(table, use_container_width=True, hide_index=True)

    csv = table.to_csv(index=False).encode("utf-8-sig")
    st.download_button("💾 Скачать в CSV", csv,
                       file_name="stavki_filtered.csv", mime="text/csv")
