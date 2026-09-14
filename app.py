# -*- coding: utf-8 -*-
"""
EL MOLTQA - Google Sheets Dashboard
------------------------------------
Reads the monthly company data directly from Google Sheets.

Expected structure:
- One Google Spreadsheet
- Monthly tabs: يناير ... سبتمبر
- Each tab contains sections:
    الحجوزات
    التعاقدات
    الالغاءات
- Each section has one/two header rows followed by data.

Google Sheet:
https://docs.google.com/spreadsheets/d/1lHloVtHag6yZs02q8XLcTu-XOF5Pn87rAQNOStkPmc4/edit?gid=928721772#gid=928721772

The sheet must be shared as:
"Anyone with the link -> Viewer"
"""

import re
from io import BytesIO
from urllib.parse import quote

import numpy as np
import pandas as pd
import streamlit as st


# =========================================================
# PAGE
# =========================================================
st.set_page_config(
    page_title="EL MOLTQA | Sales Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

SHEET_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1lHloVtHag6yZs02q8XLcTu-XOF5Pn87rAQNOStkPmc4"
    "/edit?gid=928721772#gid=928721772"
)

MONTH_TAB_MAP = {
    "يناير": "JAN",
    "فبراير": "FEB",
    "مارس": "MAR",
    "أبريل": "April",
    "مايو": "May",
    "يونيو": "june",
    "يوليو": "July",
    "أغسطس": "Aug",
    "سبتمبر": "SEP",
}
MONTHS_AR = list(MONTH_TAB_MAP.keys())

CATEGORIES = {
    "الحجوزات": "📥",
    "التعاقدات": "📝",
    "الالغاءات": "❌",
}

HEADER_HINTS = {
    "م", "كود العميل", "اسم العميل", "كود", "اسم",
    "اسم العميل ", "تيم ليدر", "رقم العميل",
}


# =========================================================
# PROFESSIONAL UI
# =========================================================
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: "Cairo", sans-serif;
    }

    .stApp {
        background: #f5f7fb;
    }

    [data-testid="stSidebar"] {
        background: #111827;
        border-right: 1px solid #1f2937;
    }

    [data-testid="stSidebar"] * {
        color: #f9fafb !important;
    }

    .main-title {
        font-size: 2.1rem;
        font-weight: 800;
        color: #111827;
        margin-bottom: 0.1rem;
    }

    .subtitle {
        color: #6b7280;
        font-size: 0.95rem;
        margin-bottom: 1.2rem;
    }

    .section-title {
        font-size: 1.25rem;
        font-weight: 800;
        color: #111827;
        margin: 1rem 0 0.6rem 0;
    }

    .kpi {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 16px;
        padding: 18px 20px;
        min-height: 125px;
        box-shadow: 0 3px 12px rgba(15, 23, 42, 0.05);
    }

    .kpi-label {
        color: #6b7280;
        font-size: 0.85rem;
        font-weight: 600;
    }

    .kpi-value {
        color: #111827;
        font-size: 1.75rem;
        font-weight: 800;
        margin-top: 6px;
    }

    .kpi-note {
        color: #9ca3af;
        font-size: 0.75rem;
        margin-top: 2px;
    }

    .status {
        display: inline-block;
        padding: 5px 11px;
        border-radius: 999px;
        background: #ecfdf5;
        color: #047857;
        font-size: 0.78rem;
        font-weight: 700;
    }

    div[data-testid="stMetric"] {
        background: white;
        border: 1px solid #e5e7eb;
        padding: 14px;
        border-radius: 14px;
        box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04);
    }

    .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# GOOGLE SHEETS HELPERS
# =========================================================
def extract_sheet_id(url_or_id: str) -> str:
    value = str(url_or_id).strip()
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", value)
    return match.group(1) if match else value


def build_csv_url(sheet_id: str, tab_name: str) -> str:
    return (
        f"https://docs.google.com/spreadsheets/d/{sheet_id}"
        f"/gviz/tq?tqx=out:csv&sheet={quote(tab_name)}"
    )


@st.cache_data(ttl=300, show_spinner=False)
def fetch_tab_raw(sheet_id: str, tab_name: str) -> pd.DataFrame | None:
    """Read one Google Sheet tab as raw text."""
    try:
        url = build_csv_url(sheet_id, tab_name)
        return pd.read_csv(
            url,
            header=None,
            dtype=str,
            keep_default_na=False,
        )
    except Exception:
        return None


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


# =========================================================
# TABLE PARSING
# =========================================================
def normalize_arabic_text(value) -> str:
    """Normalize Arabic text so section names are detected despite
    extra spaces, hamza/diacritic differences, or singular/plural wording."""
    text = clean_text(value)
    if not text:
        return ""

    # Remove Arabic diacritics and Tatweel (ـ). The source uses long
    # decorative section titles such as "التعاقدــــــــــــات".
    text = re.sub(r"[\u064B-\u065F\u0670\u0640]", "", text)
    text = text.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    text = text.replace("ى", "ي")
    text = re.sub(r"[\s\u200f\u200e]+", "", text)
    text = re.sub(r"[|:_\-–—/\\]+", "", text)
    return text.strip().lower()


SECTION_ALIASES = {
    "الحجوزات": {
        normalize_arabic_text(x) for x in
        ["الحجوزات", "حجوزات", "الحجز", "حجز"]
    },
    "التعاقدات": {
        normalize_arabic_text(x) for x in
        ["التعاقدات", "تعاقدات", "التعاقد", "تعاقد", "العقود", "عقود"]
    },
    "الالغاءات": {
        normalize_arabic_text(x) for x in
        ["الالغاءات", "الإلغاءات", "الغاءات", "إلغاءات", "الالغاء", "الإلغاء", "الغاء", "إلغاء"]
    },
}


def find_section_rows(raw_df: pd.DataFrame) -> dict:
    """Find the three real section banners in each monthly tab.

    The source uses decorative Tatweel characters and inconsistent spellings
    such as "الالغاء" / "الالغاءات" and "التعاقدــــــــات". We detect the
    stable word stem instead of relying on an exact title.
    """
    positions = {}

    for idx, row in raw_df.iterrows():
        text = " ".join(
            normalize_arabic_text(value)
            for value in row
            if clean_text(value)
        )

        if not text:
            continue

        if "الحجز" in text or "الحجوز" in text:
            positions.setdefault("الحجوزات", idx)
        elif "تعاقد" in text or "عقد" in text:
            positions.setdefault("التعاقدات", idx)
        elif "الالغاء" in text or "الغاء" in text:
            positions.setdefault("الالغاءات", idx)

    return positions

def find_header_row(
    raw_df: pd.DataFrame,
    start_idx: int,
    search_window: int = 20,
    end_idx: int | None = None,
):
    """Find the real header row after a section banner.

    The source sheets are not perfectly consistent: some sections have one
    header row, some have a title/merged row followed by the header, and some
    use Arabic/English column names. We therefore score rows by header words
    instead of assuming a fixed number of rows.
    """
    header_terms = [
        "م", "اسم العميل", "name of client", "كود العميل", "رقم العميل",
        "تيم ليدر", "team leader", "phase", "المرحله", "المرحلة",
        "unit", "unit code", "unit type", "رقم الوحده", "رقم الوحدة",
        "العماره", "العمارة", "building", "floor", "status",
        "الحالة", "حالة", "قيمه الحجز", "قيمة الحجز", "تاريخ الحجز",
        "قيمه التعاقد", "قيمة التعاقد", "تاريخ التعاقد",
        "تاريخ التعاقد", "المساحة", "المساحه", "total area",
        "ملاحظات", "note", "notes", "source", "مصدر",
    ]

    natural_limit = start_idx + 1 + search_window
    if end_idx is not None:
        natural_limit = min(natural_limit, end_idx)
    limit = min(len(raw_df), natural_limit)
    best_idx, best_score = None, -1

    for idx in range(start_idx + 1, limit):
        vals = [clean_text(v) for v in raw_df.iloc[idx] if clean_text(v)]
        if len(vals) < 3:
            continue

        score = 0
        for value in vals:
            norm = normalize_arabic_text(value)
            low = value.lower()
            if any(term in norm or term.lower() in low for term in header_terms):
                score += 2

        # Headers normally contain many non-empty cells.
        score += min(len(vals), 15) * 0.15

        if score > best_score:
            best_idx, best_score = idx, score

    return best_idx if best_score >= 3 else None

def build_columns(
    raw_df: pd.DataFrame,
    header_idx: int,
    max_col: int,
) -> list:
    top_row = raw_df.iloc[header_idx, : max_col + 1].copy()
    top_row = top_row.replace("", np.nan).ffill()

    sub_row = None
    if header_idx + 1 < len(raw_df):
        sub_row = raw_df.iloc[
            header_idx + 1, : max_col + 1
        ]

    columns = []
    seen = {}

    for i in range(max_col + 1):
        sub_value = (
            clean_text(sub_row.iloc[i])
            if sub_row is not None
            else ""
        )

        if sub_value:
            name = sub_value
        else:
            top_value = clean_text(top_row.iloc[i])
            name = top_value or f"عمود_{i + 1}"

        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 1

        columns.append(name)

    return columns


def extract_section_table(
    raw_df: pd.DataFrame,
    header_idx: int,
    next_section_idx: int | None,
) -> pd.DataFrame:
    """Extract one monthly section exactly as it appears in the source.

    Every monthly tab uses:
        section title
        blank row
        main header row
        sub-header row
        data rows

    Some sections contain a totals row or blank rows at the bottom. The first
    column ("م") is the reliable record marker, so only rows with a numeric
    record number are returned. This prevents totals/header/blank rows from
    becoming fake records.
    """
    end_idx = next_section_idx if next_section_idx is not None else len(raw_df)

    # In the real workbook the two header rows are always immediately after
    # the section banner.
    header1_idx = header_idx + 2
    header2_idx = header_idx + 3
    data_start = header_idx + 4

    if header1_idx >= end_idx:
        return pd.DataFrame()

    width = raw_df.shape[1]
    columns = []

    for col_idx in range(width):
        top = clean_text(raw_df.iat[header1_idx, col_idx])
        sub = (
            clean_text(raw_df.iat[header2_idx, col_idx])
            if header2_idx < end_idx
            else ""
        )

        # The sub-header is more specific for grouped fields such as
        # "بيانات الوحده"; otherwise keep the main header.
        name = sub or top or f"عمود_{col_idx + 1}"
        columns.append(name)

    # Make duplicate headers unique without losing any source columns.
    seen = {}
    unique_columns = []
    for name in columns:
        seen[name] = seen.get(name, 0) + 1
        unique_columns.append(
            name if seen[name] == 1 else f"{name}_{seen[name]}"
        )

    data = raw_df.iloc[data_start:end_idx, :width].copy()
    data.columns = unique_columns

    if data.empty:
        return pd.DataFrame(columns=unique_columns)

    # Remove completely blank rows.
    data = data[
        data.apply(
            lambda row: any(clean_text(v) for v in row),
            axis=1,
        )
    ]

    # The actual records in this workbook have a numeric value in "م".
    # This also removes subtotal/total rows and decorative rows.
    serial = pd.to_numeric(data.iloc[:, 0], errors="coerce")
    data = data.loc[serial.notna()].copy()

    # Keep the original serial value, but normalize whitespace in text cells.
    for col in data.columns:
        data[col] = data[col].map(clean_text)

    return data.reset_index(drop=True)



def parse_month_tab(raw_df: pd.DataFrame) -> dict:
    """Parse a monthly tab from its three explicit sections.

    This is deliberately based on the actual monthly workbook layout rather
    than guessing categories from cell values. Each tab has separate sections
    for bookings, contracts and cancellations.
    """
    result = {
        "الحجوزات": pd.DataFrame(),
        "التعاقدات": pd.DataFrame(),
        "الالغاءات": pd.DataFrame(),
    }

    positions = find_section_rows(raw_df)
    if not positions:
        return result

    ordered = sorted(positions.items(), key=lambda item: item[1])

    for i, (category, section_row) in enumerate(ordered):
        next_row = ordered[i + 1][1] if i + 1 < len(ordered) else None

        if category not in result:
            continue

        result[category] = extract_section_table(
            raw_df,
            section_row,
            next_row,
        )

    return result


# =========================================================
# DATA NORMALIZATION
# =========================================================
def normalize_column_name(name: str) -> str:
    return re.sub(
        r"\s+",
        " ",
        clean_text(name),
    ).strip()


def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    df.columns = [
        normalize_column_name(c)
        for c in df.columns
    ]

    # Remove duplicate-looking columns safely.
    df = df.loc[:, ~df.columns.duplicated()]

    return df


def find_column(df: pd.DataFrame, keywords) -> str | None:
    for col in df.columns:
        text = clean_text(col).lower()

        for keyword in keywords:
            if keyword.lower() in text:
                return col

    return None


def numeric_series(df: pd.DataFrame, column: str | None):
    if not column or column not in df.columns:
        return pd.Series(
            np.nan,
            index=df.index,
            dtype=float,
        )

    values = (
        df[column]
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("٬", "", regex=False)
        .str.replace("،", "", regex=False)
        .str.replace(" ", "", regex=False)
    )

    return pd.to_numeric(
        values,
        errors="coerce",
    )


# =========================================================
# SEARCH / FILTERS
# =========================================================
def apply_filters(
    df: pd.DataFrame,
    search_text: str,
    selected_filters: dict,
) -> pd.DataFrame:

    result = df.copy()

    if search_text.strip():
        query = search_text.strip().lower()

        mask = pd.Series(
            False,
            index=result.index,
        )

        for col in result.columns:
            mask |= (
                result[col]
                .astype(str)
                .str.lower()
                .str.contains(
                    query,
                    regex=False,
                    na=False,
                )
            )

        result = result[mask]

    for col, selected_values in selected_filters.items():
        if (
            col in result.columns
            and selected_values
        ):
            result = result[
                result[col]
                .astype(str)
                .isin(selected_values)
            ]

    return result


# =========================================================
# MAIN
# =========================================================
def main():

    sheet_id = extract_sheet_id(SHEET_URL)

    # -----------------------------------------------------
    # SIDEBAR
    # -----------------------------------------------------
    with st.sidebar:

        st.markdown(
            """
            <div style="
                font-size:24px;
                font-weight:800;
                margin-bottom:2px;
            ">
                EL MOLTQA
            </div>
            <div style="
                color:#9ca3af;
                font-size:12px;
                margin-bottom:20px;
            ">
                Sales & Contracts Dashboard
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown("### البيانات")

        selected_month = st.selectbox(
            "الشهر",
            MONTHS_AR,
            index=0,
        )

        selected_category = st.selectbox(
            "نوع البيانات",
            list(CATEGORIES.keys()),
            format_func=lambda c: (
                f"{CATEGORIES[c]}  {c}"
            ),
        )

        st.divider()

        st.markdown("### التحكم")

        if st.button(
            "🔄 تحديث البيانات",
            use_container_width=True,
        ):
            st.cache_data.clear()
            st.rerun()

        st.caption(
            "يتم تحديث البيانات تلقائياً من Google Sheets "
            "كل 5 دقائق، ويمكنك التحديث يدوياً في أي وقت."
        )

        st.divider()

        st.markdown("### مصدر البيانات")

        st.success(
            "متصل بمصدر Google Sheets"
        )

        st.caption(
            "إذا لم تظهر البيانات، تأكد أن الشيت "
            "متاح للمشاهدة لأي شخص لديه الرابط."
        )

    # -----------------------------------------------------
    # LOAD
    # -----------------------------------------------------
    # The visible month names are Arabic, while the actual Google Sheet tabs
    # use the names JAN/FEB/MAR/April/May/june/July/Aug/SEP.
    tab_name = MONTH_TAB_MAP[selected_month]

    with st.spinner(
        f"جاري تحميل بيانات {tab_name}..."
    ):
        raw_df = fetch_tab_raw(
            sheet_id,
            tab_name,
        )

    if raw_df is None:
        st.error(
            f"تعذر تحميل ورقة **{tab_name}**."
        )

        st.info(
            "تأكد من أن Google Sheet مضبوط على "
            "**Anyone with the link → Viewer** "
            "وأن اسم الـ Tab مطابق لاسم الشهر."
        )

        with st.expander("رابط مصدر البيانات"):
            st.code(SHEET_URL)

        return

    sections = parse_month_tab(raw_df)

    df = sections.get(
        selected_category,
        pd.DataFrame(),
    )

    df = normalize_dataframe(df)

    # -----------------------------------------------------
    # HEADER
    # -----------------------------------------------------
    st.markdown(
        '<div class="main-title">EL MOLTQA</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        f"""
        <div class="subtitle">
            لوحة متابعة {selected_category}
            — {selected_month}
            <span class="status">Live Data</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if df.empty:
        st.warning(
            f"لا توجد بيانات ظاهرة في قسم {selected_category} "
            f"داخل ورقة {selected_month}."
        )
        return

    # -----------------------------------------------------
    # SEARCH
    # -----------------------------------------------------
    st.markdown(
        '<div class="section-title">البحث والتصفية</div>',
        unsafe_allow_html=True,
    )

    search_col, clear_col = st.columns(
        [5, 1],
        vertical_alignment="bottom",
    )

    with search_col:
        search_text = st.text_input(
            "بحث",
            placeholder=(
                "ابحث باسم العميل، كود العميل، "
                "الوحدة، العمارة أو أي قيمة..."
            ),
            label_visibility="collapsed",
        )

    with clear_col:
        clear_filters = st.button(
            "مسح الفلاتر",
            use_container_width=True,
        )

    # Detect useful categorical columns.
    categorical_candidates = []

    for col in df.columns:
        if col == "م":
            continue

        nunique = (
            df[col]
            .astype(str)
            .replace("", np.nan)
            .nunique(dropna=True)
        )

        if 1 < nunique <= 30:
            categorical_candidates.append(col)

    categorical_candidates = categorical_candidates[:5]

    selected_filters = {}

    if categorical_candidates:
        filter_cols = st.columns(
            min(len(categorical_candidates), 5)
        )

        for i, col in enumerate(
            categorical_candidates
        ):
            values = sorted(
                [
                    str(v)
                    for v in df[col].dropna().unique()
                    if str(v).strip()
                ]
            )

            with filter_cols[i]:
                selected = st.multiselect(
                    col,
                    values,
                    key=f"filter_{selected_month}_{selected_category}_{col}",
                )

                selected_filters[col] = (
                    [] if clear_filters else selected
                )

    filtered_df = apply_filters(
        df,
        search_text,
        selected_filters,
    )

    # -----------------------------------------------------
    # DYNAMIC KPIs
    # -----------------------------------------------------
    total_records = len(filtered_df)

    total_column = find_column(
        filtered_df,
        [
            "اجمالي",
            "إجمالي",
            "القيمة",
            "السعر",
            "مبلغ",
            "قيمة",
            "total",
            "price",
            "amount",
        ],
    )

    area_column = find_column(
        filtered_df,
        [
            "المساحة",
            "المساحه",
            "area",
        ],
    )

    team_column = find_column(
        filtered_df,
        [
            "تيم ليدر",
            "team leader",
            "team",
        ],
    )

    total_value = (
        numeric_series(
            filtered_df,
            total_column,
        ).sum()
        if total_column
        else None
    )

    total_area = (
        numeric_series(
            filtered_df,
            area_column,
        ).sum()
        if area_column
        else None
    )

    teams_count = (
        filtered_df[team_column]
        .astype(str)
        .replace("", np.nan)
        .nunique()
        if team_column
        else None
    )

    st.markdown(
        '<div class="section-title">ملخص البيانات</div>',
        unsafe_allow_html=True,
    )

    k1, k2, k3, k4 = st.columns(4)

    with k1:
        st.markdown(
            f"""
            <div class="kpi">
                <div class="kpi-label">
                    إجمالي السجلات
                </div>
                <div class="kpi-value">
                    {total_records:,}
                </div>
                <div class="kpi-note">
                    بعد تطبيق البحث والفلاتر
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with k2:
        value_text = (
            f"{total_value:,.0f}"
            if total_value is not None
            else "—"
        )

        st.markdown(
            f"""
            <div class="kpi">
                <div class="kpi-label">
                    إجمالي القيمة
                </div>
                <div class="kpi-value">
                    {value_text}
                </div>
                <div class="kpi-note">
                    {total_column or "لم يتم العثور على عمود قيمة"}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with k3:
        area_text = (
            f"{total_area:,.0f}"
            if total_area is not None
            else "—"
        )

        st.markdown(
            f"""
            <div class="kpi">
                <div class="kpi-label">
                    إجمالي المساحة
                </div>
                <div class="kpi-value">
                    {area_text}
                </div>
                <div class="kpi-note">
                    {area_column or "لم يتم العثور على عمود المساحة"}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with k4:
        team_text = (
            f"{teams_count:,}"
            if teams_count is not None
            else "—"
        )

        st.markdown(
            f"""
            <div class="kpi">
                <div class="kpi-label">
                    عدد الـ Team Leaders
                </div>
                <div class="kpi-value">
                    {team_text}
                </div>
                <div class="kpi-note">
                    القيم المختلفة في البيانات الحالية
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # -----------------------------------------------------
    # CHARTS
    # -----------------------------------------------------
    st.markdown(
        '<div class="section-title">Insights</div>',
        unsafe_allow_html=True,
    )

    chart1, chart2 = st.columns(2)

    with chart1:
        chart_col = next(
            (
                col
                for col in [
                    "تيم ليدر",
                    "مصدر العميل",
                    "الحالة",
                    "المرحلة",
                    "Phase",
                    "Status",
                ]
                if col in filtered_df.columns
            ),
            None,
        )

        if chart_col:
            counts = (
                filtered_df[chart_col]
                .astype(str)
                .replace("", np.nan)
                .dropna()
                .value_counts()
                .head(10)
            )

            if not counts.empty:
                st.markdown(
                    f"**التوزيع حسب {chart_col}**"
                )
                st.bar_chart(counts)

        else:
            st.info(
                "لا يوجد عمود تصنيف مناسب لعرض الرسم."
            )

    with chart2:
        # Find a numeric column automatically.
        numeric_candidates = []

        for col in filtered_df.columns:
            series = numeric_series(
                filtered_df,
                col,
            )

            if series.notna().sum() >= 3:
                numeric_candidates.append(col)

        if numeric_candidates:
            chart_numeric_col = st.selectbox(
                "اختر المؤشر للرسم",
                numeric_candidates,
                key=(
                    f"numeric_chart_"
                    f"{selected_month}_"
                    f"{selected_category}"
                ),
            )

            values = numeric_series(
                filtered_df,
                chart_numeric_col,
            ).dropna()

            if not values.empty:
                chart_data = pd.DataFrame(
                    {
                        "القيمة": values.values
                    }
                )

                st.markdown(
                    f"**توزيع {chart_numeric_col}**"
                )
                st.line_chart(
                    chart_data,
                    use_container_width=True,
                )
        else:
            st.info(
                "لا توجد بيانات رقمية كافية لعرض الرسم."
            )

    # -----------------------------------------------------
    # DATA TABLE
    # -----------------------------------------------------
    st.markdown(
        '<div class="section-title">تفاصيل البيانات</div>',
        unsafe_allow_html=True,
    )

    st.caption(
        f"عرض {len(filtered_df):,} سجل من أصل {len(df):,}"
    )

    if filtered_df.empty:
        st.warning(
            "لا توجد نتائج مطابقة للبحث أو الفلاتر."
        )
    else:
        st.dataframe(
            filtered_df,
            use_container_width=True,
            hide_index=True,
            height=520,
        )

    # -----------------------------------------------------
    # DOWNLOAD
    # -----------------------------------------------------
    st.markdown(
        '<div class="section-title">التصدير</div>',
        unsafe_allow_html=True,
    )

    csv_bytes = filtered_df.to_csv(
        index=False
    ).encode("utf-8-sig")

    excel_buffer = BytesIO()

    try:
        with pd.ExcelWriter(
            excel_buffer,
            engine="openpyxl",
        ) as writer:
            filtered_df.to_excel(
                writer,
                index=False,
                sheet_name="Data",
            )

        excel_bytes = excel_buffer.getvalue()
    except Exception:
        excel_bytes = None

    download_col1, download_col2 = st.columns(2)

    with download_col1:
        st.download_button(
            "⬇️ تنزيل CSV",
            data=csv_bytes,
            file_name=(
                f"EL_MOLTQA_"
                f"{selected_category}_"
                f"{selected_month}.csv"
            ),
            mime="text/csv",
            use_container_width=True,
        )

    with download_col2:
        if excel_bytes:
            st.download_button(
                "⬇️ تنزيل Excel",
                data=excel_bytes,
                file_name=(
                    f"EL_MOLTQA_"
                    f"{selected_category}_"
                    f"{selected_month}.xlsx"
                ),
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                ),
                use_container_width=True,
            )

    # -----------------------------------------------------
    # FOOTER / DEBUG
    # -----------------------------------------------------
    with st.expander("معلومات المصدر"):
        st.write(
            f"Google Sheet ID: `{sheet_id}`"
        )
        st.write(
            f"Tab: `{tab_name}`"
        )
        st.write(
            f"Section: `{selected_category}`"
        )
        st.write(
            f"Raw rows loaded: `{len(raw_df):,}`"
        )

        st.caption(
            "هذه المعلومات للمساعدة في فحص الاتصال فقط."
        )


if __name__ == "__main__":
    main()
