# -*- coding: utf-8 -*-
"""
EL MOLTQA — Operations & Sales Dashboard
Reads the existing Google Sheet without changing the source.
The workbook contains monthly tabs with three separate tables:
الحجوزات / التعاقدات / الالغاءات.

This version is intentionally built around the ACTUAL worksheet structure:
- section/banner row
- one main header row beginning with "م"
- optional second header row for unit details
- records until the next section
"""

import re
from io import BytesIO
from urllib.parse import quote
import urllib.request

import numpy as np
import pandas as pd
import streamlit as st

# =========================================================
# CONFIG
# =========================================================
st.set_page_config(
    page_title="EL MOLTQA | Operations Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

SHEET_ID = "1lHloVtHag6yZs02q8XLcTu-XOF5Pn87rAQNOStkPmc4"
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit"

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
MONTHS = list(MONTH_TAB_MAP)
CATEGORIES = ["الحجوزات", "التعاقدات", "الالغاءات"]
ICON = {"الحجوزات": "📥", "التعاقدات": "📝", "الالغاءات": "↩️"}

# =========================================================
# STYLE
# =========================================================
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800&display=swap');
    html, body, [class*="css"] { font-family: "Cairo", sans-serif; }
    .stApp { background: #f6f7fb; }
    [data-testid="stSidebar"] { background: #111827; }
    [data-testid="stSidebar"] * { color: #f9fafb !important; }
    .brand { font-size: 2rem; font-weight: 800; color:#111827; margin-bottom:0; }
    .muted { color:#6b7280; }
    .hero {
      background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
      border:1px solid #e5e7eb; border-radius:20px; padding:24px 26px;
      box-shadow:0 8px 24px rgba(15,23,42,.05); margin-bottom:18px;
    }
    .hero-title { font-size:1.8rem; font-weight:800; color:#111827; }
    .hero-sub { margin-top:4px; color:#6b7280; font-size:.92rem; }
    .kpi {
      background:#fff; border:1px solid #e5e7eb; border-radius:16px;
      padding:18px 20px; min-height:118px; box-shadow:0 4px 14px rgba(15,23,42,.045);
    }
    .kpi-label { font-size:.84rem; color:#6b7280; font-weight:700; }
    .kpi-value { font-size:1.65rem; color:#111827; font-weight:800; margin-top:5px; }
    .kpi-note { font-size:.72rem; color:#9ca3af; margin-top:3px; }
    .card {
      background:#fff; border:1px solid #e5e7eb; border-radius:18px;
      padding:18px; box-shadow:0 4px 14px rgba(15,23,42,.04);
    }
    .small-tag { display:inline-block; padding:5px 10px; border-radius:999px; background:#eef2ff; color:#4338ca; font-size:.75rem; font-weight:700; }
    div[data-testid="stMetric"] { background:#fff; border:1px solid #e5e7eb; border-radius:14px; }
    .block-container { padding-top:1.5rem; padding-bottom:3rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# TEXT / STRUCTURE HELPERS
# =========================================================
def clean(v) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    return str(v).strip()


def norm(v) -> str:
    s = clean(v)
    s = re.sub(r"[\u064B-\u065F\u0670\u0640\u200f\u200e\s]+", "", s)
    s = s.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    s = s.replace("ى", "ي").lower()
    return s


def section_from_row(row) -> str | None:
    """A section banner is short and mostly empty; data notes may contain the same words."""
    vals = [clean(v) for v in row.tolist() if clean(v)]
    if not vals or len(vals) > 4:
        return None
    text = norm(" ".join(vals))
    if "حجوز" in text or text in {"حجز", "الحجز"}:
        return "الحجوزات"
    if "تعاقد" in text or text in {"عقد", "عقود", "العقد"}:
        return "التعاقدات"
    if "الغاء" in text or "الغاءات" in text:
        return "الالغاءات"
    return None


def find_section_rows(raw: pd.DataFrame):
    found = []
    for i in range(len(raw)):
        kind = section_from_row(raw.iloc[i])
        if kind:
            found.append((i, kind))
    # keep the real section sequence; ignore duplicate accidental matches close together
    result = []
    last_row = -100
    for row, kind in sorted(found, key=lambda x: x[0]):
        if row - last_row >= 3:
            result.append((row, kind))
            last_row = row
    return result


def looks_like_header(row) -> bool:
    vals = [norm(v) for v in row.tolist()]
    return (
        len(vals) > 2
        and vals[0] == "م"
        and any("اسمالعميل" in v for v in vals)
        and any("كودالعميل" in v for v in vals)
    )


def find_header_row(raw, section_row, end_row):
    for r in range(section_row + 1, min(end_row, section_row + 10)):
        if looks_like_header(raw.iloc[r]):
            return r
    return None


def is_subheader(row) -> bool:
    vals = [norm(v) for v in row.tolist() if norm(v)]
    markers = ("رقمالعمارة", "رقمالوحده", "المساحه", "الحديقه", "اسمالمرحله", "المشروع")
    return sum(any(m in v for m in markers) for v in vals) >= 2


def build_columns(raw, header_row, subheader_row):
    width = raw.shape[1]
    top = [clean(x) for x in raw.iloc[header_row, :width]]
    sub = [clean(x) for x in raw.iloc[subheader_row, :width]] if subheader_row is not None else [""] * width
    names = []
    for i in range(width):
        # Subheader is the actual name for unit-detail columns; otherwise use the main header.
        name = sub[i] or top[i] or f"عمود {i+1}"
        names.append(name)
    seen = {}
    out = []
    for n in names:
        seen[n] = seen.get(n, 0) + 1
        out.append(n if seen[n] == 1 else f"{n} ({seen[n]})")
    return out


def extract_section(raw, section_row, next_section_row):
    end = next_section_row if next_section_row is not None else len(raw)
    h = find_header_row(raw, section_row, end)
    if h is None:
        return pd.DataFrame(), {"section_row": section_row + 1, "header_row": None, "subheader_row": None}

    sub = h + 1 if h + 1 < end and is_subheader(raw.iloc[h + 1]) else None
    cols = build_columns(raw, h, sub)
    start = sub + 1 if sub is not None else h + 1
    data = raw.iloc[start:end, :len(cols)].copy()
    data.columns = cols

    # Valid source records have a populated serial in column A. This removes blank/total rows
    # without requiring the serial to be numeric.
    first = data.iloc[:, 0].map(clean) if not data.empty else pd.Series(dtype=str)
    nfirst = first.map(norm)
    keep = first.ne("") & ~nfirst.isin({"م", "total", "totals"})
    keep &= ~nfirst.str.contains("اجمال|مجموع", regex=True, na=False)
    data = data.loc[keep].copy()

    for c in data.columns:
        data[c] = data[c].map(clean)

    return data.reset_index(drop=True), {
        "section_row": section_row + 1,
        "header_row": h + 1,
        "subheader_row": sub + 1 if sub is not None else None,
    }


def parse_month(raw):
    result = {c: pd.DataFrame() for c in CATEGORIES}
    diagnostics = {c: {} for c in CATEGORIES}
    sections = find_section_rows(raw)
    for i, (section_row, category) in enumerate(sections):
        # only first occurrence of each real category is used
        if not result[category].empty or diagnostics[category].get("section_row"):
            continue
        nxt = sections[i + 1][0] if i + 1 < len(sections) else None
        result[category], diagnostics[category] = extract_section(raw, section_row, nxt)
    return result, diagnostics, sections

# =========================================================
# GOOGLE SHEETS LOADING
# =========================================================

def xlsx_export_url(sheet_id):
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"


def csv_export_url(sheet_id, tab):
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={quote(tab)}"


@st.cache_data(ttl=300, show_spinner=False)
def load_tab(tab_name):
    """Load the exact worksheet as raw cells. XLSX first; CSV fallback."""
    errors = []
    try:
        req = urllib.request.Request(
            xlsx_export_url(SHEET_ID),
            headers={"User-Agent": "Mozilla/5.0"},
        )
        with urllib.request.urlopen(req, timeout=30) as r:
            content = r.read()
        book = pd.ExcelFile(BytesIO(content), engine="openpyxl")
        actual = next((s for s in book.sheet_names if s == tab_name), None)
        if actual is None:
            actual = next((s for s in book.sheet_names if s.strip().lower() == tab_name.strip().lower()), None)
        if actual is None:
            raise ValueError(f"Tab not found: {tab_name}")
        df = pd.read_excel(book, sheet_name=actual, header=None, dtype=str, keep_default_na=False)
        return df.fillna(""), "xlsx", None
    except Exception as e:
        errors.append(str(e))

    try:
        req = urllib.request.Request(
            csv_export_url(SHEET_ID, tab_name),
            headers={"User-Agent": "Mozilla/5.0"},
        )
        with urllib.request.urlopen(req, timeout=30) as r:
            df = pd.read_csv(r, header=None, dtype=str, keep_default_na=False)
        return df.fillna(""), "csv", None
    except Exception as e:
        errors.append(str(e))

    return None, None, " | ".join(errors[-2:])


@st.cache_data(ttl=300, show_spinner=False)
def load_all_months():
    all_rows = []
    diagnostics = {}
    raw_sizes = {}
    failures = []
    for month, tab in MONTH_TAB_MAP.items():
        raw, source, err = load_tab(tab)
        if raw is None:
            failures.append(f"{month} ({tab})")
            continue
        raw_sizes[month] = (len(raw), raw.shape[1], source)
        parsed, diag, sections = parse_month(raw)
        diagnostics[month] = {"tab": tab, "diagnostics": diag, "sections": [(r + 1, k) for r, k in sections]}
        for category, frame in parsed.items():
            if frame.empty:
                continue
            x = frame.copy()
            x.insert(0, "الشهر", month)
            x.insert(1, "النوع", category)
            all_rows.append(x)
    if all_rows:
        union = pd.concat(all_rows, ignore_index=True, sort=False).fillna("")
    else:
        union = pd.DataFrame()
    return union, diagnostics, raw_sizes, failures

# =========================================================
# DATA DISPLAY HELPERS
# =========================================================
def numeric_value(series):
    s = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("،", "", regex=False)
        .str.replace("٬", "", regex=False)
        .str.replace(" ", "", regex=False)
    )
    return pd.to_numeric(s, errors="coerce")


def find_col(df, names):
    for c in df.columns:
        nc = norm(c)
        for n in names:
            if norm(n) in nc:
                return c
    return None


def money_col(df):
    return find_col(df, ["اجمالي الوحده", "اجمالي الوحدة", "مبلغ المقدم", "مبلع الحجز", "اجمالي"])


def area_col(df):
    return find_col(df, ["المساحه", "المساحة"])


def prepare_view(df):
    """Readable copy for UI only; source dataframe is not changed."""
    out = df.copy()
    date_words = ["تاريخ"]
    for c in out.columns:
        if any(w in clean(c) for w in date_words):
            # Handle Excel serial dates while preserving ordinary text dates.
            nums = pd.to_numeric(out[c], errors="coerce")
            converted = pd.Series(out[c], index=out.index, dtype="object")
            mask = nums.between(30000, 60000)
            if mask.any():
                dt = pd.to_datetime(nums, unit="D", origin="1899-12-30", errors="coerce")
                converted.loc[mask] = dt.loc[mask].dt.strftime("%Y-%m-%d")
            out[c] = converted
    return out


def reorder_columns(df):
    priority = [
        "الشهر", "النوع", "م", "كود العميل", "كود العميل  / رقم التليفون",
        "اسم العميل", "تيم ليدر", "رقم العمارة", "رقم الوحده", "المساحه",
        "الحديقه", "اسم المرحلة", "المشروع", "مصدر العميل", "اسم شركه البروكر",
        "بروكر", "اجمالي الوحده", "تاريخ الحجز", "تاريخ التعاقد", "تاريخ الالغاء",
        "تاريخ نزول العميل", "المقدم", "مبلع الحجز", " مبلغ المقدم ", "طريقة الدفع",
        "طريقه الدفع", "سعر المتر", "نسبه المقدم", "اسم البائع", "ملحوظات", "سبب الاسترداد",
    ]
    front = []
    for p in priority:
        for c in df.columns:
            if c == p and c not in front:
                front.append(c)
    rest = [c for c in df.columns if c not in front]
    return df[front + rest]

# =========================================================
# MAIN
# =========================================================
def main():
    with st.sidebar:
        st.markdown("<div style='font-size:25px;font-weight:800;'>EL MOLTQA</div>", unsafe_allow_html=True)
        st.markdown("<div style='color:#9ca3af;font-size:12px;margin-bottom:18px;'>Operations & Sales</div>", unsafe_allow_html=True)

        if st.button("🔄 تحديث البيانات", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

        st.divider()
        st.markdown("### العرض")
        view_mode = st.radio(
            "",
            ["نظرة عامة", "الحجوزات", "التعاقدات", "الالغاءات", "كل البيانات"],
            label_visibility="collapsed",
        )

        selected_month = st.selectbox("الشهر", ["كل الشهور"] + MONTHS)
        search = st.text_input("بحث", placeholder="العميل، الوحدة، العمارة، Team Leader ...")

        st.divider()
        st.caption("المصدر: Google Sheets")
        st.caption("البيانات تُحدّث عند الطلب أو تلقائياً كل 5 دقائق.")

    with st.spinner("جاري قراءة جميع الشهور والجداول..."):
        all_data, diagnostics, raw_sizes, failures = load_all_months()

    if all_data.empty:
        st.error("لم يتم العثور على سجلات. راجعي صلاحية Google Sheet: Anyone with the link → Viewer.")
        if failures:
            st.caption("تعذر تحميل: " + ", ".join(failures))
        return

    # Base view
    if view_mode == "نظرة عامة" or view_mode == "كل البيانات":
        current = all_data.copy()
    else:
        current = all_data[all_data["النوع"] == view_mode].copy()

    if selected_month != "كل الشهور":
        current = current[current["الشهر"] == selected_month].copy()

    if search.strip():
        q = search.strip().lower()
        mask = pd.Series(False, index=current.index)
        for c in current.columns:
            mask |= current[c].astype(str).str.lower().str.contains(q, regex=False, na=False)
        current = current.loc[mask].copy()

    # Header
    st.markdown(
        f"<div class='hero'><div class='hero-title'>EL MOLTQA — متابعة المبيعات والعمليات</div>"
        f"<div class='hero-sub'>بيانات فعلية من الشيت • {('كل الشهور' if selected_month == 'كل الشهور' else selected_month)}"
        f" <span class='small-tag'>{len(current):,} سجل ظاهر</span></div></div>",
        unsafe_allow_html=True,
    )

    # =====================================================
    # OVERVIEW
    # =====================================================
    if view_mode == "نظرة عامة":
        counts = all_data["النوع"].value_counts()
        total_records = len(all_data)
        value_column = money_col(all_data)
        total_value = numeric_value(all_data[value_column]).sum() if value_column else np.nan
        leaders_col = find_col(all_data, ["تيم ليدر"])
        leader_count = all_data[leaders_col].replace("", np.nan).nunique() if leaders_col else 0

        k1,k2,k3,k4 = st.columns(4)
        cards=[
            (k1,"إجمالي السجلات",f"{total_records:,}","من كل الشهور والأقسام"),
            (k2,"الحجوزات",f"{int(counts.get('الحجوزات',0)):,}","كل السجلات المحجوزة"),
            (k3,"التعاقدات",f"{int(counts.get('التعاقدات',0)):,}","كل السجلات المتعاقد عليها"),
            (k4,"الإلغاءات",f"{int(counts.get('الالغاءات',0)):,}","كل السجلات الملغاة"),
        ]
        for col,label,val,note in cards:
            with col:
                st.markdown(f"<div class='kpi'><div class='kpi-label'>{label}</div><div class='kpi-value'>{val}</div><div class='kpi-note'>{note}</div></div>",unsafe_allow_html=True)

        st.markdown("### ملخص شهري")
        if selected_month == "كل الشهور":
            summary = (all_data.groupby(["الشهر","النوع"]).size().unstack(fill_value=0).reindex(MONTHS))
            for c in CATEGORIES:
                if c not in summary.columns: summary[c]=0
            summary = summary[CATEGORIES]
            st.bar_chart(summary)

        c1,c2 = st.columns(2)
        with c1:
            st.markdown("#### توزيع الأقسام")
            st.bar_chart(counts.reindex(CATEGORIES).fillna(0))
        with c2:
            st.markdown("#### أكثر Team Leaders ظهوراً")
            if leaders_col:
                leaders = all_data[leaders_col].replace("", np.nan).dropna().value_counts().head(10)
                st.bar_chart(leaders)
            else:
                st.info("لا يوجد عمود Team Leader في البيانات الحالية.")

        st.markdown("### أهم الأرقام المتاحة")
        c1,c2,c3 = st.columns(3)
        with c1:
            st.metric("إجمالي القيمة", f"{total_value:,.0f}" if pd.notna(total_value) else "—")
        with c2:
            ac = area_col(all_data)
            area_total = numeric_value(all_data[ac]).sum() if ac else np.nan
            st.metric("إجمالي المساحة", f"{area_total:,.0f}" if pd.notna(area_total) else "—")
        with c3:
            st.metric("عدد Team Leaders", f"{leader_count:,}")

    else:
        # =================================================
        # DETAIL VIEW
        # =================================================
        if current.empty:
            st.warning("لا توجد نتائج مطابقة للاختيار الحالي.")
            return

        # Smart filters based on real columns.
        st.markdown("### تصفية البيانات")
        filter_names = ["تيم ليدر", "المشروع", "مصدر العميل", "اسم المرحلة", "اسم البائع"]
        filters = {}
        available = [c for c in filter_names if c in current.columns]
        cols = st.columns(min(5, max(1, len(available)))) if available else []
        for i,c in enumerate(available[:5]):
            vals = sorted([x for x in current[c].dropna().astype(str).unique().tolist() if x.strip()])
            with cols[i]:
                filters[c] = st.multiselect(c, vals, key=f"f_{view_mode}_{selected_month}_{c}")
        for c, vals in filters.items():
            if vals: current = current[current[c].isin(vals)]

        c1,c2,c3,c4 = st.columns(4)
        c1.metric("السجلات", f"{len(current):,}")
        vc = money_col(current)
        v = numeric_value(current[vc]).sum() if vc else np.nan
        c2.metric("إجمالي القيمة", f"{v:,.0f}" if pd.notna(v) else "—")
        ac = area_col(current)
        a = numeric_value(current[ac]).sum() if ac else np.nan
        c3.metric("إجمالي المساحة", f"{a:,.0f}" if pd.notna(a) else "—")
        lc = find_col(current,["تيم ليدر"])
        c4.metric("Team Leaders", f"{current[lc].replace('',np.nan).nunique():,}" if lc else "—")

        # Charts based on actual content.
        ch1,ch2 = st.columns(2)
        with ch1:
            by_month = current.groupby("الشهر").size().reindex(MONTHS).fillna(0)
            st.markdown("#### السجلات حسب الشهر")
            st.bar_chart(by_month)
        with ch2:
            if lc:
                leaders = current[lc].replace("",np.nan).dropna().value_counts().head(10)
                st.markdown("#### توزيع Team Leaders")
                st.bar_chart(leaders)
            else:
                st.info("لا يوجد Team Leader في هذا القسم.")

        st.markdown("### البيانات")
        display = prepare_view(reorder_columns(current))
        st.dataframe(display, use_container_width=True, hide_index=True, height=620)

        csv_bytes = display.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "⬇️ تنزيل البيانات الحالية CSV",
            csv_bytes,
            file_name=f"EL_MOLTQA_{view_mode}_{selected_month}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    # =====================================================
    # FULL DATA
    # =====================================================
    if view_mode == "كل البيانات":
        st.markdown("### البيانات الكاملة")
        display = prepare_view(reorder_columns(current))
        st.dataframe(display, use_container_width=True, hide_index=True, height=680)
        st.download_button(
            "⬇️ تنزيل كل البيانات CSV",
            display.to_csv(index=False).encode("utf-8-sig"),
            file_name="EL_MOLTQA_All_Data.csv",
            mime="text/csv",
            use_container_width=True,
        )

    # =====================================================
    # DIAGNOSTICS — useful, but not in the way
    # =====================================================
    with st.expander("فحص مصدر البيانات"):
        for month in MONTHS:
            info = diagnostics.get(month, {})
            sizes = raw_sizes.get(month, (0,0,"-"))
            st.write(
                f"**{month}** — Tab `{info.get('tab', '-')}` — "
                f"{sizes[0]:,} صف × {sizes[1]:,} عمود — مصدر القراءة: `{sizes[2]}`"
            )
            d = info.get("diagnostics", {})
            for cat in CATEGORIES:
                x = d.get(cat, {})
                if x.get("header_row"):
                    st.caption(
                        f"{ICON[cat]} {cat}: section {x.get('section_row')}, "
                        f"header {x.get('header_row')}, "
                        f"subheader {x.get('subheader_row') or 'بدون'}"
                    )
        if failures:
            st.warning("تعذر تحميل: " + ", ".join(failures))


if __name__ == "__main__":
    main()
