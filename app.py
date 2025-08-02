import streamlit as st
import pandas as pd
from io import BytesIO
from urllib.parse import quote as urlquote  # אופציונלי אם תרצה קישור וואטסאפ

# ---------- Helper: convert 0-based column index to Excel column name ----------
def xl_col_to_name_local(col_idx: int) -> str:
    col_idx += 1
    name = ""
    while col_idx > 0:
        col_idx, rem = divmod(col_idx - 1, 26)
        name = chr(65 + rem) + name
    return name

# ---------- Page & light styling ----------
st.set_page_config(page_title="FreshTrack Analytics", layout="centered")
st.markdown("""
    <style>
    body { background: linear-gradient(to bottom right, #fef3c7, #d1fae5, #dbeafe); }
    .title-text { font-size: 2.2rem; font-weight: 700;
        background: linear-gradient(to right, #16a34a, #2563eb, #7c3aed);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
    .upload-box { border: 4px dashed #6ee7b7; padding: 1rem; border-radius: 1rem; background-color: #ecfdf5; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 class='title-text' style='text-align:center'>🚀 FreshTrack Analytics</h1>", unsafe_allow_html=True)

# ---------- Uploads ----------
st.markdown("<div class='upload-box'>📥 העלה קובץ מחירון (Excel: מק\"ט, פריט, מחירון)</div>", unsafe_allow_html=True)
pricing_file = st.file_uploader("מחירון", type=["xls", "xlsx"], key="pricing")

st.markdown("<div class='upload-box'>📥 העלה קובץ הוצאות (Excel: תאריך, מק\"ט, פריט, כמות, מחיר לפני מע\"מ)</div>", unsafe_allow_html=True)
expense_file = st.file_uploader("הוצאות", type=["xls", "xlsx"], key="expense")

# ---------- Core compare (adds only one extra column: 'שינוי בשקלים') ----------
def compare_files(pricing_file, expense_file, month: int = 6):
    pricing_df = pd.read_excel(pricing_file)
    expenses_df = pd.read_excel(expense_file)

    # מחירון: בדיוק 3 עמודות בעברית
    pricing_df.columns = ['מקט', 'פריט', 'מחירון']

    # דרישות בסיס להוצאות
    if 'תאריך' not in expenses_df.columns:
        raise ValueError("בקובץ הוצאות חסרה עמודת 'תאריך'")

    expenses_df['תאריך'] = pd.to_datetime(expenses_df['תאריך'], errors='coerce')
    expenses_df = expenses_df[expenses_df['תאריך'].dt.month == month].copy()

    # מציאת עמודת מחיר בפועל
    price_col = None
    for cand in ["מחיר לפני מע\"מ", "מחיר לפני מע״מ", "מחיר_לפני_מע\"מ", "מחיר"]:
        if cand in expenses_df.columns:
            price_col = cand
            break
    if price_col is None:
        raise ValueError("בקובץ הוצאות חסרה עמודת 'מחיר לפני מע\"מ'")

    if 'מקט' not in expenses_df.columns:
        raise ValueError("בקובץ הוצאות חסרה עמודת 'מקט'")

    # מיזוג לפי מק\"ט
    merged = expenses_df.merge(pricing_df, on='מקט', how='left', suffixes=('_expenses', '_pricing'))

    # סטטוס (לוגיקה מקורית, ללא טולרנס)
    def status_row(row):
        p_list = row['מחירון']
        p_actual = row[price_col]
        if pd.isna(p_list):
            return '🟡 לא נמצא במחירון'
        if pd.isna(p_actual):
            return '⚠️ מחיר בפועל חסר'
        return '✅ תואם' if abs(p_list - p_actual) <= 0.00 else '❌ מחיר שונה'

    merged['סטטוס'] = merged.apply(status_row, axis=1)

    # -------- העמודה היחידה הנוספת: שינוי בשקלים --------
    merged['שינוי בשקלים'] = merged[price_col] - merged['מחירון']

    # שומרים את סדר העמודות המקורי; רק מבטיחים שהעמודה החדשה בסוף
    cols = list(merged.columns)
    if cols[-1] != 'שינוי בשקלים':
        cols = [c for c in cols if c != 'שינוי בשקלים'] + ['שינוי בשקלים']
    ordered = merged[cols]

    # -------- יצוא ל-Excel עם עיצוב קל, AutoFilter על כל העמודות --------
    out = BytesIO()
    with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
        ordered.to_excel(writer, index=False, sheet_name='השוואה')
        wb = writer.book
        ws = writer.sheets['השוואה']

        # כותרות
        header_fmt = wb.add_format({'bold': True, 'bg_color': '#F1F5F9', 'border': 1})
        for col_idx, col_name in enumerate(ordered.columns):
            ws.write(0, col_idx, col_name, header_fmt)
            ws.set_column(col_idx, col_idx, 16)

        # Freeze header + AutoFilter לכל הטווח
        n_rows, n_cols = ordered.shape
        ws.freeze_panes(1, 0)
        ws.autofilter(f"A1:{xl_col_to_name_local(n_cols - 1)}{n_rows + 1}")

        # עיצוב מותנה לכל השורה לפי סטטוס (לקריאות באקסל ווינדוס)
        data_range = f"A2:{xl_col_to_name_local(n_cols - 1)}{n_rows + 1}"
        green = wb.add_format({'bg_color': '#DCFCE7'})
        red   = wb.add_format({'bg_color': '#FEE2E2'})
        yellow= wb.add_format({'bg_color': '#FEF9C3'})
        ws.conditional_format(data_range, {'type': 'text', 'criteria': 'containing', 'value': '✅', 'format': green})
        ws.conditional_format(data_range, {'type': 'text', 'criteria': 'containing', 'value': '❌', 'format': red})
        ws.conditional_format(data_range, {'type': 'text', 'criteria': 'containing', 'value': '🟡', 'format': yellow})

    out.seek(0)
    return ordered, out, price_col

# ---------- Build WhatsApp message (only shown on button click) ----------
def build_message(df: pd.DataFrame, price_col: str, contact_name: str):
    diffs = df[df['סטטוס'] == '❌ מחיר שונה'].copy()
    if diffs.empty:
        return "אין הבדלים בין המחירון למה ששילמת בפועל. ✅"

    name_part = contact_name.strip() if contact_name.strip() else "[שם]"
    lines = [
        f"היי {name_part},",
        "יש הבדל בין המחירון למה ששילמתי בפועל, הנה הרשימה:",
        "מוצר | מחירון | מה ששילמתי בפועל",
    ]

    # נזהה עמודת פריט להצגה
    prod_col = 'פריט_expenses' if 'פריט_expenses' in diffs.columns else ('פריט' if 'פריט' in diffs.columns else None)

    def as_price(x):
        if pd.isna(x):
            return ""
        return f'{round(float(x), 2)} ש"ח'

    for _, row in diffs.iterrows():
        prod_name = row[prod_col] if (prod_col and prod_col in row) else ""
        list_price = as_price(row['מחירון'])
        actual = as_price(row[price_col])
        lines.append(f"{prod_name} | {list_price} | {actual}")

    return "\n".join(lines)

# ---------- UI flow ----------
if pricing_file and expense_file:
    if st.button("🔍 השווה עכשיו"):
        try:
            result_df, result_excel, actual_price_col = compare_files(pricing_file, expense_file, month=6)
            st.success("✔️ ההשוואה הושלמה.")
            st.dataframe(result_df, use_container_width=True)
            st.download_button(
                "📥 הורד את הדוח (Excel)",
                data=result_excel,
                file_name="comparison.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # ---- WhatsApp message: hidden by default, shown only when clicking a button ----
            if 'show_msg' not in st.session_state:
                st.session_state.show_msg = False

            col_a, col_b = st.columns([1, 3])
            if col_a.button("💬 הצג הודעה מוכנה לווטסאפ"):
                st.session_state.show_msg = True

            if st.session_state.show_msg:
                contact_name = col_b.text_input("שם הנמען (לא חובה)", value="")
                msg = build_message(result_df, actual_price_col, contact_name)
                st.code(msg, language=None)
                # אופציונלי: קישור וואטסאפ פתיחה עם הטקסט
                # wa_link = "https://wa.me/?text=" + urlquote(msg)
                # st.markdown(f"[פתח וואטסאפ עם ההודעה]({wa_link})")

        except Exception as e:
            st.error(f"❌ שגיאה בעת עיבוד הקבצים: {e}")
else:
    st.info("⚠️ יש להעלות את שני הקבצים כדי לבצע השוואה.")
