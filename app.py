import streamlit as st
import pandas as pd
from io import BytesIO
from xlsxwriter.utility import xl_col_to_name

# ===================== הגדרות עמוד ועיצוב =====================
st.set_page_config(page_title="FreshTrack Analytics", layout="centered")

st.markdown("""
    <style>
    body { background: linear-gradient(to bottom right, #fef3c7, #d1fae5, #dbeafe); }
    .title-text { font-size: 3rem; font-weight: bold;
        background: linear-gradient(to right, #16a34a, #2563eb, #7c3aed);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }
    .subtitle-text { font-size: 1.2rem; color: #374151; margin-bottom: 1rem; }
    .upload-box { border: 4px dashed #6ee7b7; padding: 1rem; border-radius: 1rem; background-color: #ecfdf5; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align:center' class='title-text'>🚀 FreshTrack Analytics</h1>", unsafe_allow_html=True)
st.markdown("<p class='subtitle-text' style='text-align:center'>השווה מחירון מול הוצאות במהירות 🥦🍓</p>", unsafe_allow_html=True)

# ===================== קלטים כלליים =====================
contact_name = st.text_input("שם הנמען בהודעה (אופציונלי)", value="")

st.markdown("<div class='upload-box'>📥 העלה קובץ מחירון (Excel: מק\"ט, פריט, מחירון)</div>", unsafe_allow_html=True)
pricing_file = st.file_uploader("מחירון", type=["xls", "xlsx"], key="pricing")

st.markdown("<div class='upload-box'>📥 העלה קובץ הוצאות (Excel: תאריך, מק\"ט, פריט, כמות, מחיר לפני מע\"מ)</div>", unsafe_allow_html=True)
expense_file = st.file_uploader("הוצאות", type=["xls", "xlsx"], key="expense")


# ===================== עזר: פורמט סימן מספרי =====================
def format_signed_number(x, decimals=2):
    if pd.isna(x):
        return ""
    x = round(float(x), decimals)
    sign = "+" if x > 0 else ("" if x == 0 else "")
    return f"{sign}{x:.{decimals}f}"


# ===================== לוגיקת השוואה =====================
def compare_files(pricing_file, expense_file, month: int = 6):
    # קריאת קבצים
    pricing_df = pd.read_excel(pricing_file)
    expenses_df = pd.read_excel(expense_file)

    # מחירון: 3 עמודות (מקט, פריט, מחירון)
    pricing_df.columns = ['מקט', 'פריט', 'מחירון']

    # הוצאות: דרושות 'תאריך', 'מקט' ועמודת מחיר בפועל
    if 'תאריך' not in expenses_df.columns:
        raise ValueError("בקובץ הוצאות חסרה עמודת 'תאריך'")

    expenses_df['תאריך'] = pd.to_datetime(expenses_df['תאריך'], errors='coerce')
    expenses_df = expenses_df[expenses_df['תאריך'].dt.month == month].copy()

    price_col = None
    for cand in ["מחיר לפני מע\"מ", "מחיר לפני מע״מ", "מחיר_לפני_מע\"מ", "מחיר"]:
        if cand in expenses_df.columns:
            price_col = cand
            break
    if price_col is None:
        raise ValueError("בקובץ הוצאות חסרה עמודת 'מחיר לפני מע\"מ'")

    if 'מקט' not in expenses_df.columns:
        raise ValueError("בקובץ הוצאות חסרה עמודת 'מקט'")

    merged = expenses_df.merge(pricing_df, on='מקט', how='left', suffixes=('_expenses', '_pricing'))

    # סטטוס וסטיית מחיר (לא משנים את הלוגיקה העסקית)
    def status_row(row):
        p_list = row['מחירון']
        p_actual = row[price_col]
        if pd.isna(p_list):
            return '🟡 לא נמצא במחירון'
        if pd.isna(p_actual):
            return '⚠️ מחיר בפועל חסר'
        return '✅ תואם' if abs(p_list - p_actual) <= 0.00 else '❌ מחיר שונה'

    merged['סטטוס'] = merged.apply(status_row, axis=1)
    merged['פער מחיר'] = merged[price_col] - merged['מחירון']           # מספרי
    merged['סטיית מחיר'] = merged['פער מחיר'].apply(lambda v: format_signed_number(v, 2))  # מחרוזת עם סימן

    prod_col = 'פריט_expenses' if 'פריט_expenses' in merged.columns else ('פריט' if 'פריט' in merged.columns else None)

    preferred_cols = [c for c in [
        'לקוח', 'תעודה', 'תאריך', 'אספקה',
        'מקט', prod_col, 'כמות',
        price_col, 'מחירון', 'פער מחיר', 'סטיית מחיר', 'סטטוס'
    ] if c and c in merged.columns]

    ordered = merged[preferred_cols + [c for c in merged.columns if c not in preferred_cols]]

    # יצוא אקסל מלא עם עיצוב מותנה לפי סטטוס
    out_full = BytesIO()
    with pd.ExcelWriter(out_full, engine='xlsxwriter') as writer:
        ordered.to_excel(writer, index=False, sheet_name='השוואה')
        wb = writer.book
        ws = writer.sheets['השוואה']

        # כותרות
        header_fmt = wb.add_format({'bold': True, 'bg_color': '#F1F5F9', 'border': 1})
        for col_idx, col_name in enumerate(ordered.columns):
            ws.write(0, col_idx, col_name, header_fmt)
            ws.set_column(col_idx, col_idx, 16)

        # אוטו-פילטר והקפאת שורה
        n_rows, n_cols = ordered.shape
        ws.autofilter(f"A1:{xl_col_to_name(n_cols - 1)}{n_rows + 1}")
        ws.freeze_panes(1, 0)

        # עיצוב מותנה לכל השורה לפי סטטוס
        data_range = f"A2:{xl_col_to_name(n_cols - 1)}{n_rows + 1}"
        green = wb.add_format({'bg_color': '#DCFCE7'})
        red   = wb.add_format({'bg_color': '#FEE2E2'})
        yellow= wb.add_format({'bg_color': '#FEF9C3'})
        ws.conditional_format(data_range, {'type': 'text', 'criteria': 'containing', 'value': '✅', 'format': green})
        ws.conditional_format(data_range, {'type': 'text', 'criteria': 'containing', 'value': '❌', 'format': red})
        ws.conditional_format(data_range, {'type': 'text', 'criteria': 'containing', 'value': '🟡', 'format': yellow})

    out_full.seek(0)
    return ordered, out_full, price_col, prod_col


# ===================== טמפלייט הודעה =====================
def build_message(df, price_col, prod_col, contact_name: str):
    diffs = df[df['סטטוס'] == '❌ מחיר שונה'].copy()
    if diffs.empty:
        return "אין הבדלים בין המחירון למה ששילמת בפועל. ✅"

    name_part = contact_name.strip() if contact_name.strip() else "[שם]"
    lines = [
        f"היי {name_part},",
        "יש הבדל בין המחירון למה ששילמתי בפועל, הנה הרשימה:",
        "מוצר | מחירון | מה ששילמתי בפועל",
    ]

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


# ===================== זרימת האפליקציה =====================
if pricing_file and expense_file:
    if st.button("🔍 השווה עכשיו"):
        try:
            result_df, excel_full, actual_col, product_col = compare_files(pricing_file, expense_file, month=6)

            # -------- סינון מהיר (על פי דרישתך) --------
            st.subheader("🔎 סינון מהיר")
            df_filtered = result_df.copy()

            # מזהי עמודות
            serial_candidates = ['מקט', 'מק״ט', "מק'ט"]
            serial_col = next((c for c in serial_candidates if c in df_filtered.columns), None)
            prod_display_col = product_col if (product_col and product_col in df_filtered.columns) else ('פריט' if 'פריט' in df_filtered.columns else None)

            col1, col2, col3 = st.columns(3)
            col4, col5, col6 = st.columns(3)

            # לקוח
            if 'לקוח' in df_filtered.columns:
                clients = sorted(df_filtered['לקוח'].dropna().astype(str).unique().tolist())
                chosen_clients = col1.multiselect("לקוח", options=clients, default=[])
                if chosen_clients:
                    df_filtered = df_filtered[df_filtered['לקוח'].astype(str).isin(chosen_clients)]

            # תעודה
            if 'תעודה' in df_filtered.columns:
                docs = sorted(df_filtered['תעודה'].dropna().astype(str).unique().tolist())
                chosen_docs = col2.multiselect("תעודה", options=docs, default=[])
                if chosen_docs:
                    df_filtered = df_filtered[df_filtered['תעודה'].astype(str).isin(chosen_docs)]

            # סטטוס
            if 'סטטוס' in df_filtered.columns:
                statuses = ['✅ תואם', '❌ מחיר שונה', '🟡 לא נמצא במחירון', '⚠️ מחיר בפועל חסר']
                existing_statuses = [s for s in statuses if s in df_filtered['סטטוס'].unique().tolist()]
                chosen_status = col3.multiselect("סטטוס", options=existing_statuses, default=[])
                if chosen_status:
                    df_filtered = df_filtered[df_filtered['סטטוס'].isin(chosen_status)]

            # תאריך (טווח)
            if 'תאריך' in df_filtered.columns and pd.api.types.is_datetime64_any_dtype(df_filtered['תאריך']):
                if not df_filtered['תאריך'].isna().all():
                    min_date = df_filtered['תאריך'].min().date()
                    max_date = df_filtered['תאריך'].max().date()
                    date_range = col4.date_input("טווח תאריכים (תאריך)", value=(min_date, max_date))
                    if isinstance(date_range, tuple) and len(date_range) == 2:
                        start_date, end_date = date_range
                        df_filtered = df_filtered[(df_filtered['תאריך'].dt.date >= start_date) & (df_filtered['תאריך'].dt.date <= end_date)]

            # אספקה (טווח)
            if 'אספקה' in df_filtered.columns:
                if not pd.api.types.is_datetime64_any_dtype(df_filtered['אספקה']):
                    df_filtered['אספקה'] = pd.to_datetime(df_filtered['אספקה'], errors='coerce')
                if not df_filtered['אספקה'].isna().all():
                    min_sup = df_filtered['אספקה'].min().date()
                    max_sup = df_filtered['אספקה'].max().date()
                    sup_range = col5.date_input("טווח תאריכים (אספקה)", value=(min_sup, max_sup))
                    if isinstance(sup_range, tuple) and len(sup_range) == 2:
                        start_sup, end_sup = sup_range
                        df_filtered = df_filtered[(df_filtered['אספקה'].dt.date >= start_sup) & (df_filtered['אספקה'].dt.date <= end_sup)]

            # מק״ט
            if serial_col:
                serials = sorted(df_filtered[serial_col].dropna().astype(str).unique().tolist())
                chosen_serials = col6.multiselect("מק״ט", options=serials, default=[])
                if chosen_serials:
                    df_filtered = df_filtered[df_filtered[serial_col].astype(str).isin(chosen_serials)]

            # פריט (חיפוש טקסט חופשי)
            if prod_display_col:
                query = st.text_input("חיפוש לפי פריט (מכיל)")
                if query.strip():
                    df_filtered = df_filtered[df_filtered[prod_display_col].astype(str).str.contains(query.strip(), case=False, na=False)]

            # -------- הצגה + הורדה --------
            st.success(f"נמצאו {len(df_filtered)} שורות לאחר הסינון.")
            st.dataframe(df_filtered, use_container_width=True)

            # הורדה - דוח מלא (ללא סינון)
            st.download_button(
                "📥 הורד את הדוח המלא (Excel)",
                data=excel_full,
                file_name="comparison.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # הורדה - דוח לפי סינון (עם עיצוב)
            def export_with_format(df):
                out = BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='השוואה')
                    wb = writer.book
                    ws = writer.sheets['השוואה']

                    header_fmt = wb.add_format({'bold': True, 'bg_color': '#F1F5F9', 'border': 1})
                    for col_idx, col_name in enumerate(df.columns):
                        ws.write(0, col_idx, col_name, header_fmt)
                        ws.set_column(col_idx, col_idx, 16)

                    n_rows, n_cols = df.shape
                    ws.autofilter(f"A1:{xl_col_to_name(n_cols - 1)}{n_rows + 1}")
                    ws.freeze_panes(1, 0)

                    if 'סטטוס' in df.columns:
                        data_range = f"A2:{xl_col_to_name(n_cols - 1)}{n_rows + 1}"
                        green = wb.add_format({'bg_color': '#DCFCE7'})
                        red   = wb.add_format({'bg_color': '#FEE2E2'})
                        yellow= wb.add_format({'bg_color': '#FEF9C3'})
                        ws.conditional_format(data_range, {'type': 'text', 'criteria': 'containing', 'value': '✅', 'format': green})
                        ws.conditional_format(data_range, {'type': 'text', 'criteria': 'containing', 'value': '❌', 'format': red})
                        ws.conditional_format(data_range, {'type': 'text', 'criteria': 'containing', 'value': '🟡', 'format': yellow})
                out.seek(0)
                return out

            excel_filtered = export_with_format(df_filtered)
            st.download_button(
                "📥 הורד את הדוח (לפי הסינון)",
                data=excel_filtered,
                file_name="comparison_filtered.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # -------- טמפלייט הודעה --------
            st.subheader("📝 טמפלייט הודעה (לפי הסינון)")
            msg = build_message(df_filtered, actual_col, product_col, contact_name)
            st.code(msg, language=None)

        except Exception as e:
            st.error(f"❌ שגיאה בעת עיבוד הקבצים: {e}")
else:
    st.info("⚠️ יש להעלות את שני הקבצים כדי לבצע השוואה.")
