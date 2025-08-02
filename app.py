# ================== סינון מהיר ==================
st.subheader("🔎 סינון מהיר")

df_filtered = result_df.copy()

# זיהוי שמות עמודות אפשריים
serial_candidates = ['מקט', 'מק״ט', 'מק\'ט']
serial_col = next((c for c in serial_candidates if c in df_filtered.columns), None)
prod_display_col = product_col if product_col and product_col in df_filtered.columns else ('פריט' if 'פריט' in df_filtered.columns else None)

col1, col2, col3 = st.columns(3)
col4, col5, col6 = st.columns(3)

# לקוח (רב-בחירה)
if 'לקוח' in df_filtered.columns:
    clients = sorted(df_filtered['לקוח'].dropna().astype(str).unique().tolist())
    chosen_clients = col1.multiselect("לקוח", options=clients, default=[])
    if chosen_clients:
        df_filtered = df_filtered[df_filtered['לקוח'].astype(str).isin(chosen_clients)]

# תעודה (רב-בחירה)
if 'תעודה' in df_filtered.columns:
    docs = sorted(df_filtered['תעודה'].dropna().astype(str).unique().tolist())
    chosen_docs = col2.multiselect("תעודה", options=docs, default=[])
    if chosen_docs:
        df_filtered = df_filtered[df_filtered['תעודה'].astype(str).isin(chosen_docs)]

# סטטוס (רב-בחירה)
if 'סטטוס' in df_filtered.columns:
    statuses = ['✅ תואם', '❌ מחיר שונה', '🟡 לא נמצא במחירון', '⚠️ מחיר בפועל חסר']
    existing_statuses = [s for s in statuses if s in df_filtered['סטטוס'].unique().tolist()]
    chosen_status = col3.multiselect("סטטוס", options=existing_statuses, default=[])
    if chosen_status:
        df_filtered = df_filtered[df_filtered['סטטוס'].isin(chosen_status)]

# תאריך (טווח)
if 'תאריך' in df_filtered.columns and pd.api.types.is_datetime64_any_dtype(df_filtered['תאריך']):
    min_date = df_filtered['תאריך'].min().date() if not df_filtered['תאריך'].isna().all() else None
    max_date = df_filtered['תאריך'].max().date() if not df_filtered['תאריך'].isna().all() else None
    if min_date and max_date:
        date_range = col4.date_input("טווח תאריכים (תאריך)", value=(min_date, max_date))
        if isinstance(date_range, tuple) and len(date_range) == 2:
            start_date, end_date = date_range
            df_filtered = df_filtered[(df_filtered['תאריך'].dt.date >= start_date) & (df_filtered['תאריך'].dt.date <= end_date)]

# אספקה (טווח) — אם קיימת
if 'אספקה' in df_filtered.columns:
    # המרה לתאריך אם צריך
    if not pd.api.types.is_datetime64_any_dtype(df_filtered['אספקה']):
        df_filtered['אספקה'] = pd.to_datetime(df_filtered['אספקה'], errors='coerce')
    if not df_filtered['אספקה'].isna().all():
        min_sup = df_filtered['אספקה'].min().date()
        max_sup = df_filtered['אספקה'].max().date()
        sup_range = col5.date_input("טווח תאריכים (אספקה)", value=(min_sup, max_sup))
        if isinstance(sup_range, tuple) and len(sup_range) == 2:
            start_sup, end_sup = sup_range
            df_filtered = df_filtered[(df_filtered['אספקה'].dt.date >= start_sup) & (df_filtered['אספקה'].dt.date <= end_sup)]

# מק״ט (רב-בחירה) — אם קיים
if serial_col:
    serials = df_filtered[serial_col].dropna().astype(str).unique().tolist()
    chosen_serials = col6.multiselect("מק״ט", options=sorted(serials), default=[])
    if chosen_serials:
        df_filtered = df_filtered[df_filtered[serial_col].astype(str).isin(chosen_serials)]

# פריט — חיפוש טקסט (מכיל)
if prod_display_col:
    query = st.text_input("חיפוש לפי פריט (מכיל)")
    if query.strip():
        df_filtered = df_filtered[df_filtered[prod_display_col].astype(str).str.contains(query.strip(), case=False, na=False)]

# ================== הצגה + הורדה לפי הסינון ==================
st.success(f"נמצאו {len(df_filtered)} שורות לאחר הסינון.")
st.dataframe(df_filtered, use_container_width=True)

# ייצוא אקסל עם עיצוב (אותו סגנון כמו קודם)
from io import BytesIO
def export_with_format(df):
    out = BytesIO()
    with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='השוואה')
        workbook = writer.book
        ws = writer.sheets['השוואה']

        # כותרות
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#F1F5F9', 'border': 1})
        for col_idx, col_name in enumerate(df.columns):
            ws.write(0, col_idx, col_name, header_fmt)
            ws.set_column(col_idx, col_idx, 16)

        # עיצוב מותנה לפי סטטוס (אם קיים)
        if 'סטטוס' in df.columns:
            # נקבע טווח כלל השורות (צבע לכל השורה)
            n_rows, n_cols = df.shape
            last_col = n_cols - 1
            # טווח נתונים (A2 עד עמודה אחרונה)
            from xlsxwriter.utility import xl_col_to_name
            data_range = f"A2:{xl_col_to_name(last_col)}{n_rows+1}"

            green = workbook.add_format({'bg_color': '#DCFCE7'})
            red   = workbook.add_format({'bg_color': '#FEE2E2'})
            yellow= workbook.add_format({'bg_color': '#FEF9C3'})

            ws.conditional_format(data_range, {'type':'text', 'criteria':'containing', 'value':'✅', 'format':green})
            ws.conditional_format(data_range, {'type':'text', 'criteria':'containing', 'value':'❌', 'format':red})
            ws.conditional_format(data_range, {'type':'text', 'criteria':'containing', 'value':'🟡', 'format':yellow})

    out.seek(0)
    return out

excel_filtered = export_with_format(df_filtered)

st.download_button(
    "📥 הורד את הדוח (לפי הסינון)",
    data=excel_filtered,
    file_name="comparison_filtered.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ================== טמפלייט הודעה לפי הסינון ==================
st.subheader("📝 טמפלייט הודעה (לפי הסינון)")
msg = build_message(df_filtered, actual_col, prod_display_col, contact_name)
st.code(msg, language=None)
