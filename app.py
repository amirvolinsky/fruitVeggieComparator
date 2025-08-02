import streamlit as st
import pandas as pd
from io import BytesIO

# Helper: המרת אינדקס עמודה (0-based) לאות עמודת אקסל (A, B, ..., AA)
def _xl_col_to_name(col_idx: int) -> str:
    col_idx += 1
    name = ""
    while col_idx > 0:
        col_idx, rem = divmod(col_idx - 1, 26)
        name = chr(65 + rem) + name
    return name


def load_excel(file: BytesIO) -> pd.DataFrame:
    """Load an Excel file into a pandas DataFrame."""
    return pd.read_excel(file)


def filter_june_expenses(expenses: pd.DataFrame) -> pd.DataFrame:
    """Filter the expense report DataFrame to include only rows from June."""
    df = expenses.copy()
    if 'תאריך' not in df.columns:
        return pd.DataFrame()

    df['תאריך'] = pd.to_datetime(df['תאריך'], errors='coerce')
    df = df[df['תאריך'].dt.month == 6]
    return df


def _sanitize_column_name(name: str) -> str:
    """Remove non-alphanumeric characters from a column name for robust matching."""
    import re
    return re.sub(r"[^A-Za-z\u0590-\u05FF0-9]", "", str(name))


def _find_column(df: pd.DataFrame, keywords: list[str]) -> str | None:
    """Find a column whose sanitized name contains any of the given keywords."""
    sanitized_keywords = [_sanitize_column_name(k) for k in keywords]
    for col in df.columns:
        sanitized_col = _sanitize_column_name(col)
        for key in sanitized_keywords:
            if key in sanitized_col:
                return col
    return None


def compare_prices(expenses: pd.DataFrame, price_list: pd.DataFrame) -> pd.DataFrame:
    """
    משווה מחיר בפועל מול מחיר מחירון.
    לא משנה שמות/מבנה עמודות פרט לשינוי שם 'מחירון' ל'מחיר מחירון' וקונסולידציית 'מקט'.
    מוסיף עמודה אחת בלבד: 'שוני במחיר' = מחיר בפועל - מחיר מחירון.
    """
    # עמודת מק"ט בשני הקבצים
    serial_col_exp = _find_column(expenses, ["מקט"])
    serial_col_price = _find_column(price_list, ["מקט"])
    if serial_col_exp is None or serial_col_price is None:
        raise KeyError("לא נמצאה עמודת מק\"ט באחד הקבצים. ודאו שקיימת עמודת מק\"ט.")

    # עמודת מחיר במחירון (מעדיפים 'מחירון')
    price_col_list = 'מחירון' if 'מחירון' in price_list.columns else _find_column(price_list, ["מחירון", "מחיר"])
    if price_col_list is None:
        raise KeyError("לא נמצאה עמודת מחיר במחירון. ודאו שקיימת עמודת מחירון או עמודת מחיר.")

    # איחוד שמות לשימוש פנימי בלבד
    price_list_renamed = price_list.rename(columns={price_col_list: 'מחיר מחירון', serial_col_price: 'מקט'})
    expenses_renamed = expenses.rename(columns={serial_col_exp: 'מקט'})

    # מיזוג לפי מק"ט
    merged = pd.merge(
        expenses_renamed,
        price_list_renamed[['מקט', 'מחיר מחירון']],
        on='מקט',
        how='left'
    )

    # סטטוס (ללא שינוי לוגיקה)
    def determine_status(row):
        expected = row['מחיר מחירון']
        actual = row.get('מחיר לפני מע"מ')
        if pd.isna(expected):
            return '🟡 חסר במחירון'
        try:
            if float(actual) == float(expected):
                return '✅ תואם'
            else:
                return '❌ לא תואם'
        except Exception:
            return '❌ לא תואם'

    merged['סטאטוס'] = merged.apply(determine_status, axis=1)

    # -------- עמודה יחידה: שוני במחיר (מספרי) --------
    if 'מחיר לפני מע"מ' in merged.columns:
        actual_num = pd.to_numeric(merged['מחיר לפני מע"מ'], errors='coerce')
        list_num = pd.to_numeric(merged['מחיר מחירון'], errors='coerce')
        merged['שוני במחיר'] = actual_num - list_num
    else:
        merged['שוני במחיר'] = pd.NA

    # לא משנים סדר עמודות, רק מקפידים שהעמודה החדשה תהיה בסוף
    cols = list(merged.columns)
    if cols[-1] != 'שוני במחיר':
        cols = [c for c in cols if c != 'שוני במחיר'] + ['שוני במחיר']
    merged = merged[cols]

    return merged


def create_downloadable_excel(df: pd.DataFrame) -> bytes:
    """
    יוצר קובץ אקסל:
    - כותרות מעוצבות
    - Freeze Panes לכותרות
    - AutoFilter על כל העמודות
    - צביעה מותנית לפי 'סטאטוס' (✅ ירוק, ❌ אדום, 🟡 צהוב)
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        sheet_name = 'Comparison'
        df.to_excel(writer, index=False, sheet_name=sheet_name)

        wb = writer.book
        ws = writer.sheets[sheet_name]

        # כותרות
        header_fmt = wb.add_format({'bold': True, 'bg_color': '#F1F5F9', 'border': 1})
        for c_idx, col_name in enumerate(df.columns):
            ws.write(0, c_idx, col_name, header_fmt)
            ws.set_column(c_idx, c_idx, 16)

        # Freeze header
        n_rows, n_cols = df.shape
        ws.freeze_panes(1, 0)

        # AutoFilter לכל העמודות
        ws.autofilter(f"A1:{_xl_col_to_name(n_cols - 1)}{n_rows + 1}")

        # עיצוב מותנה לפי 'סטאטוס' (אם העמודה קיימת)
        if 'סטאטוס' in df.columns:
            data_range = f"A2:{_xl_col_to_name(n_cols - 1)}{n_rows + 1}"
            green = wb.add_format({'bg_color': '#DCFCE7'})  # ✅
            red   = wb.add_format({'bg_color': '#FEE2E2'})  # ❌
            yellow= wb.add_format({'bg_color': '#FEF9C3'})  # 🟡

            ws.conditional_format(data_range, {'type': 'text', 'criteria': 'containing', 'value': '✅', 'format': green})
            ws.conditional_format(data_range, {'type': 'text', 'criteria': 'containing', 'value': '❌', 'format': red})
            ws.conditional_format(data_range, {'type': 'text', 'criteria': 'containing', 'value': '🟡', 'format': yellow})

    output.seek(0)
    return output.getvalue()


def main():
    st.set_page_config(page_title="השוואת מחירון והוצאות", page_icon="📊", layout="centered")

    st.title("השוואת מחירון מול דוח הוצאות")
    st.markdown(
        """
        ## ברוכים הבאים!
        העלו קובץ מחירון וקובץ הוצאות, בצעו השוואה והורידו קובץ תוצאה.
        """
    )

    price_file = st.file_uploader("העלה קובץ מחירון", type=["xlsx", "xls"], key="price")
    expense_file = st.file_uploader("העלה קובץ הוצאות", type=["xlsx", "xls"], key="expenses")

    if price_file is not None and expense_file is not None:
        if st.button("השווה עכשיו", key="compare_button"):
            try:
                price_df = load_excel(price_file)
                expenses_df = load_excel(expense_file)
            except Exception as e:
                st.error(f"אירעה שגיאה בקריאת הקבצים: {e}")
                return

            june_expenses = filter_june_expenses(expenses_df)
            if june_expenses.empty:
                st.warning("לא נמצאו הוצאות לחודש יוני.")
                return

            try:
                comparison_df = compare_prices(june_expenses, price_df)
            except Exception as e:
                st.error(f"שגיאה בהשוואה: {e}")
                return

            # הורדה לאקסל עם עיצוב + פילטרים
            excel_bytes = create_downloadable_excel(comparison_df)

            st.success("השוואה הושלמה בהצלחה!")
            st.dataframe(comparison_df, use_container_width=True)
            st.download_button(
                label="📥 הורד קובץ השוואה",
                data=excel_bytes,
                file_name="comparison.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.info("אנא העלו גם קובץ מחירון וגם קובץ הוצאות כדי להמשיך.")


if __name__ == "__main__":
    main()
