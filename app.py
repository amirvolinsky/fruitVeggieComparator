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
    משווה מחיר בפועל מול מחיר מחירון תוך התחשבות במק"ט ובתאריך.

    מהלך הפעולה:
    1. מאתר את עמודות המק"ט והתאריך בשני הקבצים וממפה אותן לשמות אחידים.
    2. מבצע מיזוג לפי מק"ט ותאריך כך שהשוואה תתבצע מול המחיר הרלוונטי לתאריך ההוצאה.
    3. מאתר את עמודות המחיר בפועל במסמך ההוצאות (למשל "מחיר לפני מע"מ" ו"מחיר לאחר מע"מ").
       במקרה שיש יותר מעמודה אחת, מנסה להתאים את המחיר במחירון לעמודה שתואמת את המחיר בפועל לאחר מע"מ (כיוון שבקבצי המחירון המחיר כולל מע"מ).
       אם המחיר במחירון תואם לאחד העמודות — הסטטוס יסומן כ"✅ תואם" והפרש המחיר יחושב ביחס לאותה עמודה.
       אם אין התאמה מלאה, נבחר את העמודה בעלת הפרש מינימלי ונחשב את הסטטוס כ"❌ לא תואם".
    4. מחזיר DataFrame עם עמודת סטטוס ועמודת "שוני במחיר" (הפרש המחיר בפועל מהמחירון).

    שים לב: אם המחיר במחירון חסר (NaN) עבור המקט/תאריך – הסטטוס יסומן כ"🟡 חסר במחירון" והפרש המחיר יהיה ריק.
    """
    # עמודת מק"ט בשני הקבצים
    serial_col_exp = _find_column(expenses, ["מקט"])
    serial_col_price = _find_column(price_list, ["מקט"])
    if serial_col_exp is None or serial_col_price is None:
        raise KeyError("לא נמצאה עמודת מק\"ט באחד הקבצים. ודאו שקיימת עמודת מק\"ט.")

    # עמודת מחירון (עשויה לכלול רווח בסוף)
    if 'מחירון ' in price_list.columns:
        price_col_list = 'מחירון '
    elif 'מחירון' in price_list.columns:
        price_col_list = 'מחירון'
    else:
        price_col_list = _find_column(price_list, ["מחירון", "מחיר"])
    if price_col_list is None:
        raise KeyError("לא נמצאה עמודת מחיר במחירון. ודאו שקיימת עמודת מחירון או עמודת מחיר.")

    # עמודות תאריך בשני הקבצים
    date_col_price = _find_column(price_list, ["תאריך", "date"])
    date_col_exp = _find_column(expenses, ["תאריך", "date"])
    if date_col_price is None or date_col_exp is None:
        raise KeyError("לא נמצאה עמודת תאריך באחד הקבצים.")

    # איחוד שמות לשימוש פנימי בלבד
    price_list_renamed = price_list.rename(columns={
        price_col_list: 'מחיר מחירון',
        serial_col_price: 'מקט',
        date_col_price: 'תאריך'
    })
    expenses_renamed = expenses.rename(columns={
        serial_col_exp: 'מקט',
        date_col_exp: 'תאריך'
    })

    # המרה לתאריכים (לאובייקט date בלבד כדי שהשוואה תהיה ללא זמן)
    price_list_renamed['תאריך'] = pd.to_datetime(price_list_renamed['תאריך'], errors='coerce').dt.date
    expenses_renamed['תאריך'] = pd.to_datetime(expenses_renamed['תאריך'], errors='coerce').dt.date

    # מיזוג לפי מק"ט ותאריך
    merged = pd.merge(
        expenses_renamed,
        price_list_renamed[['מקט', 'תאריך', 'מחיר מחירון']],
        on=['מקט', 'תאריך'],
        how='left'
    )

    # מציאת כל עמודות המחיר הרלוונטיות במסמך ההוצאות
    # נבדוק גם שמות בעברית לפני/אחרי מע"מ וגם שמות מסוננים
    def _find_price_columns(columns: list[str]) -> list[str]:
        candidates = []
        for candidate in [
            'מחיר לפני מע"מ', 'מחיר לפני מע"מ', 'מחיר לפני מע\"מ',
            'מחיר לאחר מע"מ', 'מחיר לאחר מע\"מ'
        ]:
            if candidate in columns:
                candidates.append(candidate)
        # הוספת עמודות מסוננות שעלולות להכיל את המילים הללו
        for col in columns:
            sanitized = _sanitize_column_name(col)
            if 'מחירלפנימעם' in sanitized or 'מחירלאחרמעם' in sanitized:
                if col not in candidates:
                    candidates.append(col)
        return candidates

    price_columns = _find_price_columns(list(expenses_renamed.columns))
    # אם לא מצאנו עמודות מחיר, נשתמש במשתנה ריק בשביל תנאי מאוחר
    if not price_columns:
        price_columns = []

    # קביעת סטטוס והפרש מחיר בהתבסס על העמודות שנמצאו
    def determine_status_and_diff(row):
        expected = row['מחיר מחירון']
        # אם אין מחיר במחירון, נציין כחסר
        if pd.isna(expected):
            return '🟡 חסר במחירון', pd.NA
        # אסוף את המחירים בפועל מהרשימה
        actual_values = []
        for col in price_columns:
            # במידה והעמודה לא קיימת בשורה (למשל אחרי merge), נמשיך
            if col in row.index:
                actual_values.append(row[col])
        # המרת המחירים למספרים
        numeric_actuals = []
        for val in actual_values:
            try:
                numeric_actuals.append(float(val))
            except Exception:
                numeric_actuals.append(float('nan'))
        expected_num = float(expected)
        # אם אחת מהעמודות תואמת בדיוק את המחירון, נגדיר כתואם עם הפרש 0
        for act in numeric_actuals:
            # התאמה תוך התעלמות מהפרש זעיר (בעיות חישוב עשרוני)
            if pd.notna(act) and abs(act - expected_num) < 1e-6:
                return '✅ תואם', act - expected_num
        # אחרת נמצא את העמודה עם הפרש מינימלי, אם קיימת
        if numeric_actuals:
            # רשימת הפרשים בין מחירים בפועל למחירון; אם הערך לא מספרי ניקח אינסוף כדי לא לבחור בו
            diffs = [act - expected_num if pd.notna(act) else float('inf') for act in numeric_actuals]
            # מציאת הפרש מינימלי בערכים סופיים בלבד
            import math
            valid_diffs = [d for d in diffs if not math.isinf(d)]
            if valid_diffs:
                min_diff = min(valid_diffs, key=lambda x: abs(x))
                return '❌ לא תואם', min_diff
        # אם אין עמודות מחיר תקינות, נחזיר סטטוס לא תואם והפרש לא ידוע
        return '❌ לא תואם', pd.NA

    # החלת הפונקציה על כל השורות
    status_list = []
    diff_list = []
    for _, row in merged.iterrows():
        status, diff = determine_status_and_diff(row)
        status_list.append(status)
        diff_list.append(diff)
    merged['סטאטוס'] = status_list
    merged['שוני במחיר'] = diff_list

    # דואגים שהעמודה החדשה תהיה בסוף
    cols = list(merged.columns)
    # אם סדר העמודות אינו מסתיים ב'שוני במחיר', נסדר אותו
    if cols[-1] != 'שוני במחיר':
        cols = [c for c in cols if c not in ['סטאטוס', 'שוני במחיר']] + ['סטאטוס', 'שוני במחיר']
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
