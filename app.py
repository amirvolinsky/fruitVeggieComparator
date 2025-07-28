import streamlit as st
import pandas as pd
from io import BytesIO


def load_excel(file: BytesIO) -> pd.DataFrame:
    """Load an Excel file into a pandas DataFrame.

    Parameters
    ----------
    file : BytesIO
        Uploaded file object from Streamlit file uploader.

    Returns
    -------
    pd.DataFrame
        DataFrame containing the contents of the first sheet of the Excel file.
    """
    return pd.read_excel(file)


def filter_june_expenses(expenses: pd.DataFrame) -> pd.DataFrame:
    """Filter the expense report DataFrame to include only rows from June.

    The function attempts to parse the 'תאריך' column to datetime. If parsing
    fails for some rows, those rows are removed before filtering to June.

    Parameters
    ----------
    expenses : pd.DataFrame
        DataFrame containing the expense report.

    Returns
    -------
    pd.DataFrame
        Filtered DataFrame containing only expenses from June (month == 6).
    """
    df = expenses.copy()
    if 'תאריך' not in df.columns:
        # If no date column exists, return empty DataFrame
        return pd.DataFrame()

    # Convert date column to datetime, coercing errors to NaT
    df['תאריך'] = pd.to_datetime(df['תאריך'], errors='coerce')
    # Filter to June rows (month == 6)
    df = df[df['תאריך'].dt.month == 6]
    return df


def _sanitize_column_name(name: str) -> str:
    """Remove non-alphanumeric characters from a column name for robust matching.

    Parameters
    ----------
    name : str
        Original column name.

    Returns
    -------
    str
        Sanitized column name containing only Hebrew/Latin letters and digits.
    """
    import re

    return re.sub(r"[^A-Za-z\u0590-\u05FF0-9]", "", str(name))


def _find_column(df: pd.DataFrame, keywords: list[str]) -> str | None:
    """Find a column in the DataFrame whose sanitized name contains any of the given keywords.

    This helper allows matching of column names that might include punctuation or quotes,
    which are removed by sanitization. For example, "מק\"ט" will sanitize to "מקט".

    Parameters
    ----------
    df : pd.DataFrame
        DataFrame to search.
    keywords : list[str]
        List of substrings to look for in sanitized column names.

    Returns
    -------
    str | None
        The first matching column name, or None if no match is found.
    """
    sanitized_keywords = [_sanitize_column_name(k) for k in keywords]
    for col in df.columns:
        sanitized_col = _sanitize_column_name(col)
        for key in sanitized_keywords:
            if key in sanitized_col:
                return col
    return None


def compare_prices(expenses: pd.DataFrame, price_list: pd.DataFrame) -> pd.DataFrame:
    """Compare actual charged prices against expected prices from the price list.

    This function attempts to be flexible with column names by matching columns
    based on key substrings (e.g., matching 'מקט' even if written as "מק\"ט") and
    detecting a price column containing the substring 'מחיר'. It performs a left
    join from the expenses DataFrame to the price_list on the detected serial
    number column and adds a status indicating price match, mismatch, or missing.

    Parameters
    ----------
    expenses : pd.DataFrame
        Filtered expense report DataFrame (June expenses).
    price_list : pd.DataFrame
        Price list DataFrame.

    Returns
    -------
    pd.DataFrame
        DataFrame containing the comparison results with an added 'סטאטוס' column.
    """
    # Detect serial number column in both DataFrames
    serial_col_exp = _find_column(expenses, ["מקט"])
    serial_col_price = _find_column(price_list, ["מקט"])

    if serial_col_exp is None or serial_col_price is None:
        raise KeyError("לא נמצאה עמודת מק""ט באחד הקבצים. ודאו שקיימת עמודת מק""ט.")

    # Detect price column in price list: try to find 'מחירון' or any column containing 'מחיר'
    price_col = None
    # Prefer exact match to 'מחירון'
    if 'מחירון' in price_list.columns:
        price_col = 'מחירון'
    else:
        price_col = _find_column(price_list, ["מחירון", "מחיר"])

    if price_col is None:
        raise KeyError("לא נמצאה עמודת מחיר במחירון. ודאו שקיימת עמודת מחירון או עמודת מחיר.")

    # Rename columns for consistent handling
    price_list_renamed = price_list.rename(columns={price_col: 'מחיר מחירון', serial_col_price: 'מקט'})
    expenses_renamed = expenses.rename(columns={serial_col_exp: 'מקט'})

    # Perform left merge on 'מקט'
    merged = pd.merge(
        expenses_renamed,
        price_list_renamed[['מקט', 'מחיר מחירון']],
        on='מקט',
        how='left'
    )

    # Determine status based on presence of expected price and price comparison
    def determine_status(row):
        expected = row['מחיר מחירון']
        actual = row.get('מחיר לפני מע"מ')
        # If there is no expected price (NaN), mark as missing from price list
        if pd.isna(expected):
            return '🟡 חסר במחירון'
        # If the actual charged price equals the expected price
        try:
            # Compare numeric values (convert both to float)
            if float(actual) == float(expected):
                return '✅ תואם'
            else:
                return '❌ לא תואם'
        except Exception:
            return '❌ לא תואם'

    merged['סטאטוס'] = merged.apply(determine_status, axis=1)
    return merged


def create_downloadable_excel(df: pd.DataFrame) -> bytes:
    """Create an Excel file in memory containing the comparison results.

    Parameters
    ----------
    df : pd.DataFrame
        DataFrame containing the comparison results.

    Returns
    -------
    bytes
        Bytes representing the Excel file ready for download.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Comparison')
    # Seek to start of the BytesIO buffer
    output.seek(0)
    return output.getvalue()


def main():
    """Main function to run the Streamlit app."""
    st.set_page_config(
        page_title="השוואת מחירון והוצאות",
        page_icon="📊",
        layout="centered"
    )

    # App title and description (in Hebrew)
    st.title("השוואת מחירון מול דוח הוצאות")
    st.markdown(
        """
        ## ברוכים הבאים!
        כאן תוכלו להשוות את דוח ההוצאות שלכם מול מחירון עדכני, ולזהות במהירות התאמות או
        הבדלים במחיר. האפליקציה עובדת באופן הבא:
        1. מעלים קובץ מחירון וקובץ הוצאות.
        2. לוחצים על "השווה עכשיו" לקבלת דוח השוואה.
        3. מורידים את קובץ התוצאה שמופיע לאחר ההשוואה.
        """
    )

    # File uploaders
    price_file = st.file_uploader("העלה קובץ מחירון", type=["xlsx", "xls"], key="price")
    expense_file = st.file_uploader("העלה קובץ הוצאות", type=["xlsx", "xls"], key="expenses")

    # Ensure both files are uploaded before proceeding
    if price_file is not None and expense_file is not None:
        if st.button("השווה עכשיו", key="compare_button"):
            # Load files into DataFrames
            try:
                price_df = load_excel(price_file)
                expenses_df = load_excel(expense_file)
            except Exception as e:
                st.error(f"אירעה שגיאה בקריאת הקבצים: {e}")
                return

            # Filter June expenses
            june_expenses = filter_june_expenses(expenses_df)
            if june_expenses.empty:
                st.warning("לא נמצאו הוצאות לחודש יוני.")
                return

            # Perform comparison
            comparison_df = compare_prices(june_expenses, price_df)

            # Create downloadable Excel file
            excel_bytes = create_downloadable_excel(comparison_df)

            # Show download button with download icon
            st.success("השוואה הושלמה בהצלחה!")
            st.dataframe(comparison_df)  # Display comparison results
            st.download_button(
                label="📥 הורד קובץ השוואה",
                data=excel_bytes,
                file_name="comparison.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        # Inform the user that both files are required
        st.info("אנא העלו גם קובץ מחירון וגם קובץ הוצאות כדי להמשיך.")


if __name__ == "__main__":
    main()