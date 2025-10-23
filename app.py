#make venv
#py -3.13 -m venv .venv  

#bypass for windows
#Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

#activate venv
#.\.venv\Scripts\Activate.ps1

#download libraries

#pip install yfinance pandas streamlit openpyxl xlsxwriter

#verifying installation
#python -c "import yfinance, pandas, streamlit; print('All good')"

#this is how you run thge app through streamlit
#streamlit run app.py
#OR python -m streamlit run app.py (if referencing wrong venv)

#when done type
#crtl + c 

import streamlit as st
import yfinance as yf
import pandas as pd
from io import BytesIO
from datetime import datetime

# configuring page. tab identity (streamlit)
st.set_page_config(page_title="Income Statement Viewer", page_icon="💰", layout="wide")

# setting title (streamlit)
st.title("Income Statement Viewer")

# user ticket impuit (streamlit)
ticker_symbol = st.text_input( "Enter a Stock Ticker:", "" )

# annual or quarterly selector
period_type = st.radio("Period:", ["Annual", "Quarterly"])

#formatting the excel sheet to look pretty (income statement)
def build_formatted_excel(df: pd.DataFrame, ticker: str, period_label: str) -> bytes:

    # make copy of dataframe, names dataframe indexes (columns), moveing index to column
    export_df = df.copy()
    export_df.index.name = "Line Item"
    export_df.reset_index(inplace=True)

    #in-memory binary stream to house the excel file
    buffer = BytesIO()

    #opening excelwriter engine, setting tab name, writeing dataframe onto the sheet
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        sheet_name = "Income Statement"
        export_df.to_excel(writer, index=False, sheet_name=sheet_name)

        #wb references entire workbook for edits, ws references worksheet and sheet name
        wb = writer.book
        ws = writer.sheets[sheet_name]

        # formatting
        header_fmt = wb.add_format({
            "bold": True, "text_wrap": True, "align": "center", "valign": "vcenter"
        })
        name_fmt = wb.add_format({
            "font_color": "#1f4e79", "valign": "vcenter"
        })
        millions_fmt = wb.add_format({
            "num_format": '#,##0.00,,"M"', "valign": "vcenter"
        })
        title_fmt = wb.add_format({
            "bold": True, "font_size": 14, "valign": "vcenter"
        })

        # title row
        title = f"{ticker.upper()} {period_label} Income Statement – exported {datetime.now():%Y-%m-%d}"
        ws.write(0, 0, title, title_fmt)

        # header formatting
        for col_idx, col_name in enumerate(export_df.columns):
            ws.write(0, col_idx, col_name, header_fmt)
        ws.set_row(0, 24)  # taller header

        # column formatting
        max_name_len = max(len(str(x)) for x in export_df["Line Item"].astype(str))
        ws.set_column(0, 0, max(14, min(60, int(max_name_len * 0.95))), name_fmt)
        ws.set_column(1, export_df.shape[1] - 1, 14, millions_fmt)

        # freezes first row and first column, gives user dropdowns on header row
        ws.freeze_panes(1, 1)
        ws.autofilter(0, 0, 0, export_df.shape[1] - 1)

        # highlighting key metrics
        bold_fmt = wb.add_format({"bold": True, "font_color": "#1f4e79"})
        highlight_items = {"Total Revenue", "Operating Income", "Net Income", "Gross Profit"}
        for r in range(1, export_df.shape[0] + 1):
            if str(export_df.iloc[r-1, 0]) in highlight_items:
                ws.write(r, 0, export_df.iloc[r-1, 0], bold_fmt)

    return buffer.getvalue()

# data after button is pressed (income statement)
if st.button("Fetch Data"):
    if ticker_symbol.strip() == "":
        st.warning("Please enter a valid ticker symbol.")
    else:
        try:
           #this part fetches the income statement based on whats answeres on the toggle
            ticker = yf.Ticker(ticker_symbol)
            income_stmt = (
                ticker.quarterly_income_stmt
                if period_type == "Quarterly"
                else ticker.income_stmt
            )

            if income_stmt is None or income_stmt.empty:
                st.error(f"No income statement data found for {ticker_symbol}.")
            else:
                st.success(f"Successfully retrieved data for {ticker_symbol}.")

                # display dataframe
                st.dataframe(income_stmt)

                # show basic info
                st.write(f"**Number of line items:** {income_stmt.shape[0]}")
                st.write(f"**Number of periods:** {income_stmt.shape[1]}")

                excel_bytes = build_formatted_excel(income_stmt, ticker_symbol, period_type)

                # Create the download button for user
                filename = f"{ticker_symbol.upper()}_{period_type.lower()}_Income_Statement.xlsx"
                st.download_button(
                 label=" Download Formatted Excel File",
                  data=excel_bytes,
                 file_name=filename,
                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                # -------------------------------------------

        except Exception as e:
            st.error(f"Error fetching data: {e}")