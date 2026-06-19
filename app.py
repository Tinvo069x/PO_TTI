import streamlit as st
import pandas as pd
from datetime import datetime
from collections import defaultdict
import io

# ===================================
# CONFIG
# ===================================

st.set_page_config(
    page_title="PO Tracking Processor",
    layout="wide"
)

st.title("PO Tracking Processor")
st.markdown("Upload Open CD file and generate PO Tracking report")

# ===================================
# SAFE RENAME
# ===================================

def safe_rename_column(df, old_name, new_name):

    if old_name not in df.columns:
        return df

    target = new_name
    i = 1

    while target in df.columns:
        target = f"{new_name}_{i}"
        i += 1

    return df.rename(columns={old_name: target})


# ===================================
# PROCESS FUNCTION
# ===================================

def process_file(df):

    today_col_name = datetime.today().strftime("%d-%b-%Y")

    df = safe_rename_column(
        df,
        "Past due",
        today_col_name
    )

    if "Type" not in df.columns:
        raise ValueError(
            "Input file must contain column 'Type'"
        )

    df = df[
        df["Type"].isin(
            ["Firm", "Forecast"]
        )
    ].copy()

    # ==========================
    # MONTH GROUPING
    # ==========================

    col_month_map = {}

    for col in df.columns:

        parsed = pd.to_datetime(
            str(col),
            errors="coerce",
            dayfirst=True
        )

        if pd.notna(parsed):

            col_month_map[col] = parsed.strftime(
                "%Y-%m"
            )

    month_groups = defaultdict(list)

    for col, month in col_month_map.items():
        month_groups[month].append(col)

    info_cols = [
        c
        for c in df.columns
        if c not in col_month_map
    ]

    df_info = df[info_cols].copy()

    df_months = pd.DataFrame(index=df.index)

    for month, cols in sorted(
        month_groups.items()
    ):

        df_months[month] = (
            df[cols]
            .apply(
                pd.to_numeric,
                errors="coerce"
            )
            .sum(
                axis=1,
                skipna=True
            )
        )

    df_result = pd.concat(
        [
            df_info.reset_index(drop=True),
            df_months.reset_index(drop=True)
        ],
        axis=1
    )

    groupby_cols = [
        c
        for c in info_cols
        if c != "Type"
    ]

    df_result = (
        df_result
        .groupby(
            groupby_cols,
            dropna=False
        )
        .sum(
            numeric_only=True
        )
        .reset_index()
    )

    # ==========================
    # INSERT COLUMNS
    # ==========================

    new_cols = [
        "Status",
        "Run_out_stock_until",
        "RYG_PO Qty Proposal",
        "Qty_Excess",
        "Focus_Vendor",
        "Control WOS(3month)_ASN",
        "Classification",
        "Reserved_2",
        "Reserved_3",
        "Reserved_4"
    ]

    for col in reversed(new_cols):

        df_result.insert(
            0,
            col,
            ""
        )

    return df_result


# ===================================
# FILE UPLOAD
# ===================================

uploaded_file = st.file_uploader(
    "Upload Excel File",
    type=["xlsx", "xls"]
)

if uploaded_file:

    try:

        with st.spinner("Reading file..."):

            df = pd.read_excel(
                uploaded_file,
                keep_default_na=False
            )

        st.success(
            f"Loaded {len(df):,} rows"
        )

        st.dataframe(
            df.head(),
            use_container_width=True
        )

        if st.button(
            "Run Data Processing",
            type="primary"
        ):

            result = process_file(df)

            st.success(
                f"Completed - {len(result):,} rows"
            )

            st.dataframe(
                result.head(),
                use_container_width=True
            )

            output = io.BytesIO()

            with pd.ExcelWriter(
                output,
                engine="openpyxl"
            ) as writer:

                result.to_excel(
                    writer,
                    sheet_name="Result",
                    index=False
                )

            today_str = datetime.today().strftime(
                "%Y%m%d"
            )

            st.download_button(
                label="Download Result",
                data=output.getvalue(),
                file_name=f"PO_Tracking_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:

        st.error(str(e))