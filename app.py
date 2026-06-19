
import streamlit as st
import pandas as pd
from datetime import datetime
from collections import defaultdict
import io

# =====================================================
# CONFIG
# =====================================================

st.set_page_config(
    page_title="PO Tracking Processor",
    layout="wide"
)

st.title("PO Tracking Processor")
st.markdown("Upload Open CD file and generate PO Tracking report")

# =====================================================
# HELPER FUNCTIONS
# =====================================================

def safe_rename_column(df, old_name, new_name):

    if old_name not in df.columns:
        return df

    target = new_name
    i = 1

    while target in df.columns:
        target = f"{new_name}_{i}"
        i += 1

    return df.rename(columns={old_name: target})


def detect_month_columns(df):

    month_cols = []

    for col in df.columns:

        try:

            dt = pd.to_datetime(
                str(col),
                errors="raise"
            )

            if dt.year >= 2024:
                month_cols.append(col)

        except:
            pass

    return month_cols


def calculate_runout(row, month_cols):

    supply = (
        pd.to_numeric(
            row.get("Store_Qty", 0),
            errors="coerce"
        )
        +
        pd.to_numeric(
            row.get("IQC_QTY", 0),
            errors="coerce"
        )
        +
        pd.to_numeric(
            row.get("PO_QTY", 0),
            errors="coerce"
        )
    )

    supply = 0 if pd.isna(supply) else supply

    balance = supply

    for col in month_cols:

        demand = pd.to_numeric(
            row[col],
            errors="coerce"
        )

        demand = 0 if pd.isna(demand) else demand

        balance -= demand

        if balance < 0:
            return str(col)

    return "999"


def calculate_excess(row, month_cols):

    supply = (
        pd.to_numeric(
            row.get("Store_Qty", 0),
            errors="coerce"
        )
        +
        pd.to_numeric(
            row.get("IQC_QTY", 0),
            errors="coerce"
        )
        +
        pd.to_numeric(
            row.get("PO_QTY", 0),
            errors="coerce"
        )
    )

    supply = 0 if pd.isna(supply) else supply

    demand = (
        pd.to_numeric(
            row[month_cols],
            errors="coerce"
        )
        .fillna(0)
        .sum()
    )

    return max(
        0,
        supply - demand
    )


def calculate_classification(row, month_cols):

    active_months = (
        pd.to_numeric(
            row[month_cols],
            errors="coerce"
        )
        .fillna(0)
        .gt(0)
        .sum()
    )

    ratio = active_months / len(month_cols)

    if ratio >= 0.8:
        return "Fast Moving"

    return "Slow Moving"


def calculate_status(row):

    runout = row["Run_out_stock_until"]

    if runout == "999":
        return "Excess PO"

    return "Follow"


# =====================================================
# MAIN PROCESS
# =====================================================

def process_file(df):

    today_col_name = datetime.today().strftime(
        "%d-%b-%Y"
    )

    df = safe_rename_column(
        df,
        "Past due",
        today_col_name
    )

    if "Type" not in df.columns:
        raise ValueError(
            "Input file must contain column Type"
        )

    df = df[
        df["Type"].isin(
            ["Firm", "Forecast"]
        )
    ].copy()

    # ==========================================
    # GROUP MONTH
    # ==========================================

    col_month_map = {}

    for col in df.columns:

        try:

            dt = pd.to_datetime(
                str(col),
                errors="raise"
            )

            col_month_map[col] = (
                dt.strftime("%Y-%m")
            )

        except:
            pass

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

    # ==========================================
    # GROUP KEY
    # ==========================================

    groupby_cols = []

    if "Vendor_Code" in df_result.columns:
        groupby_cols.append("Vendor_Code")

    if "Site" in df_result.columns:
        groupby_cols.append("Site")

    if "Part_No" in df_result.columns:
        groupby_cols.append("Part_No")

    if len(groupby_cols) == 0:
        raise ValueError(
            "Cannot find Vendor_Code / Site / Part_No"
        )

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

    # ==========================================
    # CALCULATE LOGIC
    # ==========================================

    month_cols = detect_month_columns(
        df_result
    )

    df_result["Run_out_stock_until"] = (
        df_result.apply(
            lambda r:
            calculate_runout(
                r,
                month_cols
            ),
            axis=1
        )
    )

    df_result["Qty_Excess"] = (
        df_result.apply(
            lambda r:
            calculate_excess(
                r,
                month_cols
            ),
            axis=1
        )
    )

    df_result["Classification"] = (
        df_result.apply(
            lambda r:
            calculate_classification(
                r,
                month_cols
            ),
            axis=1
        )
    )

    df_result["Status"] = (
        df_result.apply(
            calculate_status,
            axis=1
        )
    )

    return df_result


# =====================================================
# UPLOAD
# =====================================================

uploaded_file = st.file_uploader(
    "Upload Excel File",
    type=["xlsx", "xls", "xlsb"]
)

if uploaded_file:

    try:

        with st.spinner(
            "Reading file..."
        ):

            if uploaded_file.name.endswith(".xlsb"):

                df = pd.read_excel(
                    uploaded_file,
                    engine="pyxlsb"
                )

            else:

                df = pd.read_excel(
                    uploaded_file
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
                result.head(20),
                use_container_width=True
            )

            output = io.BytesIO()

            with pd.ExcelWriter(
                output,
                engine="openpyxl"
            ) as writer:

                result.to_excel(
                    writer,
                    index=False,
                    sheet_name="Result"
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

