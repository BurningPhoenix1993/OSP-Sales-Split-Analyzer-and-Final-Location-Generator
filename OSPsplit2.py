import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="OSP Final Split Calculator", layout="wide")
st.title("📊 OSP Final Location + Sales Split Analysis")

# Upload files
st.header("Upload Files")
sales_file = st.file_uploader("Upload Sales History Excel", type=["xlsx"])
osp_file = st.file_uploader("Upload OSP Final Location Excel", type=["xlsx"])

if sales_file and osp_file:
    try:
        sales_df = pd.read_excel(sales_file, dtype=str).fillna("")
        osp_df = pd.read_excel(osp_file, dtype=str).fillna("")

        week_cols = [col for col in sales_df.columns if col.startswith("WK")]
        base_cols = ["SKU", "Ship to", "IPAG", "SCH", "Location"]
        sales_df[week_cols] = sales_df[week_cols].apply(pd.to_numeric, errors='coerce')

        osp_cols = [col for col in osp_df.columns if col.startswith("OSP WK")]
        merge_cols = ["SKU", "Ship to"]

        df = pd.merge(sales_df, osp_df[merge_cols + osp_cols], on=merge_cols, how="left")

        df["Total sum"] = df[week_cols].sum(axis=1)
        df["SUMIFS SKU-IPAG"] = df.groupby(["SKU", "IPAG"])["Total sum"].transform("sum")
        df["SUMIFS SKU-SCH"] = df.groupby(["SKU", "SCH"])["Total sum"].transform("sum")

        for col in osp_cols:
            wk = col.replace("OSP ", "").replace(" Final Location", "")
            ipag_col = f"SUMIFS SKU-IPAG-LOC-{wk}"
            sch_col = f"SUMIFS SKU-SCH-LOC-{wk}"
            split_ipag = f"CSF Split% SKU-IPAG {wk}"
            split_sch = f"CSF Split% SKU-SCH {wk}"

            df[ipag_col] = df.groupby(["SKU", "IPAG", col])["Total sum"].transform("sum")
            df[sch_col] = df.groupby(["SKU", "SCH", col])["Total sum"].transform("sum")

            df[split_ipag] = (df[ipag_col] / df["SUMIFS SKU-IPAG"]).fillna(0)
            df[split_sch] = (df[sch_col] / df["SUMIFS SKU-SCH"]).fillna(0)

        # Create tables for export
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wk_nums = [col.split()[1] for col in osp_cols]  # e.g., 'WK1', 'WK2', ...
            ipag_tables = []
            sch_tables = []

            for wk in wk_nums:
                ipag_cols = ["SKU", "IPAG", f"OSP {wk} Final Location", f"CSF Split% SKU-IPAG {wk}"]
                sch_cols = ["SKU", "SCH", f"OSP {wk} Final Location", f"CSF Split% SKU-SCH {wk}"]

                ipag_table = df[ipag_cols].drop_duplicates().reset_index(drop=True)
                sch_table = df[sch_cols].drop_duplicates().reset_index(drop=True)

                ipag_tables.append(ipag_table)
                sch_tables.append(sch_table)

            # Write SKU-IPAG tables horizontally with gaps
            sheet = writer.book.add_worksheet("SKU-IPAG Tables")
            start_col = 0
            for table in ipag_tables:
                table.to_excel(writer, sheet_name="SKU-IPAG Tables", index=False, startcol=start_col, startrow=0, header=True)
                start_col += len(table.columns) + 1  # Add a 1-column gap

            # Write SKU-SCH tables horizontally with gaps
            sheet = writer.book.add_worksheet("SKU-SCH Tables")
            start_col = 0
            for table in sch_tables:
                table.to_excel(writer, sheet_name="SKU-SCH Tables", index=False, startcol=start_col, startrow=0, header=True)
                start_col += len(table.columns) + 1

        output.seek(0)

        # Display and download
        st.success("✅ Excel with structured tables created!")
        st.download_button(
            label="📥 Download Structured Excel File",
            data=output,
            file_name="OSP_Split_Tables.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
else:
    st.info("👆 Please upload both required Excel files to continue.")
