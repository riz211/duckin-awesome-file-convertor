import pandas as pd
import streamlit as st
from io import BytesIO
import re
from openpyxl.styles import PatternFill
import os

# App title
st.title("Fuckin' Awesome File Convertor")

# Step 1: File uploader
st.header("Upload Excel Files")
uploaded_files = st.file_uploader("Upload one or more Excel files", type=["xlsx"], accept_multiple_files=True)

# Initialize an empty DataFrame to avoid undefined variable errors
combined_df = pd.DataFrame()

if uploaded_files:
    all_data = []

    # Step 2: Process each uploaded file
    for uploaded_file in uploaded_files:
        try:
            excel_file = pd.ExcelFile(uploaded_file)
            for sheet_name in excel_file.sheet_names:
                sheet_data = pd.read_excel(uploaded_file, sheet_name=sheet_name, usecols="B,E,G,H,I")
                all_data.append(sheet_data)
        except Exception as e:
            st.error(f"Error reading file {uploaded_file.name}: {e}")

    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        st.write("### Combined Data Preview (Before Renaming)")
        st.dataframe(combined_df)

        # Step 3.1: Add HANDLING COST column
        combined_df["HANDLING COST"] = 0.75
        st.success("HANDLING COST column added with default value 0.75.")

        # Step 4: Standardize and Rename Columns
        combined_df.columns = combined_df.columns.str.strip()
        column_mapping = {
            "Product Details": "TITLE",
            "Brand": "BRAND",
            "Product ID": "SKU",
            "UPC Code": "UPC/ISBN",
            "Price": "COST_PRICE"
        }
        combined_df.rename(columns=column_mapping, inplace=True)

        if "COST_PRICE" not in combined_df.columns:
            st.error("COST_PRICE column is missing. Ensure the input file has a 'Price' column.")
        else:
            st.success("Columns renamed successfully.")

        # Format SKU and UPC/ISBN columns
        if "SKU" in combined_df.columns:
            combined_df["SKU"] = combined_df["SKU"].astype(str).str.replace(",", "").str.strip()

        if "UPC/ISBN" in combined_df.columns:
            combined_df["UPC/ISBN"] = (
                combined_df["UPC/ISBN"]
                .apply(lambda x: str(int(float(x))) if pd.notnull(x) else "")
                .str.zfill(12)
            )

        # Format COST_PRICE column
        if "COST_PRICE" in combined_df.columns:
            combined_df["COST_PRICE"] = (
                combined_df["COST_PRICE"]
                .astype(str)
                .str.replace(r"[$,]", "", regex=True)
                .astype(float)
                .round(2)
            )

        # Add QUANTITY and ITEM LOCATION columns
        combined_df["QUANTITY"] = 1
        combined_df["ITEM LOCATION"] = "WALMART"

        # ITEM WEIGHT calculation
        def extract_weight_with_packs(title):
            try:
                match_weight = re.search(r"(\d+(\.\d+)?)\s*(?:oz|ounces|fl oz|fluid ounces)", title, re.IGNORECASE)
                single_unit_weight = float(match_weight.group(1)) if match_weight else None

                match_pack = re.search(r"(?:\b(\d+)\s*pack\b|\bpack of\s*(\d+))", title, re.IGNORECASE)
                pack_size = int(match_pack.group(1) or match_pack.group(2)) if match_pack else 1

                if single_unit_weight is not None:
                    single_unit_weight += 10 if "fl oz" in match_weight.group(0).lower() else 6
                    return round((single_unit_weight * pack_size) / 16, 2)
            except Exception as e:
                st.error(f"Error processing title '{title}': {e}")
            return None

        if "TITLE" in combined_df.columns:
            combined_df["ITEM WEIGHT (pounds)"] = combined_df["TITLE"].apply(
                lambda x: extract_weight_with_packs(x) if isinstance(x, str) else None
            )

        # SHIPPING COST calculation
        shipping_legend_path = "project-folder/data/default_shipping_legend.xlsx"
        if os.path.exists(shipping_legend_path):
            shipping_legend = pd.read_excel(shipping_legend_path, engine="openpyxl")
            if {"Weight Range Min (lb)", "Weight Range Max (lb)", "SHIPPING COST"}.issubset(shipping_legend.columns):
                def calculate_shipping_cost(weight, legend):
                    if pd.isnull(weight):
                        return None
                    for _, row in legend.iterrows():
                        if row["Weight Range Min (lb)"] <= weight <= row["Weight Range Max (lb)"]:
                            return row["SHIPPING COST"]
                    return None

                combined_df["SHIPPING COST"] = combined_df["ITEM WEIGHT (pounds)"].apply(
                    lambda w: calculate_shipping_cost(w, shipping_legend)
                )

        # RETAIL PRICE and other calculations
        if all(col in combined_df.columns for col in ["COST_PRICE", "SHIPPING COST", "HANDLING COST"]):
            combined_df["RETAIL PRICE"] = combined_df.apply(
                lambda row: round((row["COST_PRICE"] + row["SHIPPING COST"] + row["HANDLING COST"]) * 1.35, 2)
                if not any(pd.isnull(row[col]) for col in ["COST_PRICE", "SHIPPING COST", "HANDLING COST"])
                else None,
                axis=1
            )
            combined_df["MIN PRICE"] = combined_df["RETAIL PRICE"]
            combined_df["MAX PRICE"] = combined_df["RETAIL PRICE"].apply(lambda x: round(x * 1.35, 2) if x else None)

        combined_df.drop_duplicates(inplace=True)

        # Metrics Summary
        total_input = len(all_data)
        total_output = len(combined_df)
        missing_weights = combined_df["ITEM WEIGHT (pounds)"].isnull().sum()

        st.markdown(f"""
        - **Total Listings in Input Files:** {total_input}
        - **Total Listings in Output File:** {total_output}
        - **Missing Weights (Highlighted Rows):** {missing_weights}
        """)

        # Export Logic
        if st.button("Export to Excel"):
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                combined_df.to_excel(writer, index=False, sheet_name="Consolidated Data")
                worksheet = writer.sheets["Consolidated Data"]

                # Highlight missing weights
                red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
                for row, weight in enumerate(combined_df["ITEM WEIGHT (pounds)"], start=2):
                    if pd.isnull(weight):
                        for col in range(1, len(combined_df.columns) + 1):
                            worksheet.cell(row=row, column=col).fill = red_fill

            st.download_button(
                "Download Consolidated File",
                data=buffer.getvalue(),
                file_name="Consolidated_File.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
else:
    st.info("Upload files to start processing.")
