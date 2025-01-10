import pandas as pd
import streamlit as st
from io import BytesIO
import re
from openpyxl.styles import PatternFill
import os

# Display the GIF centered above the title
st.markdown(
    """
    <div style="text-align: center;">
        <img src="project-folder/assets/chillguy.gif" alt="Loading Animation" style="width:300px; height:auto;">
    </div>
    """,
    unsafe_allow_html=True
)

# Streamlit app title
st.title("Fuckin' Awesome File Convertor")

# Step 1: File uploader
st.header("Upload Excel Files")
uploaded_files = st.file_uploader("Upload one or more Excel files", type=["xlsx"], accept_multiple_files=True)

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
        # Step 3: Combine all sheets into one DataFrame
        combined_df = pd.concat(all_data, ignore_index=True)
        st.write("### Combined Data Preview (Before Renaming)")
        st.dataframe(combined_df)

        # Step 3.1: Add HANDLING COST column
        st.write("### Adding HANDLING COST Column")
        combined_df["HANDLING COST"] = 0.75
        st.success("HANDLING COST column added with default value 0.75.")

        # Step 4: Standardize and Rename Columns
        st.write("### Renaming Columns")
        combined_df.columns = combined_df.columns.str.strip()  # Strip column headers of extra spaces
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

        # Step 5: Clean TITLE column
        if "TITLE" in combined_df.columns:
            combined_df["TITLE"] = (
                combined_df["TITLE"]
                .str.replace(r"\(W\+\)", "", regex=True)
                .str.replace(r"\(SP\)", "", regex=True)
                .str.replace(r"\(P\)", "", regex=True)
                .str.strip()
            )
            st.success("Unwanted patterns removed from TITLE column.")

        # Step 6: Format UPC/ISBN column
        if "UPC/ISBN" in combined_df.columns:
            combined_df["UPC/ISBN"] = combined_df["UPC/ISBN"].astype(str).str.zfill(12)
            st.success("UPC/ISBN column formatted to have a minimum of 12 digits.")

        # Step 7: Format COST_PRICE column
        st.write("### Formatting COST_PRICE Column")
        if "COST_PRICE" in combined_df.columns:
            combined_df["COST_PRICE"] = (
                combined_df["COST_PRICE"]
                .astype(str)
                .str.replace(r"[$,]", "", regex=True)  # Remove currency symbols and commas
                .astype(float)
                .round(2)
            )
            st.success("COST_PRICE column formatted to numeric with two decimal places.")

        # Step 8: Add QUANTITY and ITEM LOCATION columns
        combined_df["QUANTITY"] = 1
        combined_df["ITEM LOCATION"] = "WALMART"

        # Step 9: Add ITEM WEIGHT (pounds) column
        st.write("### Adding ITEM WEIGHT (pounds) Column")

        def extract_weight(title):
            match = re.search(r"(\d+(\.\d+)?)\s*(?:oz|ounces|fl oz|fluid ounces)", title, re.IGNORECASE)
            if match:
                weight = float(match.group(1))
                if re.search(r"fl oz|fluid ounces", title, re.IGNORECASE):
                    weight += 10
                else:
                    weight += 6
                return weight
            return None

        combined_df["ITEM WEIGHT (pounds)"] = combined_df["TITLE"].apply(lambda x: extract_weight(x) if isinstance(x, str) else None)
        combined_df["ITEM WEIGHT (pounds)"] = combined_df["ITEM WEIGHT (pounds)"].apply(
            lambda x: round(x / 16, 2) if x is not None else None
        )

        # Step 9.1: Highlight rows with missing weights
        st.write("### Highlighting Rows with Missing ITEM WEIGHT (pounds)")

        # Step 10: Add SHIPPING COST column
        st.write("### Adding SHIPPING COST Column")

        shipping_legend_path = "project-folder/data/default_shipping_legend.xlsx"  # Relative path in the repository

        if not os.path.exists(shipping_legend_path):
            st.error(f"The shipping legend file does not exist at the specified path: {shipping_legend_path}")
        else:
            try:
                shipping_legend = pd.read_excel(shipping_legend_path, engine="openpyxl")
                st.success("Shipping legend file loaded successfully.")
            except Exception as e:
                st.error(f"Error reading shipping legend file: {e}")
                shipping_legend = None

        if shipping_legend is not None and {"Weight Range Min (lb)", "Weight Range Max (lb)", "SHIPPING COST"}.issubset(shipping_legend.columns):
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
        else:
            st.error("Shipping legend file is missing required columns: 'Weight Range Min (lb)', 'Weight Range Max (lb)', 'SHIPPING COST'.")

        # Step 10.1: Add RETAIL PRICE column
        if all(col in combined_df.columns for col in ["COST_PRICE", "SHIPPING COST", "HANDLING COST"]):
            combined_df["RETAIL PRICE"] = combined_df.apply(
                lambda row: round(
                    (row["COST_PRICE"] + row["SHIPPING COST"] + row["HANDLING COST"]) * 1.35, 2
                ) if not (pd.isnull(row["COST_PRICE"]) or pd.isnull(row["SHIPPING COST"]) or pd.isnull(row["HANDLING COST"])) else None,
                axis=1
            )

        # Step 10.2: Add MIN PRICE and MAX PRICE columns
        if all(col in combined_df.columns for col in ["SHIPPING COST", "ITEM WEIGHT (pounds)", "RETAIL PRICE"]):
            combined_df["MIN PRICE"] = combined_df["RETAIL PRICE"]
            combined_df["MAX PRICE"] = combined_df["RETAIL PRICE"].apply(lambda x: round(x * 1.35, 2) if x is not None else None)

        # Step 10.3: Remove rows with "Great Value"
        combined_df = combined_df[~combined_df["TITLE"].str.contains("Great Value", case=False, na=False)]

        # Step 11: Remove duplicate rows
        combined_df.drop_duplicates(inplace=True)

        # Step 12: Final Output and Download
        st.write("### Final Data Preview")
        st.dataframe(combined_df)

        if st.button("Export to Excel"):
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                combined_df.to_excel(writer, index=False, sheet_name="Consolidated Data")
                worksheet = writer.sheets["Consolidated Data"]

                red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
                for row_index, weight in enumerate(combined_df["ITEM WEIGHT (pounds)"], start=2):
                    if pd.isnull(weight):
                        for col_index in range(1, len(combined_df.columns) + 1):
                            worksheet.cell(row=row_index, column=col_index).fill = red_fill

            buffer.seek(0)
            st.download_button(
                label="Download Excel File",
                data=buffer.getvalue(),
                file_name="Consolidated_Data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
else:
    st.info("Upload one or more Excel files to get started.")
