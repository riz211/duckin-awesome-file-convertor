import pandas as pd
import streamlit as st
from io import BytesIO
import re
from openpyxl.styles import PatternFill
import os

combined_df = pd.DataFrame()  # Initialize an empty DataFrame

# App title
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
        if not combined_df.empty:  # Check if combined_df exists and is not empty
            if "ITEM WEIGHT (pounds)" in combined_df.columns:
                combined_df["Missing Weight"] = combined_df["ITEM WEIGHT (pounds)"].isnull()
                st.write("Missing weights flagged successfully.")
            else:
                st.error("ITEM WEIGHT (pounds) column is missing.")
    else:
        st.error("The DataFrame is not defined or is empty. Please upload files to process.")
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

    # Step 4.1: Format SKU column to remove commas and ensure it is displayed as a string
    if "SKU" in combined_df.columns:
        combined_df["SKU"] = combined_df["SKU"].astype(str).str.replace(",", "").str.strip()
        st.success("SKU column formatted to remove commas.")

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
    st.write("### Formatting UPC/ISBN Column")
    if "UPC/ISBN" in combined_df.columns:
        combined_df["UPC/ISBN"] = (
            combined_df["UPC/ISBN"]
            .apply(lambda x: str(int(float(x))) if pd.notnull(x) else "")  # Convert to integer, then string
            .str.zfill(12)  # Add leading zeros to ensure 12 digits
        )
        st.success("UPC/ISBN column formatted to have a minimum of 12 digits as a string.")

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

    # Function to extract weight and handle pack sizes
    def extract_weight_with_packs(title):
        """
        Extract the weight and account for pack size in the TITLE.
        """
        try:
            match_weight = re.search(r"(\d+(\.\d+)?)\s*(?:oz|ounces|ounce|fl. oz.|fluid ounce|fl oz|fluid ounces)", title, re.IGNORECASE)
            single_unit_weight = float(match_weight.group(1)) if match_weight else None
            match_pack = re.search(r"(?:\b(\d+)\s*pack\b|\bpack of\s*(\d+))", title, re.IGNORECASE)
            pack_size = int(match_pack.group(1) or match_pack.group(2)) if match_pack else 1
            if single_unit_weight is not None:
                if match_weight and "fl oz" in match_weight.group(0).lower():
                    single_unit_weight += 10
                else:
                    single_unit_weight += 6
                total_weight = (single_unit_weight * pack_size) / 16
                return round(total_weight, 2)
        except Exception as e:
            st.error(f"Error processing title '{title}': {e}")
        return None

    if "TITLE" in combined_df.columns:
        combined_df["ITEM WEIGHT (pounds)"] = combined_df["TITLE"].apply(
            lambda x: extract_weight_with_packs(x) if isinstance(x, str) else None
        )
        st.success("ITEM WEIGHT (pounds) column updated to account for pack sizes.")

    st.write("### Highlighting Rows with Missing ITEM WEIGHT (pounds)")

    if "ITEM WEIGHT (pounds)" in combined_df.columns:
        combined_df["Missing Weight"] = combined_df["ITEM WEIGHT (pounds)"].isnull()
        st.write("Missing weights have been flagged successfully.")
    else:
        st.error("ITEM WEIGHT (pounds) column is missing.")

    # Additional logic for shipping costs and other columns (same as before)

else:
    st.info("Please upload one or more Excel files to start processing.")
