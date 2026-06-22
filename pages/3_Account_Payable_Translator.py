import io
import pandas as pd
import streamlit as st

REQUIRED_COLUMNS = ['code', 'invoice#', 'invoice date', 'due date','retainage','sub-total','total due']

def normalize(s):
    try:
        return str(s).strip().lower()
    except:
        return ''

def find_header_positions(df, required_columns, max_rows=50):
    """
    Find column positions for each required column (can be on different rows)
    """
    mapping = {}
    for i in range(min(max_rows, len(df))):
        row = df.iloc[i]
        for j, cell in enumerate(row):
            val = normalize(cell)
            if val in required_columns and val not in mapping:
                mapping[val] = j
        if len(mapping) == len(required_columns):
            break

    missing = set(required_columns) - set(mapping.keys())
    if missing:
        raise ValueError(f"Maestro file is missing required columns: {missing}")
    return mapping, i  # i = last header row

def read_maestro_file(uploaded_file):
    """
    Reads Maestro file, detects scattered headers, keeps all columns.
    """
    uploaded_file.seek(0)
    raw_df = pd.read_excel(uploaded_file, header=None, engine='openpyxl')

    mapping, last_header_row = find_header_positions(raw_df, REQUIRED_COLUMNS)

    # Read full file again, no header
    uploaded_file.seek(0)
    full_df = pd.read_excel(uploaded_file, header=None, engine='openpyxl')

    # Set proper column names for required columns
    for col_name, col_idx in mapping.items():
        full_df.rename(columns={col_idx: col_name.title()}, inplace=True)

    # Drop rows above last header row
    df = full_df.iloc[last_header_row+1:].reset_index(drop=True)

    # Drop rows where all required columns are blank
    df.dropna(subset=[col.title() for col in REQUIRED_COLUMNS], how='all', inplace=True)


    return df

def merge_vendor_ids(df_maestro, df_mapping):
    """
    Merge Maestro DataFrame with mapping file to get Vendor_ID without dropping other data.
    """
    df_mapping['Code'] = df_mapping['Code'].astype(str).str.strip()
    df_maestro['Code'] = df_maestro['Code'].astype(str).str.strip()

    df_merged = pd.merge(df_maestro, df_mapping[['Code', 'Intact_Vendor_ID']],
                         on='Code', how='left')

    # Reorder columns: VENDOR_ID + required columns
    first_cols = ['Intact_Vendor_ID'] + [col.title() for col in REQUIRED_COLUMNS]
    other_cols = [c for c in df_merged.columns if c not in first_cols]
    df_merged = df_merged[first_cols + other_cols]

    return df_merged

# ----------------- Streamlit UI -----------------
st.title("AP Data Translator")

maestro_file = st.file_uploader("Upload Maestro Excel file", type=["xlsx", "xls"])
mapping_file = st.file_uploader("Upload Vendor ID Mapping Excel file", type=["xlsx", "xls"])

if maestro_file and mapping_file:
    try:
        df_maestro = read_maestro_file(maestro_file)
        df_mapping = pd.read_excel(mapping_file, engine='openpyxl')

        # Merge to get VENDOR_ID
        df_translated = merge_vendor_ids(df_maestro, df_mapping)
        
        #Rename columns for output only
        df_translated.rename(columns={
            'Code': 'Maestro_ID',  # rename Code to Maestro_ID
        }, inplace=True)

        # Define final columns for output
        final_cols = ['Intact_Vendor_ID', 'Maestro_ID'] + [col.title() for col in REQUIRED_COLUMNS if col.lower() != 'code']
        # Keep only the desired columns
        #final_cols = ['VENDOR_ID'] + [col.title() for col in REQUIRED_COLUMNS]
        df_translated = df_translated[final_cols]

        st.success("Files processed successfully!")
        st.dataframe(df_translated)

        # Provide download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_translated.to_excel(writer, index=False)

        output.seek(0)

        st.download_button(
            label="Download Translated Excel",
            data=output,
            file_name="translated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error processing files: {e}")
