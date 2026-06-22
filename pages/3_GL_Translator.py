import streamlit as st
import pandas as pd
import io

st.title("GL Account Translator")

# ----------------- Config -----------------
REQUIRED_GL_COLUMNS = ['account', 'description', 'current balance']
REQUIRED_MAP_COLUMNS = ['maestro_account_no', '#old_acct_name', 'sage_account_no', 'acct_name']
MAX_HEADER_ROWS = 50

# ----------------- Helpers -----------------
def normalize(s):
    try:
        return str(s).strip().lower()
    except:
        return ''

def find_header_positions(df, required_columns, max_rows=50):
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
        raise ValueError(f"File is missing required columns: {missing}")
    return mapping, i

def read_messy_file(uploaded_file, required_columns):
    uploaded_file.seek(0)
    raw_df = pd.read_excel(uploaded_file, header=None, engine='openpyxl')
    mapping, last_row = find_header_positions(raw_df, required_columns, MAX_HEADER_ROWS)
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, header=None, engine='openpyxl')
    for col_name, col_idx in mapping.items():
        df.rename(columns={col_idx: col_name}, inplace=True)
    df = df.iloc[last_row+1:].reset_index(drop=True)
    df.columns = [normalize(c) for c in df.columns]
    df.dropna(subset=required_columns, how='all', inplace=True)
    return df

def merge_gl(df_gl, df_map):
    df_gl['maestro_account'] = df_gl['maestro_account'].astype(str).str.strip()
    df_map['maestro_account'] = df_map['maestro_account'].astype(str).str.strip()

    one_to_many = df_map.groupby('maestro_account').filter(lambda x: len(x) > 1)
    split_mapping = {}

    if not one_to_many.empty:
        st.info("One-to-many mapping detected. Enter split percentages:")
        for acct in one_to_many['maestro_account'].unique():
            rows = one_to_many[one_to_many['maestro_account']==acct]
            splits = []
            st.write(f"Maestro Account: {acct}")
            for i, (_, row) in enumerate(rows.iterrows()):
                percent = st.number_input(
                    f"Split % for Sage {row['sage_account']} ({row['sage_account_name']})",
                    min_value=0.0, max_value=100.0, value=100.0/len(rows),
                    key=f"split_{acct}_{i}"
                )
                splits.append(percent)
            if sum(splits)!=100:
                st.warning(f"Splits for {acct} do not sum to 100%. They will be normalized automatically.")
            normalized = [p*100/sum(splits) for p in splits]
            split_mapping[acct] = normalized

    # Merge
    df_merged = df_gl.merge(
        df_map[['maestro_account','sage_account','sage_account_name']],
        on='maestro_account', how='left'
    )

    # Apply splits
    if split_mapping:
        final_rows = []
        for _, row in df_merged.iterrows():
            acct = row['maestro_account']
            multiples = df_map[df_map['maestro_account']==acct]
            if acct in split_mapping:
                for i, (_, map_row) in enumerate(multiples.iterrows()):
                    new_row = row.copy()
                    new_row['sage_account'] = map_row['sage_account']
                    new_row['sage_account_name'] = map_row['sage_account_name']
                    new_row['balance'] = row['balance'] * split_mapping[acct][i]/100
                    final_rows.append(new_row)
            else:
                final_rows.append(row)
        df_merged = pd.DataFrame(final_rows)

    df_merged = df_merged[['maestro_account','sage_account','sage_account_name','description','balance']]
    return df_merged

# ----------------- Streamlit UI -----------------
mapping_file = st.file_uploader("Upload Mapping File (Maestro → Sage)", type=['xlsx'])
gl_file = st.file_uploader("Upload Maestro GL File", type=['xlsx'])

if mapping_file and gl_file:
    try:
        # Read files
        df_map = read_messy_file(mapping_file, REQUIRED_MAP_COLUMNS)
        df_map.rename(columns={
            'maestro_account_no':'maestro_account',
            '#old_acct_name':'old_name',
            'sage_account_no':'sage_account',
            'acct_name':'sage_account_name'
        }, inplace=True)

        df_gl = read_messy_file(gl_file, REQUIRED_GL_COLUMNS)
        df_gl.rename(columns={
            'account':'maestro_account',
            'description':'description',
            'current balance':'balance'
        }, inplace=True)

        # Merge
        df_final = merge_gl(df_gl, df_map)

        st.success("GL Mapping Completed Successfully!")
        st.dataframe(df_final)

        # Download
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='GL_Translated')
        buffer.seek(0)

        st.download_button(
            "Download Translated GL File",
            data=buffer,
            file_name="translated_gl.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")
