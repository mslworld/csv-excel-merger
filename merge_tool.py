import streamlit as st
import pandas as pd
import os
import tempfile
import io

st.set_page_config(page_title="CSV/Excel Merger Tool", layout="centered")
st.title("📁 Merge CSV, Excel, and TXT Files into One CSV")

# Pagination settings
FILES_PER_PAGE = 20
CHUNK_SIZE = 100000  # rows per chunk for large CSV/TXT

uploaded_files = st.file_uploader(
    "Upload multiple CSV, Excel, or TXT files (any size, any mix)",
    type=["csv", "xls", "xlsx", "txt"],
    accept_multiple_files=True,
    key="file_uploader"
)

if uploaded_files:
    page_number = st.number_input("Page", min_value=1, max_value=(len(uploaded_files) - 1) // FILES_PER_PAGE + 1, value=1, step=1)
    start_index = (page_number - 1) * FILES_PER_PAGE
    end_index = start_index + FILES_PER_PAGE

    st.markdown("### Uploaded Files:")
    for file in uploaded_files[start_index:end_index]:
        st.write(f"📄 {file.name} ({round(file.size / 1024 / 1024, 2)} MB)")

    if st.button("🔄 Merge Files"):
        merged_df = pd.DataFrame()

        with st.spinner("Merging files. Please wait..."):
            for uploaded_file in uploaded_files:
                try:
                    filename = uploaded_file.name
                    file_buffer = io.BytesIO(uploaded_file.read())
                    file_buffer.seek(0)

                    if filename.endswith('.csv') or filename.endswith('.txt'):
                        delimiter = '\t' if filename.endswith('.txt') else ','
                        chunk_iter = pd.read_csv(file_buffer, delimiter=delimiter, chunksize=CHUNK_SIZE)
                        for chunk in chunk_iter:
                            chunk.columns = [f"Column{i+1}" for i in range(len(chunk.columns))]
                            if merged_df.empty:
                                merged_df = chunk.copy()
                            else:
                                chunk.columns = merged_df.columns[:len(chunk.columns)]
                                merged_df = pd.concat([merged_df, chunk], ignore_index=True)

                    elif filename.endswith(('.xls', '.xlsx')):
                        df = pd.read_excel(file_buffer, engine='openpyxl')
                        df.columns = [f"Column{i+1}" for i in range(len(df.columns))]
                        if merged_df.empty:
                            merged_df = df.copy()
                        else:
                            df.columns = merged_df.columns[:len(df.columns)]
                            merged_df = pd.concat([merged_df, df], ignore_index=True)

                    else:
                        st.warning(f"Unsupported file type: {filename}")

                except Exception as e:
                    st.error(f"Error processing {uploaded_file.name}: {e}")

        if not merged_df.empty:
            temp_dir = tempfile.mkdtemp()
            output_path = os.path.join(temp_dir, "merged_output.csv")
            merged_df.to_csv(output_path, index=False)

            st.success("✅ Files merged successfully!")
            with open(output_path, "rb") as f:
                st.download_button("📥 Download Merged CSV", f, file_name="merged_output.csv")
        else:
            st.warning("No data was merged. Please check the files.") 
 


