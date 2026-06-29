import streamlit as st
import pandas as pd
import os
import time
import io
import tempfile
import csv
from datetime import datetime
import math

# Set Page Config
st.set_page_config(
    page_title="Pro Spreadsheet Merger",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for Professional Look
st.markdown("""
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .stButton>button {
        width: 100%;
        border-radius: 5px;
        height: 3em;
        background-color: #007bff;
        color: white;
    }
    .stProgress > div > div > div > div {
        background-color: #28a745;
    }
    .reportview-container .main .block-container {
        padding-top: 2rem;
    }
    .metric-container {
        background-color: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    </style>
""", unsafe_allow_html=True)

# Helper Constants
SUPPORTED_EXTENSIONS = [".csv", ".txt", ".xls", ".xlsx", ".xml", ".ods", ".dif", ".slk", ".prn"]

def get_file_info(uploaded_file):
    """Extract metadata from uploaded file."""
    name = uploaded_file.name
    ext = os.path.splitext(name)[1].lower()
    size_kb = uploaded_file.size / 1024
    return {
        "name": name,
        "extension": ext,
        "size": size_kb,
        "raw": uploaded_file
    }

def detect_delimiter(file_obj):
    """Attempt to detect CSV delimiter."""
    try:
        sample = file_obj.read(2048).decode('utf-8', errors='ignore')
        file_obj.seek(0)
        sniffer = csv.Sniffer()
        return sniffer.sniff(sample).delimiter
    except:
        file_obj.seek(0)
        return ','

def get_all_columns(files):
    """
    Scan all files briefly to determine the union of all column names.
    This ensures proper alignment during streaming.
    """
    all_cols = []
    seen = set()
    
    status_text = st.empty()
    status_text.text("Analyzing column structures...")
    
    for f_info in files:
        ext = f_info['extension']
        f = f_info['raw']
        try:
            # We only read the first row to get columns
            if ext in ['.csv', '.txt']:
                sep = detect_delimiter(f)
                df = pd.read_csv(f, sep=sep, nrows=0)
            elif ext in ['.xls', '.xlsx']:
                df = pd.read_excel(f, nrows=0)
            elif ext == '.ods':
                df = pd.read_excel(f, engine='odf', nrows=0)
            elif ext == '.xml':
                df = pd.read_xml(f, nrows=0)
            else:
                # Fallback for other formats, try reading small chunk
                df = pd.read_csv(f, nrows=0)
            
            for col in df.columns:
                if col not in seen:
                    all_cols.append(col)
                    seen.add(col)
            f.seek(0)
        except Exception:
            f.seek(0)
            continue
            
    status_text.empty()
    return all_cols

def process_merge(files, output_path, all_columns):
    """
    High-performance streaming merge logic.
    """
    total_files = len(files)
    processed_files = 0
    skipped_files = []
    total_rows = 0
    start_time = time.time()
    
    # Progress UI placeholders
    progress_bar = st.progress(0)
    status_label = st.empty()
    metrics_row = st.columns(4)
    m1, m2, m3, m4 = metrics_row
    
    # Open the output file for writing
    with open(output_path, 'w', newline='', encoding='utf-8') as out_f:
        writer = csv.DictWriter(out_f, fieldnames=all_columns)
        writer.writeheader()
        
        for idx, f_info in enumerate(files):
            filename = f_info['name']
            ext = f_info['extension']
            f = f_info['raw']
            f.seek(0)
            
            status_label.info(f"Processing ({idx+1}/{total_files}): {filename}")
            
            try:
                # Determine how to read the file
                if ext in ['.csv', '.txt']:
                    sep = detect_delimiter(f)
                    # Use chunking for large text files
                    chunks = pd.read_csv(f, sep=sep, chunksize=50000, low_memory=False)
                    for chunk_idx, chunk in enumerate(chunks):
                        # FIX: Remove index if it was read as a column
                        if 'Unnamed: 0' in chunk.columns:
                            chunk = chunk.drop(columns=['Unnamed: 0'])
                            
                        # Align columns: ensure data stays in correct named column
                        chunk = chunk.reindex(columns=all_columns)
                        
                        # FIX: index=False is critical to prevent shifting data to the right
                        chunk.to_csv(out_f, header=False, index=False, quoting=csv.QUOTE_MINIMAL)
                        
                        total_rows += len(chunk)
                        m3.metric("Rows Processed", f"{total_rows:,}")
                        
                else:
                    # For Excel/other non-streaming formats, we have to load fully 
                    # but we do it one file at a time to save RAM.
                    if ext in ['.xls', '.xlsx']:
                        df = pd.read_excel(f)
                    elif ext == '.ods':
                        df = pd.read_excel(f, engine='odf')
                    elif ext == '.xml':
                        df = pd.read_xml(f)
                    elif ext == '.slk':
                        # SYLK handling
                        df = pd.read_csv(f, sep=';', engine='python')
                    elif ext == '.prn':
                        # PRN handling (fixed width usually, but often treated as space-sep)
                        df = pd.read_csv(f, sep='\s+', engine='python')
                    else:
                        df = pd.read_csv(f)
                    
                    # FIX: Ensure we don't have an "Unnamed" index column from the original file
                    if 'Unnamed: 0' in df.columns:
                        df = df.drop(columns=['Unnamed: 0'])

                    # FIX: Align columns strictly and drop the index to prevent shifting
                    df = df.reindex(columns=all_columns)
                    df.to_csv(out_f, header=False, index=False, quoting=csv.QUOTE_MINIMAL)
                    total_rows += len(df)
                
                processed_files += 1
                
            except Exception as e:
                skipped_files.append({"file": filename, "error": str(e)})
            
            # Update overall progress
            elapsed = time.time() - start_time
            progress = (idx + 1) / total_files
            progress_bar.progress(progress)
            
            avg_time_per_file = elapsed / (idx + 1)
            remaining_files = total_files - (idx + 1)
            est_remaining = avg_time_per_file * remaining_files
            
            m1.metric("Files Done", f"{processed_files}/{total_files}")
            m2.metric("Elapsed", f"{elapsed:.1f}s")
            m4.metric("Est. Left", f"{est_remaining:.1f}s")

    return processed_files, skipped_files, total_rows, time.time() - start_time

def main():
    st.title("🚀 Professional Spreadsheet Merger")
    st.markdown("Merge CSV, Excel, ODS, and more with high-performance streaming.")

    if 'uploaded_files_list' not in st.session_state:
        st.session_state.uploaded_files_list = []

    # Sidebar / Top Actions
    with st.expander("⬆️ Upload Files", expanded=True):
        new_files = st.file_uploader(
            "Drag and drop files here", 
            accept_multiple_files=True,
            type=[ext.strip('.') for ext in SUPPORTED_EXTENSIONS]
        )
        
        if new_files:
            for f in new_files:
                if f.name not in [x['name'] for x in st.session_state.uploaded_files_list]:
                    st.session_state.uploaded_files_list.append(get_file_info(f))

    if st.session_state.uploaded_files_list:
        # Action Buttons
        col_btns = st.columns([1, 1, 4])
        if col_btns[0].button("🗑️ Clear All"):
            st.session_state.uploaded_files_list = []
            st.rerun()
            
        # Search and Filter
        search_query = st.text_input("🔍 Search files by name...", "")
        
        filtered_files = [
            f for f in st.session_state.uploaded_files_list 
            if search_query.lower() in f['name'].lower()
        ]
        
        # Pagination
        items_per_page = 20
        total_pages = math.ceil(len(filtered_files) / items_per_page)
        page = st.number_input("Page", min_value=1, max_value=max(1, total_pages), step=1)
        
        start_idx = (page - 1) * items_per_page
        end_idx = start_idx + items_per_page
        
        # Display File List
        st.subheader(f"Files to Merge ({len(filtered_files)})")
        
        # Summary Header
        total_size = sum(f['size'] for f in filtered_files)
        st.info(f"Total Files: **{len(filtered_files)}** | Total Size: **{total_size/1024:.2f} MB**")
        
        # Table Header
        h_col1, h_col2, h_col3, h_col4 = st.columns([4, 2, 2, 1])
        h_col1.write("**Filename**")
        h_col2.write("**Size (KB)**")
        h_col3.write("**Ext**")
        h_col4.write("**Action**")
        
        for i, f_info in enumerate(filtered_files[start_idx:end_idx]):
            c1, c2, c3, c4 = st.columns([4, 2, 2, 1])
            c1.write(f_info['name'])
            c2.write(f"{f_info['size']:.2f}")
            c3.write(f_info['extension'].upper())
            if c4.button("❌", key=f"del_{f_info['name']}"):
                st.session_state.uploaded_files_list = [
                    x for x in st.session_state.uploaded_files_list if x['name'] != f_info['name']
                ]
                st.rerun()

        st.divider()
        
        # Merge Section
        if st.button("🏗️ Start Merging Files", type="primary"):
            if not st.session_state.uploaded_files_list:
                st.error("No files to merge!")
                return

            with tempfile.NamedTemporaryFile(delete=False, suffix=".csv") as tmp_out:
                output_path = tmp_out.name
            
            try:
                # 1. Get Union of Columns
                all_columns = get_all_columns(st.session_state.uploaded_files_list)
                
                # 2. Process Merge
                processed, skipped, total_rows, duration = process_merge(
                    st.session_state.uploaded_files_list, 
                    output_path, 
                    all_columns
                )
                
                # 3. Success Summary
                st.success("✅ Merging Completed!")
                
                stats_col1, stats_col2, stats_col3 = st.columns(3)
                stats_col1.metric("Total Rows", f"{total_rows:,}")
                stats_col2.metric("Time Taken", f"{duration:.2f}s")
                out_size = os.path.getsize(output_path) / (1024 * 1024)
                stats_col3.metric("Output Size", f"{out_size:.2f} MB")
                
                # Download Button
                with open(output_path, "rb") as f:
                    st.download_button(
                        label="📥 Download Merged CSV",
                        data=f,
                        file_name=f"merged_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv"
                    )
                
                # Error Log
                if skipped:
                    with st.expander("⚠️ Skipped Files / Errors"):
                        for s in skipped:
                            st.error(f"**{s['file']}**: {s['error']}")
                            
            except Exception as e:
                st.error(f"A critical error occurred: {str(e)}")
            finally:
                # Cleanup handled by user downloading or session end
                pass
    else:
        st.info("Please upload some spreadsheet files to get started.")
        st.image("https://img.icons8.com/clouds/200/000000/data-configuration.png", width=200)

if __name__ == "__main__":
    main()
