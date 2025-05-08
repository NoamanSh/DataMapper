import streamlit as st
import pandas as pd
import tempfile
import os
import json
import base64
from io import BytesIO

from process import extract_data_from_one_xml, write_dataframes_to_excel # Updated import
from utils import extract_excel_columns, extract_xml_tags

st.set_page_config(page_title="XML to Excel Mapping Tool", layout="wide")
st.title("🧩 XML to Excel Column Mapping Tool")

# Initialize session state
base_keys = [
    'excel_temp_path', 'output_excel_path', 'excel_columns', 
    'last_excel_file_id', 'mappings', 'active_xml_source_id_for_mapping'
]

for key in base_keys:
    if key not in st.session_state:
        st.session_state[key] = None

if 'flat_excel_columns' not in st.session_state: 
    st.session_state['flat_excel_columns'] = []
if 'xml_sources_data' not in st.session_state: # Will store list of dicts for each XML
    st.session_state['xml_sources_data'] = []
if 'processed_xml_file_ids' not in st.session_state: # To track already processed files
    st.session_state['processed_xml_file_ids'] = set()
if 'next_xml_internal_id_counter' not in st.session_state:
    st.session_state['next_xml_internal_id_counter'] = 1

# Upload files
uploaded_xml_files = st.file_uploader("📄 Upload XML files", type="xml", accept_multiple_files=True)
excel_file = st.file_uploader("📊 Upload Excel file with target columns", type=["xlsx"])

# Handle file uploads
if excel_file: # Excel file is mandatory to proceed
    # Check if Excel file has changed or if there are new XML files to process
    excel_changed = excel_file.file_id != st.session_state.get('last_excel_file_id')
    
    # If Excel changes, all XMLs need to be available for re-mapping if desired, but retain their data.
    # Mappings are reset if Excel changes.
    if excel_changed:
        st.info("Processing new files...")
        # Clean up previous temp files if they exist
        if st.session_state.get('excel_temp_path') and os.path.exists(st.session_state['excel_temp_path']):
            try:
                os.remove(st.session_state['excel_temp_path'])
            except OSError as e:
                st.warning(f"Could not remove old temp Excel file: {e}")
            st.session_state['excel_temp_path'] = None

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
            tmp_excel.write(excel_file.getvalue())
            st.session_state['excel_temp_path'] = tmp_excel.name
        st.session_state['last_excel_file_id'] = excel_file.file_id

        excel_data = extract_excel_columns(st.session_state['excel_temp_path'])
        st.session_state['excel_columns'] = excel_data
        flat_columns = [f"{sheet}/{col}" for sheet, cols in excel_data.items() for col in cols]
        st.session_state['flat_excel_columns'] = flat_columns
        
        st.session_state['mappings'] = [] # Reset mappings if Excel file changes
        st.session_state['output_excel_path'] = None
        st.success("Excel file processed. Upload or re-confirm XML files for mapping.")
        # Do not rerun yet, allow XML processing below

    # Process uploaded XML files
    if uploaded_xml_files:
        for xml_file_obj in uploaded_xml_files:
            if xml_file_obj.file_id not in st.session_state.processed_xml_file_ids:
                st.info(f"Processing new XML file: {xml_file_obj.name}...")
                temp_xml_path = None
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xml") as tmp_xml:
                    tmp_xml.write(xml_file_obj.getvalue())
                    temp_xml_path = tmp_xml.name
                
                xml_tags = extract_xml_tags(temp_xml_path)
                xml_internal_id = f"xml_{st.session_state.next_xml_internal_id_counter}"
                st.session_state.xml_sources_data.append({
                    'id': xml_internal_id,
                    'display_name': f"XML {st.session_state.next_xml_internal_id_counter} ({xml_file_obj.name})",
                    'temp_path': temp_xml_path,
                    'tags': xml_tags,
                    'file_id': xml_file_obj.file_id 
                })
                st.session_state.processed_xml_file_ids.add(xml_file_obj.file_id)
                st.session_state.next_xml_internal_id_counter += 1
                st.success(f"XML file '{xml_file_obj.name}' processed and added as source.")
                if not st.session_state.active_xml_source_id_for_mapping and st.session_state.xml_sources_data:
                    st.session_state.active_xml_source_id_for_mapping = st.session_state.xml_sources_data[0]['id']
                st.rerun() # Rerun to update UI with new XML source

    # Initial message if no XMLs are processed yet with a valid Excel
    if not st.session_state.xml_sources_data and st.session_state.excel_temp_path:
        st.info("Please upload one or more XML files to begin mapping.")


# Show interactive mapping interface if Excel is processed and at least one XML source exists
if st.session_state.get("flat_excel_columns") and st.session_state.get("xml_sources_data"):
    st.subheader("🔗 Mapping Interface")

    excel_nodes = st.session_state['flat_excel_columns']

    # Selector for active XML source for mapping
    xml_source_options = {source['id']: source['display_name'] for source in st.session_state.xml_sources_data}
    
    if not xml_source_options:
        st.warning("No XML sources available. Please upload XML files.")
        st.stop()

    # Ensure active_xml_source_id_for_mapping is valid
    if st.session_state.active_xml_source_id_for_mapping not in xml_source_options:
        st.session_state.active_xml_source_id_for_mapping = next(iter(xml_source_options)) # Default to first available

    selected_xml_source_id = st.radio(
        "Select XML Source for Tag Selection:",
        options=list(xml_source_options.keys()),
        format_func=lambda x: xml_source_options[x],
        key='active_xml_source_id_for_mapping_radio', # Use a different key for radio
        horizontal=True,
        index=list(xml_source_options.keys()).index(st.session_state.active_xml_source_id_for_mapping) # Ensure correct default
    )
    st.session_state.active_xml_source_id_for_mapping = selected_xml_source_id

    # Get tags for the currently active XML source
    active_xml_tags = []
    for source in st.session_state.xml_sources_data:
        if source['id'] == st.session_state.active_xml_source_id_for_mapping:
            active_xml_tags = source['tags']
            break

    # Debug information
    with st.expander("Debug Information"):
        st.write("Excel columns (first 5):", excel_nodes[:5])
        st.write(f"XML tags from {xml_source_options[st.session_state.active_xml_source_id_for_mapping]} (first 5):", active_xml_tags[:5])
        st.write("Current XML Sources Data:", st.session_state.xml_sources_data)
        st.write("Current Mappings:", st.session_state.mappings)
    
    # Create two columns for the mapping interface
    col1, col2 = st.columns(2)
    
    # Excel columns selection
    with col1:
        st.subheader("Excel Columns")
        selected_excel = st.selectbox("Select Excel Column", excel_nodes)
    
    # XML tags selection
    with col2:
        st.subheader("XML Tags")
        selected_xml_tag = st.selectbox(f"Select XML Tag (from {xml_source_options[st.session_state.active_xml_source_id_for_mapping]})", active_xml_tags)
    
    # Button to add mapping
    if st.button("Add Mapping"):
        # Check if this Excel column is already mapped
        excel_already_mapped = False
        for i, mapping in enumerate(st.session_state['mappings']):
            if mapping['excel'] == selected_excel:
                # Update the existing mapping
                st.session_state['mappings'][i]['xml_path'] = selected_xml_tag
                st.session_state['mappings'][i]['xml_source_id'] = st.session_state.active_xml_source_id_for_mapping
                excel_already_mapped = True
                st.success(f"Updated mapping: {selected_excel} → {selected_xml_tag} (from {xml_source_options[st.session_state.active_xml_source_id_for_mapping]})")
                break
        
        if not excel_already_mapped:
            # Add new mapping
            st.session_state['mappings'].append({
                'excel': selected_excel,
                'xml_path': selected_xml_tag,
                'xml_source_id': st.session_state.active_xml_source_id_for_mapping # Store which XML this mapping belongs to
            })
            st.success(f"Added mapping: {selected_excel} → {selected_xml_tag} (from {xml_source_options[st.session_state.active_xml_source_id_for_mapping]})")
    
    # Display current mappings
    if st.session_state['mappings']:
        st.subheader("Current Mappings")
        # Enhance DataFrame display to show XML source
        display_mappings = []
        for m in st.session_state['mappings']:
            source_display_name = xml_source_options.get(m['xml_source_id'], "Unknown XML Source")
            display_mappings.append({
                "Excel Column": m['excel'],
                "XML Path": m['xml_path'],
                "XML Source": source_display_name
            })
        mapping_data = pd.DataFrame(display_mappings)
        st.dataframe(mapping_data)
        
        # Button to remove a mapping
        if st.button("Remove Last Mapping"):
            if st.session_state['mappings']:
                st.session_state['mappings'].pop()
                st.success("Removed last mapping")
                st.rerun()
    
    # Process mappings and generate Excel
    if st.button("📥 Generate Excel"):
        current_mappings = st.session_state.get("mappings")
        if isinstance(current_mappings, list) and current_mappings:
            try:
                # Prepare mappings for each XML source
                # mappings_by_source will be: {'xml_1': {'Sheet1': {'ColA': 'path'}}, 'xml_2': ...}
                mappings_by_source_id = {}
                for m_item in current_mappings:
                    if '/' in m_item['excel']:
                        sheet, col = m_item['excel'].split('/', 1)
                        source_id_for_mapping = m_item['xml_source_id'] 
                        
                        if source_id_for_mapping not in mappings_by_source_id:
                            mappings_by_source_id[source_id_for_mapping] = {}
                        if sheet not in mappings_by_source_id[source_id_for_mapping]:
                            mappings_by_source_id[source_id_for_mapping][sheet] = {}
                        mappings_by_source_id[source_id_for_mapping][sheet][col] = m_item['xml_path']
                
                # Ensure directory exists
                os.makedirs('data', exist_ok=True)
                output_excel_file_path = 'data/output_filled.xlsx'

                # Extract data from each XML source that has mappings
                all_extracted_data_by_sheet_and_source = [] # List of dicts: [{'sheet_name': df_for_xml1}, {'sheet_name': df_for_xml2}]

                # Step 1: Determine max rows for each sheet across all relevant XML sources
                max_rows_per_sheet = {}
                for xml_s_data in st.session_state.xml_sources_data:
                    current_source_id = xml_s_data['id']
                    if current_source_id in mappings_by_source_id:
                        # Temporarily extract data to find max_len for each sheet from this source
                        temp_data_for_this_xml = extract_data_from_one_xml(
                            xml_s_data['temp_path'],
                            mappings_by_source_id[current_source_id]
                        )
                        for sheet_name, df_temp in temp_data_for_this_xml.items():
                            max_rows_per_sheet[sheet_name] = max(max_rows_per_sheet.get(sheet_name, 0), len(df_temp))

                # Step 2: Extract data again and reindex to global max_rows for that sheet
                for xml_s_data in st.session_state.xml_sources_data:
                    current_source_id = xml_s_data['id']
                    if current_source_id in mappings_by_source_id:
                        st.write(f"Extracting data for {xml_s_data['display_name']}...")
                        data_for_this_xml = extract_data_from_one_xml(
                            xml_s_data['temp_path'],
                            mappings_by_source_id[current_source_id]
                        )
                        # Reindex DFs in data_for_this_xml
                        reindexed_data_for_this_xml = {}
                        for sheet_name, df_original in data_for_this_xml.items():
                            if sheet_name in max_rows_per_sheet:
                                target_rows = max_rows_per_sheet[sheet_name]
                                if len(df_original) < target_rows:
                                    # Create an index with the target number of rows
                                    new_index = pd.RangeIndex(start=0, stop=target_rows, step=1)
                                    reindexed_data_for_this_xml[sheet_name] = df_original.reindex(new_index)
                                else:
                                    reindexed_data_for_this_xml[sheet_name] = df_original
                            else: # Should not happen if max_rows_per_sheet was populated correctly
                                reindexed_data_for_this_xml[sheet_name] = df_original
                        all_extracted_data_by_sheet_and_source.append(reindexed_data_for_this_xml)
                
                # Step 3: Merge data from all XML sources
                final_data_to_write_by_sheet = {}
                for data_from_one_source in all_extracted_data_by_sheet_and_source: # data_from_one_source is {'Sheet1': df, 'Sheet2': df}
                    for sheet_name, df_from_current_source in data_from_one_source.items():
                        if sheet_name not in final_data_to_write_by_sheet:
                            final_data_to_write_by_sheet[sheet_name] = df_from_current_source.copy()
                        else:
                            df_existing = final_data_to_write_by_sheet[sheet_name]
                            for col_to_add in df_from_current_source.columns:
                                if col_to_add not in df_existing.columns:
                                    df_existing[col_to_add] = df_from_current_source[col_to_add]
                            final_data_to_write_by_sheet[sheet_name] = df_existing

                write_dataframes_to_excel(st.session_state['excel_temp_path'], final_data_to_write_by_sheet, output_excel_file_path)
                st.session_state['output_excel_path'] = output_excel_file_path
                st.success("✅ Excel generated successfully.")
                
                # Download button logic
                output_file_to_download = st.session_state['output_excel_path'] # Use the variable directly
                if os.path.exists(output_file_to_download):
                    with open(output_file_to_download, "rb") as f:
                        file_bytes = f.read()
                    st.download_button(
                        "⬇️ Download Mapped Excel", 
                        file_bytes, 
                        file_name="output_filled.xlsx", # Keep consistent filename
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error(f"Generated file not found at {output_file_to_download}")
            except Exception as e:
                st.error(f"❌ Failed to generate Excel: {str(e)}")
                st.exception(e) # Show full traceback for debugging
        else:
            st.warning("⚠️ Please upload and map at least one XML file first.")
elif not excel_file:
    st.info("📊 Please upload an Excel file to begin.")