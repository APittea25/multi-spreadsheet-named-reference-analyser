import streamlit as st
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Excel Named Reference Analyzer", layout="wide")
st.title("🔍 Excel Named Reference Analyzer")

# Step 1: Upload multiple Excel files
uploaded_files = st.file_uploader("Upload Excel files (.xlsx)", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    all_named_references = []

    for uploaded_file in uploaded_files:
        st.subheader(f"📄 File: {uploaded_file.name}")

        try:
            # Load workbook from uploaded file
            in_memory_file = BytesIO(uploaded_file.read())
            wb = load_workbook(in_memory_file, data_only=False)

            file_named_refs = []

            for defined_name in wb.defined_names.definedName:
                name = defined_name.name
                formula = defined_name.attr_text
                scope = defined_name.scope  # Can be used for sheet-level vs workbook-level
                file_named_refs.append((name, formula))

            if file_named_refs:
                st.write("🔗 Named References:")
                for name, formula in file_named_refs:
                    st.code(f"{name} = {formula}", language="text")
                all_named_references.extend([(uploaded_file.name, name, formula) for name, formula in file_named_refs])
            else:
                st.write("⚠️ No named ranges found.")

        except Exception as e:
            st.error(f"❌ Error reading {uploaded_file.name}: {e}")

    # Future step: Use `all_named_references` to build cross-file dependency graph

else:
    st.info("⬆️ Upload at least one `.xlsx` file to begin.")

