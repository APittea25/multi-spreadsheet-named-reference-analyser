import streamlit as st
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Excel Named Reference Analyzer", layout="wide")
st.title("🔍 Excel Named Reference Analyzer")

# Upload Excel files
uploaded_files = st.file_uploader("Upload Excel files (.xlsx)", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    all_named_references = []

    for uploaded_file in uploaded_files:
        st.subheader(f"📄 File: {uploaded_file.name}")

        try:
            in_memory_file = BytesIO(uploaded_file.read())
            wb = load_workbook(in_memory_file, data_only=False)

            file_named_refs = []

            for defined_name in wb.defined_names.definedName:
                name = defined_name.name
                destinations = list(defined_name.destinations)  # (sheet_name, range)

                if destinations:
                    for sheet_name, cell_range in destinations:
                        formula = f"{sheet_name}!{cell_range}"
                        scope = sheet_name
                        file_named_refs.append((name, formula, scope))
                else:
                    formula = defined_name.attr_text
                    scope = "Workbook-level"
                    file_named_refs.append((name, formula, scope))

            if file_named_refs:
                st.write("🔗 Named References Found:")
                for name, formula, scope in file_named_refs:
                    st.code(f"{name} ({scope}) = {formula}", language="text")
                    all_named_references.append({
                        "file": uploaded_file.name,
                        "name": name,
                        "formula": formula,
                        "scope": scope
                    })
            else:
                st.warning("⚠️ No named references found in this file.")

        except Exception as e:
            st.error(f"❌ Error reading {uploaded_file.name}: {e}")

    st.success(f"✅ Extracted {len(all_named_references)} named references from {len(uploaded_files)} file(s).")

else:
    st.info("⬆️ Upload one or more `.xlsx` Excel files to begin.")
