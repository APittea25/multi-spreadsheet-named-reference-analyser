import streamlit as st
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Excel Named Reference Analyzer", layout="wide")
st.title("🔍 Excel Named Reference Analyzer")

# Step 1: Upload Excel files
uploaded_files = st.file_uploader("Upload one or more Excel files (.xlsx)", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    all_named_references = []

    for uploaded_file in uploaded_files:
        st.subheader(f"📄 File: {uploaded_file.name}")

        try:
            # Load the file in memory
            in_memory_file = BytesIO(uploaded_file.read())
            wb = load_workbook(in_memory_file, data_only=False)

            # Extract all named ranges
            named_ranges = wb.defined_names.defined_names
            file_named_refs = []

            for defined_name in named_ranges:
                name = defined_name.name
                formula = defined_name.attr_text
                scope = "Workbook-level" if defined_name.scope is None else wb.sheetnames[defined_name.scope]
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

    st.success(f"✅ Extracted {len(all_named_references)} named references across {len(uploaded_files)} file(s).")

else:
    st.info("⬆️ Please upload at least one `.xlsx` Excel file to begin.")
