import streamlit as st
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Excel Named Reference Analyzer", layout="wide")
st.title("🔍 Excel Named Reference Analyzer")

# Upload multiple Excel files
uploaded_files = st.file_uploader("Upload Excel files (.xlsx)", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    all_named_references = []

    for uploaded_file in uploaded_files:
        st.subheader(f"📄 File: {uploaded_file.name}")

        try:
            # Load the file into memory for openpyxl
            in_memory_file = BytesIO(uploaded_file.read())
            wb = load_workbook(in_memory_file, data_only=False)

            file_named_refs = []

            # ✅ Correct way to iterate over defined names
            for defined_name in wb.defined_names.definedName:
                name = defined_name.name
                formula = defined_name.attr_text

                # Determine if it's workbook or sheet scope
                if defined_name.scope is None:
                    scope = "Workbook-level"
                else:
                    try:
                        scope = wb.sheetnames[defined_name.scope]
                    except IndexError:
                        scope = f"Sheet index {defined_name.scope} (invalid or missing)"

                file_named_refs.append((name, formula, scope))

            # Display named references
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
    st.info("⬆️ Please upload one or more `.xlsx` Excel files to begin.")
