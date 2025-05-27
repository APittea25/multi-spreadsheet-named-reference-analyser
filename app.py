import streamlit as st
from openpyxl import load_workbook
import graphviz
import io
import re

st.set_page_config(page_title="Full Excel Formula Audit", layout="wide")

# --- Clean formula: remove external links like 'C:\\...\\file.xlsx'! or '[file.xlsx]Sheet'! ---
def clean_formula(formula):
    if not formula:
        return ""
    return re.sub(r"(?:'[^']*\.xlsx'!|'[^']*'!|\[[^\]]+\][^!]*!)", "", formula)

# --- Extract formulas and dependencies ---
def extract_all_formulas(wb):
    formulas = {}
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(values_only=False):
            for cell in row:
                if cell.value and str(cell.value).strip().startswith("="):
                    cell_id = f"{sheet.title}!{cell.coordinate}"
                    formula = clean_formula(str(cell.value).strip())
                    formulas[cell_id] = formula
    return formulas

# --- Extract dependencies from formula text ---
def find_formula_dependencies(formulas):
    dependencies = {}
    labels = list(formulas.keys())

    for target_cell, formula in formulas.items():
        formula_upper = formula.upper()
        refs = []
        for other_cell in labels:
            if other_cell == target_cell:
                continue
            other_label = other_cell.split("!")[-1].upper()
            if re.search(rf"\b{re.escape(other_label)}\b", formula_upper):
                refs.append(other_cell)
        dependencies[target_cell] = refs

    return dependencies

# --- Graph building ---
def create_dependency_graph(dependencies, all_nodes):
    dot = graphviz.Digraph()
    for node in all_nodes:
        dot.node(node)
    for target, sources in dependencies.items():
        for source in sources:
            dot.edge(source, target)
    return dot

# --- Streamlit UI ---
st.title("📊 Full Excel Formula Dependency Auditor")

uploaded_file = st.file_uploader("Upload an Excel (.xlsx) file", type=["xlsx"])

if uploaded_file:
    try:
        wb = load_workbook(io.BytesIO(uploaded_file.read()), data_only=False)

        st.subheader("📌 Extracted Formulas")
        all_formulas = extract_all_formulas(wb)
        st.json(all_formulas)

        st.subheader("🔗 Dependency Graph")
        dependencies = find_formula_dependencies(all_formulas)
        dot = create_dependency_graph(dependencies, all_formulas.keys())
        st.graphviz_chart(dot)

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("⬆️ Upload a `.xlsx` file to begin.")
