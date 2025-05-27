import streamlit as st
from openpyxl import load_workbook
import graphviz
import io
import re
from collections import defaultdict

st.set_page_config(page_title="Named Range Formula Dependency Viewer", layout="wide")

# --- Simplify formulas (remove external references) ---
def simplify_formula(formula):
    if not formula:
        return ""
    return re.sub(r"(?:'[^']*\.xlsx'!|'[^']*'!|\[[^\]]+\][^!]*!)", "", formula)

# --- Extract named ranges and top-left formulas only ---
def extract_named_references(wb, file_label):
    named_refs = {}
    for name in wb.defined_names:
        dn = wb.defined_names[name]
        if dn.attr_text and not dn.is_external:
            for sheet_name, ref in dn.destinations:
                sheet = wb[sheet_name]
                label = name
                try:
                    coord = ref.replace("$", "").split("!")[-1]
                    top_left_cell = coord.split(":")[0]  # get top-left if it's a range
                    cell = sheet[top_left_cell]
                    raw_value = str(cell.value or "").strip()
                    if raw_value.startswith("="):
                        named_refs[label] = {
                            "sheet": sheet_name,
                            "ref": ref,
                            "formula": simplify_formula(raw_value),
                            "file": file_label
                        }
                except Exception:
                    pass
    return named_refs

# --- Detect dependencies between formulas ---
def find_dependencies(named_refs):
    dependencies = defaultdict(list)
    labels = list(named_refs.keys())

    for target_label, info in named_refs.items():
        formula = (info.get("formula") or "").upper()
        for other_label in labels:
            if other_label == target_label:
                continue
            if re.search(rf"\b{re.escape(other_label.upper())}\b", formula):
                dependencies[target_label].append(other_label)
    return dependencies

# --- Graphviz generation ---
def create_dependency_graph(dependencies, all_labels):
    dot = graphviz.Digraph()
    for label in all_labels:
        dot.node(label)
    for target, sources in dependencies.items():
        for source in sources:
            dot.edge(source, target)
    return dot

# --- Streamlit UI ---
st.title("📊 Named Range Formula Dependency Viewer")

uploaded_file = st.file_uploader("Upload an Excel (.xlsx) file", type=["xlsx"])

if uploaded_file:
    try:
        wb = load_workbook(io.BytesIO(uploaded_file.read()), data_only=False)

        st.subheader("📌 Named References")
        named_refs = extract_named_references(wb, uploaded_file.name)
        st.json(named_refs)

        st.subheader("🔗 Dependency Graph")
        dependencies = find_dependencies(named_refs)
        dot = create_dependency_graph(dependencies, named_refs.keys())
        st.graphviz_chart(dot)

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("⬆️ Upload a `.xlsx` file to begin.")
