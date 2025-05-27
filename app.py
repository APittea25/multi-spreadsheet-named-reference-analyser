import streamlit as st
from openai import OpenAI
from openpyxl import load_workbook
import graphviz
import io
import re

st.set_page_config(page_title="Excel Named Reference Analyzer", layout="wide")

# Initialize OpenAI client securely
openai_api_key = st.secrets.get("OPENAI_API_KEY")
if not openai_api_key:
    st.error("❌ OPENAI_API_KEY not found in secrets.")
    st.stop()
client = OpenAI(api_key=openai_api_key)

# --- Extract named references from a workbook ---
def extract_named_references(wb, file_label):
    named_refs = {}
    for name in wb.defined_names:
        defined_name = wb.defined_names[name]
        if defined_name.attr_text and not defined_name.is_external:
            dests = list(defined_name.destinations)
            for sheet_name, ref in dests:
                full_name = f"{file_label}::{defined_name.name}"
                named_refs[full_name] = {
                    "sheet": sheet_name,
                    "ref": ref,
                    "formula": None,
                    "file": file_label,
                    "pure_name": defined_name.name
                }
                try:
                    sheet = wb[sheet_name]
                    cell_ref = ref.split('!')[-1]
                    cell = sheet[cell_ref]
                    if cell.data_type == 'f':
                        named_refs[full_name]["formula"] = cell.value
                except Exception:
                    pass
    return named_refs

# --- Find dependencies between named references by name ---
def find_dependencies(named_refs):
    dependencies = {}

    # Map pure names (e.g., Annuity_Type) → [full names]
    pure_to_full = {}
    for full_name, info in named_refs.items():
        pure = info["pure_name"].upper()
        pure_to_full.setdefault(pure, []).append(full_name)

    for current_name, info in named_refs.items():
        formula = info.get("formula", "") or ""
        formula_upper = formula.upper()
        deps = []

        for pure_name, full_names in pure_to_full.items():
            if current_name in full_names:
                continue  # Skip self

            # Match pure name as whole word in formula
            if re.search(rf"\b{re.escape(pure_name)}\b", formula_upper):
                deps.extend(full_names)

        dependencies[current_name] = deps

    return dependencies

# --- Create dependency graph ---
def create_dependency_graph(dependencies):
    dot = graphviz.Digraph()
    for node in dependencies:
        dot.node(node)
    for node, deps in dependencies.items():
        for dep in deps:
            dot.edge(dep, node)
    return dot

# --- OpenAI explanation ---
@st.cache_data(show_spinner=False)
def call_openai(prompt, max_tokens=100):
    try:
        response = client.chat.completions.create(
            model="gpt-4",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2,
            max_tokens=max_tokens
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"(Error: {e})"

# --- Generate AI explanations for each formula ---
@st.cache_data(show_spinner=False)
def generate_ai_outputs(named_refs):
    results = []
    for name, info in named_refs.items():
        excel_formula = info.get("formula", "")
        if not excel_formula:
            doc = "No formula."
            py = ""
        else:
            doc_prompt = f"Explain what this Excel formula does:\n{excel_formula}"
            py_prompt = f"Translate this Excel formula into Python:\n{excel_formula}"
            doc = call_openai(doc_prompt, max_tokens=100)
            py = call_openai(py_prompt, max_tokens=100)

        results.append({
            "Named Reference": name,
            "AI Documentation": doc,
            "Excel Formula": excel_formula,
            "Python Formula": py,
        })
    return results

# --- Markdown table rendering ---
def render_markdown_table(rows):
    headers = ["Named Reference", "AI Documentation", "Excel Formula", "Python Formula"]
    md = "| " + " | ".join(headers) + " |\n"
    md += "| " + " | ".join(["---"] * len(headers)) + " |\n"
    for row in rows:
        md += "| " + " | ".join([
            str(row.get("Named Reference", "")),
            str(row.get("AI Documentation", "")).replace("\n", " "),
            str(row.get("Excel Formula", "")).replace("\n", " "),
            str(row.get("Python Formula", "")).replace("\n", " "),
        ]) + " |\n"
    return md

# --- Streamlit UI ---
st.title("📊 Excel Named Reference Dependency Viewer (Multi-Workbook)")

uploaded_files = st.file_uploader("Upload one or more Excel (.xlsx) files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    combined_named_refs = {}

    for uploaded_file in uploaded_files:
        try:
            wb = load_workbook(filename=io.BytesIO(uploaded_file.read()), data_only=False)
            file_named_refs = extract_named_references(wb, uploaded_file.name)
            combined_named_refs.update(file_named_refs)
        except Exception as e:
            st.error(f"❌ Failed to read {uploaded_file.name}: {e}")

    if combined_named_refs:
        st.subheader("📌 Named References Extracted")
        st.json(combined_named_refs)

        st.subheader("🔗 Dependency Graph")
        dependencies = find_dependencies(combined_named_refs)
        dot = create_dependency_graph(dependencies)
        st.graphviz_chart(dot)

        st.subheader("🧠 AI-Generated Formula Explanations")
        with st.spinner("Asking GPT for help..."):
            table_rows = generate_ai_outputs(combined_named_refs)
            st.markdown(render_markdown_table(table_rows), unsafe_allow_html=True)
    else:
        st.warning("No named references found.")
else:
    st.info("⬆️ Upload `.xlsx` files to get started.")
