import streamlit as st
from openai import OpenAI
from openpyxl import load_workbook
import graphviz
import io

st.set_page_config(page_title="Excel Named Range Visualizer", layout="wide")

# Initialize OpenAI client
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# --- Extract named references from a single workbook ---
def extract_named_references(wb, file_label):
    named_refs = {}
    for name in wb.defined_names:
        defined_name = wb.defined_names[name]
        if defined_name.attr_text and not defined_name.is_external:
            dests = list(defined_name.destinations)
            for sheet_name, ref in dests:
                full_name = f"{file_label}::{defined_name.name}"  # Unique name per file
                named_refs[full_name] = {
                    "sheet": sheet_name,
                    "ref": ref,
                    "formula": None,
                    "file": file_label
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

# --- Detect dependencies between named references ---
def find_dependencies(named_refs):
    dependencies = {}
    all_names = list(named_refs.keys())
    for name, info in named_refs.items():
        formula = info.get("formula", "")
        if formula:
            formula = formula.upper()
            deps = [other for other in all_names if other != name and other.split("::")[-1].upper() in formula]
            dependencies[name] = deps
        else:
            dependencies[name] = []
    return dependencies

# --- Create Graphviz dependency graph ---
def create_dependency_graph(dependencies):
    dot = graphviz.Digraph()
    for ref in dependencies:
        dot.node(ref)
    for ref, deps in dependencies.items():
        for dep in deps:
            dot.edge(dep, ref)
    return dot

# --- Call OpenAI GPT for formula interpretation ---
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

# --- Generate AI explanations ---
@st.cache_data(show_spinner=False)
def generate_ai_outputs(named_refs):
    results = []
    for name, info in named_refs.items():
        excel_formula = info.get("formula", "")
        if not excel_formula:
            doc = "No formula."
            py = ""
        else:
            doc_prompt = f"Explain what the following Excel formula does:\n{excel_formula}"
            py_prompt = f"Translate this Excel formula into a clean, readable Python expression:\n{excel_formula}"
            doc = call_openai(doc_prompt, max_tokens=100)
            py = call_openai(py_prompt, max_tokens=100)

        results.append({
            "Named Reference": name,
            "AI Documentation": doc,
            "Excel Formula": excel_formula,
            "Python Formula": py,
        })
    return results

# --- Markdown table formatter ---
def render_markdown_table(rows):
    headers = ["Named Reference", "AI Documentation", "Excel Formula", "Python Formula"]
    md = "| " + " | ".join(headers) + " |\n"
    md += "| " + " | ".join(["---"] * len(headers)) + " |\n"
    for row in rows:
        md += "| " + " | ".join([
            str(row.get("Named Reference", "") or ""),
            str(row.get("AI Documentation", "") or "").replace("\n", " "),
            str(row.get("Excel Formula", "") or "").replace("\n", " "),
            str(row.get("Python Formula", "") or "").replace("\n", " "),
        ]) + " |\n"
    return md

# --- Streamlit UI ---
st.title("📊 Multi-Workbook Named Range Dependency Viewer with AI")

uploaded_files = st.file_uploader("Upload one or more Excel (.xlsx) files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    combined_named_refs = {}

    for uploaded_file in uploaded_files:
        try:
            wb = load_workbook(filename=io.BytesIO(uploaded_file.read()), data_only=False)
            file_named_refs = extract_named_references(wb, uploaded_file.name)
            combined_named_refs.update(file_named_refs)
        except Exception as e:
            st.error(f"❌ Failed to process {uploaded_file.name}: {e}")

    if combined_named_refs:
        st.subheader("📌 Named References Across Files")
        st.json(combined_named_refs)

        st.subheader("🔗 Cross-Workbook Dependency Graph")
        dependencies = find_dependencies(combined_named_refs)
        dot = create_dependency_graph(dependencies)
        st.graphviz_chart(dot)

        st.subheader("🧠 AI-Generated Documentation and Python Translations")
        with st.spinner("Generating AI explanations using GPT..."):
            table_rows = generate_ai_outputs(combined_named_refs)
            st.markdown(render_markdown_table(table_rows), unsafe_allow_html=True)
    else:
        st.warning("⚠️ No named references were found in the uploaded files.")
else:
    st.info("⬆️ Please upload one or more `.xlsx` Excel files to begin.")
