import streamlit as st
from openai import OpenAI
from openpyxl import load_workbook
import graphviz
import io
import re

st.set_page_config(page_title="Excel Named Reference Dependency Viewer", layout="wide")

# --- OpenAI setup ---
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
            for sheet_name, ref in defined_name.destinations:
                full_name = f"{file_label}::{defined_name.name}"
                named_refs[full_name] = {
                    "sheet": sheet_name,
                    "ref": ref,
                    "formula": None,
                    "file": file_label,
                    "label": defined_name.name  # the actual name used in formulas
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

# --- Detect dependencies: if a named reference is used in another formula, it's a dependency ---
def find_dependencies(named_refs):
    dependencies = {}
    all_names = list(named_refs.keys())

    for current_name, info in named_refs.items():
        formula = (info.get("formula") or "").upper()
        deps = []

        for other_name in all_names:
            if other_name == current_name:
                continue

            other_label = named_refs[other_name]["label"].upper()

            if re.search(rf"\b{re.escape(other_label)}\b", formula):
                deps.append(other_name)

        dependencies[current_name] = deps

    return dependencies

# --- Graphviz graph generation ---
def create_dependency_graph(dependencies):
    dot = graphviz.Digraph()
    for node in dependencies:
        dot.node(node)
    for node, deps in dependencies.items():
        for dep in deps:
            dot.edge(dep, node)
    return dot

# --- GPT explanation functions ---
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

@st.cache_data(show_spinner=False)
def generate_ai_outputs(named_refs):
    results = []
    for name, info in named_refs.items():
        formula = info.get("formula", "")
        if not formula:
            doc = "No formula."
            py = ""
        else:
            doc = call_openai(f"Explain this Excel formula:\n{formula}")
            py = call_openai(f"Translate this Excel formula to Python:\n{formula}")
        results.append({
            "Named Reference": name,
            "AI Documentation": doc,
            "Excel Formula": formula,
            "Python Formula": py
        })
    return results

# --- Markdown table rendering ---
def render_markdown_table(rows):
    headers = ["Named Reference", "AI Documentation", "Excel Formula", "Python Formula"]
    md = "| " + " | ".join(headers) + " |\n"
    md += "| " + " | ".join(["---"] * len(headers)) + " |\n"
    for row in rows:
        md += "| " + " | ".join([
            str(row["Named Reference"]),
            str(row["AI Documentation"]).replace("\n", " "),
            str(row["Excel Formula"]).replace("\n", " "),
            str(row["Python Formula"]).replace("\n", " ")
        ]) + " |\n"
    return md

# --- Streamlit UI ---
st.title("📊 Excel Named Reference Dependency Viewer (Multi-Workbook)")

uploaded_files = st.file_uploader("Upload Excel files (.xlsx)", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    combined_named_refs = {}

    for uploaded_file in uploaded_files:
        try:
            wb = load_workbook(io.BytesIO(uploaded_file.read()), data_only=False)
            refs = extract_named_references(wb, uploaded_file.name)
            combined_named_refs.update(refs)
        except Exception as e:
            st.error(f"❌ Error reading {uploaded_file.name}: {e}")

    if combined_named_refs:
        st.subheader("📌 Named References Extracted")
        st.json(combined_named_refs)

        st.subheader("🔗 Dependency Graph")
        dependencies = find_dependencies(combined_named_refs)
        dot = create_dependency_graph(dependencies)
        st.graphviz_chart(dot)

        st.subheader("🧠 AI Explanations")
        with st.spinner("Calling GPT-4..."):
            rows = generate_ai_outputs(combined_named_refs)
            st.markdown(render_markdown_table(rows), unsafe_allow_html=True)
    else:
        st.warning("No named references found.")
else:
    st.info("⬆️ Upload one or more `.xlsx` files to begin.")
