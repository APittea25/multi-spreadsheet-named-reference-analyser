import streamlit as st
from openai import OpenAI
from openpyxl import load_workbook
import graphviz
import io
import re

st.set_page_config(page_title="Named Reference Dependency Viewer", layout="wide")

# --- OpenAI setup ---
openai_api_key = st.secrets.get("OPENAI_API_KEY")
if not openai_api_key:
    st.error("❌ OPENAI_API_KEY not found in secrets.")
    st.stop()
client = OpenAI(api_key=openai_api_key)

# --- Extract named references (labels only) ---
def extract_named_references(wb):
    named_refs = {}
    for name in wb.defined_names:
        defined_name = wb.defined_names[name]
        if defined_name.attr_text and not defined_name.is_external:
            for sheet_name, ref in defined_name.destinations:
                label = defined_name.name
                named_refs[label] = {
                    "sheet": sheet_name,
                    "ref": ref,
                    "formula": None,
                }
                try:
                    sheet = wb[sheet_name]
                    cell_ref = ref.split('!')[-1]
                    cell = sheet[cell_ref]
                    if cell.data_type == 'f':
                        named_refs[label]["formula"] = cell.value
                except Exception:
                    pass
    return named_refs

# --- Find dependencies based on formula references ---
def find_dependencies(named_refs):
    dependencies = {}
    all_labels = list(named_refs.keys())

    for target_label, info in named_refs.items():
        formula = (info.get("formula") or "")
        deps = []

        for source_label in all_labels:
            if source_label == target_label:
                continue

            if re.search(rf"\b{re.escape(source_label)}\b", formula):
                deps.append(source_label)

        dependencies[target_label] = deps

    return dependencies

# --- Create Graphviz graph (label → label) ---
def create_dependency_graph(dependencies):
    dot = graphviz.Digraph()
    for node in dependencies:
        dot.node(node)
    for node, deps in dependencies.items():
        for dep in deps:
            dot.edge(dep, node)
    return dot

# --- GPT explanations ---
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
    for label, info in named_refs.items():
        formula = info.get("formula", "")
        if not formula:
            doc = "No formula."
            py = ""
        else:
            doc = call_openai(f"Explain this Excel formula:\n{formula}")
            py = call_openai(f"Translate this Excel formula to Python:\n{formula}")
        results.append({
            "Named Reference": label,
            "AI Documentation": doc,
            "Excel Formula": formula,
            "Python Formula": py
        })
    return results

# --- Markdown table ---
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
st.title("📊 Excel Named Reference Dependency Viewer (Simplified, No File Prefix)")

uploaded_file = st.file_uploader("Upload an Excel (.xlsx) file", type=["xlsx"])

if uploaded_file:
    try:
        wb = load_workbook(io.BytesIO(uploaded_file.read()), data_only=False)
        named_refs = extract_named_references(wb)

        st.subheader("📌 Named References Extracted")
        st.json(named_refs)

        st.subheader("🔗 Dependency Graph")
        dependencies = find_dependencies(named_refs)
        dot = create_dependency_graph(dependencies)
        st.graphviz_chart(dot)

        st.subheader("🧠 AI Formula Explanations")
        with st.spinner("Calling GPT-4..."):
            rows = generate_ai_outputs(named_refs)
            st.markdown(render_markdown_table(rows), unsafe_allow_html=True)

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("⬆️ Upload a `.xlsx` file to begin.")
