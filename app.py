import streamlit as st
from openai import OpenAI
from openpyxl import load_workbook
import graphviz
import io
import re
from collections import defaultdict

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
                label = defined_name.name
                if label not in named_refs:
                    named_refs[label] = []
                item = {
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
                        item["formula"] = cell.value
                except Exception:
                    pass
                named_refs[label].append(item)
    return named_refs

# --- Build dependency map based on formula usage ---
def find_dependencies(named_refs):
    dependencies = defaultdict(list)
    labels = list(named_refs.keys())

    for target_label in labels:
        for instance in named_refs[target_label]:
            formula = (instance.get("formula") or "")
            for source_label in labels:
                if source_label == target_label:
                    continue
                if re.search(rf"\b{re.escape(source_label)}\b", formula):
                    dependencies[target_label].append(source_label)
    return dependencies

# --- Create Graphviz dependency graph with all labels ---
def create_dependency_graph(dependencies, all_labels):
    dot = graphviz.Digraph()
    for label in all_labels:
        dot.node(label)
    for target, sources in dependencies.items():
        for source in sources:
            dot.edge(source, target)
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
    for label, instances in named_refs.items():
        formula = instances[0].get("formula", "") if instances else ""
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

uploaded_files = st.file_uploader("Upload one or more Excel (.xlsx) files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    combined_named_refs = defaultdict(list)
    collision_tracker = defaultdict(set)

    for uploaded_file in uploaded_files:
        try:
            wb = load_workbook(io.BytesIO(uploaded_file.read()), data_only=False)
            refs = extract_named_references(wb, uploaded_file.name)
            for label, entries in refs.items():
                combined_named_refs[label].extend(entries)
                collision_tracker[label].add(uploaded_file.name)
        except Exception as e:
            st.error(f"❌ Error reading {uploaded_file.name}: {e}")

    if combined_named_refs:
        st.subheader("📌 Named References Extracted")
        st.json(combined_named_refs)

        collisions = {k: list(v) for k, v in collision_tracker.items() if len(v) > 1}
        if collisions:
            st.warning("⚠️ Naming collisions detected across files:")
            st.json(collisions)

        st.subheader("🔗 Dependency Graph")
        dependencies = find_dependencies(combined_named_refs)
        dot = create_dependency_graph(dependencies, combined_named_refs.keys())
        st.graphviz_chart(dot)

        st.subheader("🧠 AI Formula Explanations")
        with st.spinner("Calling GPT-4..."):
            rows = generate_ai_outputs(combined_named_refs)
            st.markdown(render_markdown_table(rows), unsafe_allow_html=True)
    else:
        st.warning("No named references found.")
else:
    st.info("⬆️ Upload one or more `.xlsx` files to begin.")
