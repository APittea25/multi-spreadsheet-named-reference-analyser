import streamlit as st
from openai import OpenAI
from openpyxl import load_workbook
import graphviz
import io

st.set_page_config(page_title="Excel Named Range Visualizer", layout="wide")

# Initialize OpenAI
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# --- Extract named references ---
@st.cache_data(show_spinner=False)
def extract_named_references(_wb):
    named_refs = {}
    for name in _wb.defined_names:
        defined_name = _wb.defined_names[name]
        if defined_name.attr_text and not defined_name.is_external:
            dests = list(defined_name.destinations)
            for sheet_name, ref in dests:
                named_refs[defined_name.name] = {
                    "sheet": sheet_name,
                    "ref": ref,
                    "formula": None
                }
                try:
                    sheet = _wb[sheet_name]
                    cell_ref = ref.split('!')[-1]
                    cell = sheet[cell_ref]
                    if cell.data_type == 'f':
                        named_refs[defined_name.name]["formula"] = cell.value
                except Exception:
                    pass
    return named_refs

# --- Find dependencies among named references ---
@st.cache_data(show_spinner=False)
def find_dependencies(named_refs):
    dependencies = {}
    for name, info in named_refs.items():
        formula = info.get("formula", "")
        if formula:
            formula = formula.upper()
            deps = [other for other in named_refs if other != name and other.upper() in formula]
            dependencies[name] = deps
        else:
            dependencies[name] = []
    return dependencies

# --- Create graph ---
def create_dependency_graph(dependencies):
    dot = graphviz.Digraph()
    for ref in dependencies:
        dot.node(ref)
    for ref, deps in dependencies.items():
        for dep in deps:
            dot.edge(dep, ref)
    return dot

# --- OpenAI explanations ---
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

def render_markdown_table(rows):
    headers = ["Named Reference", "AI Documentation", "Excel Formula", "Python Formula"]
    md = "| " + " | ".join(headers) + " |\n"
    md += "| " + " | ".join(["---"] * len(headers)) + " |\n"
    for row in rows:
        md += "| " + " | ".join([
            row["Named Reference"],
            row["AI Documentation"].replace("\n", " "),
            row["Excel Formula"].replace("\n", " "),
            row["Python Formula"].replace("\n", " ")
        ]) + " |\n"
    return md

# --- UI ---
st.title("📊 Excel Named Reference Dependency Viewer")

uploaded_file = st.file_uploader("Upload an Excel (.xlsx) file", type=["xlsx"])

if uploaded_file:
    try:
        wb = load_workbook(io.BytesIO(uploaded_file.read()), data_only=False)

        st.subheader("📌 Named References")
        named_refs = extract_named_references(wb)
        st.json(named_refs)

        st.subheader("🔗 Dependency Graph")
        dependencies = find_dependencies(named_refs)
        dot = create_dependency_graph(dependencies)
        st.graphviz_chart(dot)

        st.subheader("🧠 AI Formula Explanations")
        with st.spinner("Asking GPT..."):
            rows = generate_ai_outputs(named_refs)
            st.markdown(render_markdown_table(rows), unsafe_allow_html=True)

    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
else:
    st.info("⬆️ Upload a `.xlsx` file to begin.")
