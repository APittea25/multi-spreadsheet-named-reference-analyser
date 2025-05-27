import streamlit as st
from openai import OpenAI
from openpyxl import load_workbook
import graphviz
import io
import re
from collections import defaultdict

st.set_page_config(page_title="Named Reference Dependency Viewer", layout="wide")

# --- OpenAI setup ---
openai_api_key = st.secrets.get("OPENAI_API_KEY")
if not openai_api_key:
    st.error("❌ OPENAI_API_KEY not found in secrets.")
    st.stop()
client = OpenAI(api_key=openai_api_key)

# --- Formula cleaner ---
def simplify_formula(formula):
    if not formula:
        return ""
    formula = re.sub(r"'[^']+\.xlsx'!", "", formula)
    formula = re.sub(r"\[[^\]]+\][^!]*!", "", formula)
    return formula

# --- Extract named references with reliable formula handling ---
def extract_named_references(wb, file_label):
    named_refs = {}

    for name in wb.defined_names:
        dn = wb.defined_names[name]
        if dn.is_external or not dn.attr_text:
            continue
        for sheet_name, ref in dn.destinations:
            label = name
            try:
                sheet = wb[sheet_name]
                coord = ref.replace("$", "").split("!")[-1]
                top_left_cell = coord.split(":")[0]
                cell = sheet[top_left_cell]

                raw_formula = None

                if isinstance(cell.value, str) and cell.value.startswith("="):
                    raw_formula = cell.value.strip()
                elif hasattr(cell.value, "text"):
                    raw_formula = str(cell.value.text).strip()

                if raw_formula and raw_formula.startswith("="):
                    simplified = simplify_formula(raw_formula)
                    st.write(f"✅ `{label}` at `{sheet_name}!{top_left_cell}` = {raw_formula} → simplified: `{simplified}`")
                    formulas = [simplified]
                else:
                    st.write(f"⚠️ `{label}` at `{sheet_name}!{top_left_cell}` has no formula. Value = `{cell.value}`")
                    formulas = []

                named_refs[label] = {
                    "sheet": sheet_name,
                    "ref": ref,
                    "formulas": formulas,
                    "file": file_label
                }

            except Exception as e:
                st.write(f"❌ Error processing `{label}` → {e}")

    return named_refs

# --- Dependency detection ---
def find_dependencies(named_refs):
    dependencies = defaultdict(list)
    all_labels = list(named_refs.keys())

    for target_label, info in named_refs.items():
        formula_text = " ".join(info.get("formulas", []))
        for other_label in all_labels:
            if other_label == target_label:
                continue
            if re.search(rf"\b{re.escape(other_label)}\b", formula_text):
                dependencies[target_label].append(other_label)

    return dependencies

# --- Graphviz graph ---
def create_dependency_graph(dependencies, all_labels):
    dot = graphviz.Digraph()
    for label in all_labels:
        dot.node(label)
    for target, sources in dependencies.items():
        for source in sources:
            dot.edge(source, target)
    return dot

# --- GPT Explanation ---
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
        formulas = info.get("formulas", [])
        combined_formula = " + ".join(formulas)
        if not combined_formula:
            doc = "No formula."
            py = ""
        else:
            doc = call_openai(f"Explain this Excel formula:\n{combined_formula}")
            py = call_openai(f"Translate this Excel formula to Python:\n{combined_formula}")
        results.append({
            "Named Reference": label,
            "AI Documentation": doc,
            "Excel Formula": combined_formula,
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

# --- NEW SECTION: Recreate Spreadsheet Calculations in Python ---
def topological_sort(dependencies):
    from collections import defaultdict, deque

    in_degree = defaultdict(int)
    graph = defaultdict(list)

    for node in dependencies:
        for dep in dependencies[node]:
            graph[dep].append(node)
            in_degree[node] += 1

    queue = deque([n for n in dependencies if in_degree[n] == 0])
    sorted_order = []

    while queue:
        node = queue.popleft()
        sorted_order.append(node)
        for neighbor in graph[node]:
            in_degree[neighbor] -= 1
            if in_degree[neighbor] == 0:
                queue.append(neighbor)

    return sorted_order


def generate_end_to_end_python(named_refs, dependencies):
    order = topological_sort(dependencies)
    lines = ["# Full spreadsheet logic recreated in Python"]

    for name in order:
        formula = " + ".join(named_refs[name].get("formulas", []))
        if formula:
            lines.append(f"# {name}")
            lines.append(f"# Excel: {formula}")
        else:
            lines.append(f"# {name} is likely an input")
            lines.append(f"{name} = ...")
            lines.append("")
            continue

        context = "\n".join(
            f"{n}: {named_refs[n].get('formulas', [''])[0]}" for n in order if n != name
        )

        full_prompt = (
            f"You are recreating a spreadsheet in Python. Here are the named references and formulas:\n"
            f"{context}\n\n"
            f"Now translate the formula for '{name}':\n{formula}\n"
            f"Return only the Python expression for {name}."
        )

        python_translation = call_openai(full_prompt, max_tokens=200)
        lines.append(f"{name} = {python_translation}")
        lines.append("")

    return "\n".join(lines)

# --- Streamlit UI ---
st.title("📊 Excel Named Reference Dependency Viewer (Cloud-Safe)")

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
        dot = create_dependency_graph(dependencies, combined_named_refs.keys())
        st.graphviz_chart(dot)

        st.subheader("🧠 AI-Generated Formula Explanations")
        with st.spinner("Calling GPT-4..."):
            rows = generate_ai_outputs(combined_named_refs)
            st.markdown(render_markdown_table(rows), unsafe_allow_html=True)

        st.subheader("🧮 End-to-End Spreadsheet Logic in Python")
        with st.spinner("Generating full procedural script from GPT..."):
            final_script = generate_end_to_end_python(combined_named_refs, dependencies)
            st.code(final_script, language="python")

    else:
        st.warning("No named references found.")
else:
    st.info("⬆️ Upload one or more `.xlsx` files to begin.")
