# =============================================================================
# SciWrite AI  v2.0  —  Production-Grade Research Paper Generator (Flask/Vercel)
# =============================================================================
import os
# CRITICAL VERCEL MATPLOTLIB SETUP: Serverless environments are strictly read-only 
# except for the /tmp directory. MPLCONFIGDIR must be redirected before import.
os.environ["MPLCONFIGDIR"] = "/tmp"

from flask import Flask, render_template, request, jsonify
import re, textwrap, tempfile, subprocess, base64, json
from pathlib import Path
from typing import Optional
import pandas as pd
import jinja2
from google import genai
import io
import zipfile
import matplotlib
matplotlib.use('Agg')  # Force non-interactive backend to prevent GUI thread collision errors
import matplotlib.pyplot as plt
import urllib.request
import urllib.parse
import xml.etree.ElementTree as ET
import time

app = Flask(__name__)

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────
GEMINI_MODEL = "gemini-2.5-flash-lite"

SECTION_TAGS = [
    "ABSTRACT", "INTRODUCTION", "RELATED_WORK", "METHODOLOGY",
    "RESULTS_AND_ANALYSIS", "DISCUSSION", "CONCLUSION", "ACKNOWLEDGEMENTS",
]

WORDS_PER_PAGE_SINGLE = 450
WORDS_PER_PAGE_DOUBLE = 600

SECTION_WEIGHT: dict[str, float] = {
    "ABSTRACT":             0.05,
    "INTRODUCTION":         0.16,
    "RELATED_WORK":         0.14,
    "METHODOLOGY":          0.22,
    "RESULTS_AND_ANALYSIS": 0.20,
    "DISCUSSION":           0.12,
    "CONCLUSION":           0.07,
    "ACKNOWLEDGEMENTS":     0.04,
}

# ─────────────────────────────────────────────────────────────────────────────
# UTILITIES
# ─────────────────────────────────────────────────────────────────────────────

def escape_for_latex(text: str) -> str:
    if not isinstance(text, str):
        text = str(text)
    text = text.replace("\\", r"\textbackslash{}")
    text = text.replace("&",  r"\&")
    text = text.replace("%",  r"\%")
    text = text.replace("$",  r"\$")
    text = text.replace("#",  r"\#")
    text = text.replace("_",  r"\_")
    text = text.replace("{",  r"\{")
    text = text.replace("}",  r"\}")
    text = text.replace("~",  r"\textasciitilde{}")
    text = text.replace("^",  r"\textasciicircum{}")
    return text

def extract_xml_section(tag: str, raw: str) -> str:
    m = re.search(rf"<{tag}>(.*?)</{tag}>", raw, re.DOTALL | re.IGNORECASE)
    return m.group(1).strip() if m else ""

def compute_section_targets(pages: int, two_col: bool) -> dict[str, int]:
    wpp = WORDS_PER_PAGE_DOUBLE if two_col else WORDS_PER_PAGE_SINGLE
    total = pages * wpp
    return {tag: max(80, int(total * w)) for tag, w in SECTION_WEIGHT.items()}

def fetch_arxiv_literature(query: str, max_results: int = 8) -> tuple[str, str]:
    if not query.strip():
        query = "machine learning" 
        
    safe_query = urllib.parse.quote(query)
    url = f"https://export.arxiv.org/api/query?search_query=all:{safe_query}&start=0&max_results={max_results}&sortBy=relevance"
    
    bibtex_entries = []
    abstract_summaries = []
    
    try:
        with urllib.request.urlopen(url, timeout=10) as response:
            xml_data = response.read()
            
        root = ET.fromstring(xml_data)
        ns = {'atom': 'http://www.w3.org/2005/Atom'}
        
        for i, entry in enumerate(root.findall('atom:entry', ns)):
            title = entry.find('atom:title', ns).text.replace('\n', ' ').strip()
            summary = entry.find('atom:summary', ns).text.replace('\n', ' ').strip()
            published = entry.find('atom:published', ns).text[:4]
            authors = [author.find('atom:name', ns).text for author in entry.findall('atom:author', ns)]
            author_str = " and ".join(authors)
            
            last_name = authors[0].split()[-1].lower() if authors else "unknown"
            first_word = re.sub(r'[^a-zA-Z0-9]', '', title.split()[0].lower())
            cite_key = f"{last_name}{published}{first_word}"
            
            bibtex = textwrap.dedent(f"""
            @article{{{cite_key},
              title={{{title}}},
              author={{{author_str}}},
              journal={{arXiv preprint}},
              year={{{published}}}
            }}
            """).strip()
            
            bibtex_entries.append(bibtex)
            abstract_summaries.append(f"[{cite_key}] {title} ({published}): {summary}")
            
        return "\n\n".join(bibtex_entries), "\n\n".join(abstract_summaries)
        
    except Exception as e:
        print(f"arXiv API fetch failed: {e}. Proceeding with minimal context.")
        return "", ""

# ─────────────────────────────────────────────────────────────────────────────
# GEMINI
# ─────────────────────────────────────────────────────────────────────────────

def call_gemini_with_retry(api_key: str, prompt: str, retries: int = 3) -> str:
    client = genai.Client(api_key=api_key)
    for attempt in range(retries):
        try:
            response = client.models.generate_content(
                model=GEMINI_MODEL, 
                contents=prompt,
                config=genai.types.GenerateContentConfig(temperature=0.2)
            )
            return response.text
        except Exception as e:
            if "503" in str(e) or "429" in str(e):
                if attempt < retries - 1:
                    time.sleep(2 ** attempt)
                    continue
            raise e
    return ""

def clean_llm_latex_output(raw_output: str) -> str:
    clean_text = raw_output.strip()
    if clean_text.startswith("```"):
        clean_text = re.sub(r'^```[a-zA-Z]*\n', '', clean_text)
        clean_text = re.sub(r'\n```$', '', clean_text)
    return clean_text.strip()

def generate_section_prompt(
    tag: str, title: str, arxiv_context: str, bibtex: str, 
    user_notes: str, previous_sections: dict, target_words: int
) -> str:
    history = ""
    if previous_sections:
        history = "PREVIOUSLY GENERATED SECTIONS:\n"
        for k, v in previous_sections.items():
            history += f"--- {k} ---\n{v[:500]}... [truncated]\n\n"
            
    cite_keys = re.findall(r"@\w+\{(\w+),", bibtex)
    cite_str = ", ".join(cite_keys) if cite_keys else "(none provided)"

    return textwrap.dedent(f"""
    You are an elite academic researcher. You are writing ONE specific section of a research paper.
    
    PAPER TITLE: {title}
    CURRENT SECTION TO WRITE: {tag}
    TARGET WORD COUNT: ~{target_words} words.
    
    REAL LITERATURE CONTEXT (Use these to ground your claims):
    {arxiv_context}
    
    AVAILABLE CITATION KEYS: {cite_str}
    
    USER NOTES FOR THIS PAPER:
    {user_notes}
    
    {history}
    
    INSTRUCTIONS:
    1. Write ONLY the '{tag}' section. Do not write any other sections.
    2. Write in formal, third-person academic LaTeX prose. 
    3. Use \\cite{{key}} frequently and accurately based on the Literature Context provided.
    4. Do not output standard markdown code fences (like ```latex). Output raw text.
    5. Do not output the section header (e.g., no \\section{{{tag}}}), just the body paragraphs.
    6. Ensure the narrative flows logically from the previously generated sections.
    """).strip()

def verify_and_enrich_bibliography(api_key: str, paper_title: str, keywords: str, user_bibtex: str) -> str:
    if user_bibtex.strip():
        return user_bibtex.strip()
        
    client = genai.Client(api_key=api_key)
    
    grounding_prompt = f"""
    Find 4 genuine, highly-cited academic research papers relevant to this domain to construct a BibTeX library.
    
    Domain Context:
    Title: {paper_title}
    Keywords: {keywords}
    
    CRITICAL GROUNDING REQUIREMENTS:
    1. Do NOT invent or guess references. Every paper must be a real, verifiable publication indexed in Google Scholar, IEEE, or arXiv.
    2. Format the response strictly as a single text block containing valid BibTeX records.
    3. Ensure every record includes author, title, journal/booktitle, volume, year, and a clean alphanumeric citation key.
    4. Output ONLY valid BibTeX syntax. No markdown wrappers, no explanations.
    """
    
    try:
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=grounding_prompt,
            config=genai.types.GenerateContentConfig(
                temperature=0.0,
                tools=[{"google_search": {}}],
                system_instruction="You are a strict citation validation agent. You output true, verifiable academic BibTeX data blocks only."
            )
        )
        return response.text.strip().replace("```bibtex", "").replace("```", "")
    except Exception:
        return textwrap.dedent("""
        @article{vaswani2017attention,
          title={Attention is all you need},
          author={Vaswani, Ashish and Shazeer, Noam and Parmar, Niki and Uszkoreit, Jakob and Jones, Llion and Gomez, Aidan N and Kaiser, {\L}ukasz and Polosukhin, Illia},
          journal={Advances in neural information processing systems},
          volume={30},
          year={2017}
        }
        """).strip()

def generate_academic_chart(api_key: str, paper_title: str, results_notes: str, figure_index: int) -> tuple[str, bytes] | None:
    client = genai.Client(api_key=api_key)
    chart_filename = f"/tmp/generated_chart_{figure_index}.png"
    
    prompt = f"""
    You are an expert Data Scientist and Academic typesetter. Write an isolated Python script using matplotlib to generate a publication-quality chart for a research paper.
    
    PAPER CONTEXT:
    Title: {paper_title}
    Experimental Results Data: {results_notes}
    Target Figure Filename: {chart_filename}
    
    CRITICAL DESIGN REQUIREMENTS:
    1. Aesthetics: Use a clean academic style. Set dark grey or black gridlines, clear legible labels, axis titles, and a descriptive legend if applicable.
    2. Data: Translate the user's unstructured metrics into explicit arrays or dataframes within the script.
    3. Output: The script MUST save the file to the exact path string: '{chart_filename}' using plt.savefig('{chart_filename}', dpi=300, bbox_inches='tight').
    4. Safety: Do NOT use plt.show(). Use plt.close('all') at the absolute end of the execution string.
    5. Formatting: Output ONLY raw executable Python code. No markdown formatting. No ```python blocks. No explanations.
    """
    
    try:
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt,
            config=genai.types.GenerateContentConfig(
                temperature=0.1,
                system_instruction="You are an automated Python script generator that outputs raw executable code strings with zero markdown encapsulation."
            )
        )
        
        clean_code = response.text.strip().replace("```python", "").replace("```", "")
        
        local_scope = {"plt": plt, "pd": pd, "re": re}
        plt.close('all')
        
        exec(clean_code, globals(), local_scope)
        
        if os.path.exists(chart_filename):
            with open(chart_filename, "rb") as f:
                img_bytes = f.read()
            os.remove(chart_filename)
            plt.close('all')
            return f"generated_chart_{figure_index}.png", img_bytes
            
    except Exception as e:
        print(f"Chart Engine Failure on Asset {figure_index}: {str(e)}")
        plt.close('all')
        return None

def generate_overleaf_zip(
    latex_src: str, 
    bibtex_data: str, 
    uploaded_figs: list, 
    generated_figs: list[tuple[str, bytes]]
) -> bytes:
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr("main.tex", latex_src)
        if bibtex_data.strip():
            zip_file.writestr("references.bib", bibtex_data.strip())
            
        if uploaded_figs:
            for uf in uploaded_figs:
                filename = getattr(uf, "filename", getattr(uf, "name", "figure.png"))
                safe_name = re.sub(r"[^\w.\-]", "_", filename)
                uf.seek(0)
                zip_file.writestr(safe_name, uf.read())
                
        if generated_figs:
            for filename, img_bytes in generated_figs:
                zip_file.writestr(filename, img_bytes)
                
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

# ─────────────────────────────────────────────────────────────────────────────
# JINJA2 + LATEX TEMPLATE
# ─────────────────────────────────────────────────────────────────────────────

def make_jinja_env() -> jinja2.Environment:
    return jinja2.Environment(
        block_start_string="[%", block_end_string="%]",
        variable_start_string="[[", variable_end_string="]]",
        comment_start_string="[#", comment_end_string="#]",
        trim_blocks=True, lstrip_blocks=True,
        autoescape=False, undefined=jinja2.Undefined,
    )

LATEX_TMPL = r"""
\documentclass[11pt[% if two_col %],twocolumn[% endif %]]{article}
\usepackage[T1]{fontenc}
\usepackage[utf8]{inputenc}
\usepackage{lmodern}
\usepackage{microtype}
\usepackage{amsmath,amssymb,amsfonts,mathtools,bm}
\usepackage[
[% if two_col %]
    letterpaper,top=0.9in,bottom=0.9in,left=0.65in,right=0.65in,columnsep=18pt,
[% else %]
    letterpaper,top=1in,bottom=1in,left=1.1in,right=1.1in,
[% endif %]
]{geometry}
\usepackage{graphicx,booktabs,tabularx,multirow,array,float}
\usepackage[numbers,sort&compress]{natbib}
\usepackage{hyperref,url}
\hypersetup{colorlinks=true,linkcolor=blue!65!black,citecolor=green!55!black,urlcolor=blue!75!black,
    pdftitle={[[ escaped_title ]]},pdfauthor={[[ first_author ]]},pdfkeywords={[[ escaped_kw ]]}}
\usepackage{parskip,setspace,titlesec,abstract,authblk,fancyhdr,xcolor,soul}
\usepackage{algorithm,algpseudocode}
\titleformat{\section}{\large\bfseries\scshape}{}{0em}{}[\vspace{-2pt}\rule{\linewidth}{0.5pt}\vspace{2pt}]
\titleformat{\subsection}{\normalsize\bfseries}{}{0em}{}
\titlespacing*{\section}{0pt}{14pt}{6pt}
\titlespacing*{\subsection}{0pt}{10pt}{4pt}
\pagestyle{fancy}\fancyhf{}
\renewcommand{\headrulewidth}{0.35pt}
\fancyhead[L]{\small\scshape\textcolor{gray}{[[ short_title ]]}}
\fancyhead[R]{\small\textcolor{gray}{\thepage}}
\fancyfoot[C]{\small\textcolor{gray!60}{Generated by SciWrite AI \textperiodcentered{} \today}}
\renewcommand{\abstractname}{\normalsize\bfseries\scshape Abstract}
\setlength{\absleftindent}{0pt}\setlength{\absrightindent}{0pt}
\title{\vspace{-1.5em}\Large\bfseries [[ escaped_title ]]%
[% if kw_raw %]\\\vspace{0.3em}\normalsize\normalfont\textit{Keywords:}\ \small [[ escaped_kw ]][% endif %]}
[% for a in authors %]
\author[[[ loop.index ]]]{\textbf{[[ a.name ]]}[% if a.email %]\thanks{\href{mailto:[[ a.email ]]}{[[ a.email ]]}}\vspace{-0.5em}[% endif %]}
\affil[[[ loop.index ]]]{\small\textit{[[ a.affil ]]}}
[% endfor %]
\date{\today}
\begin{document}
\maketitle\thispagestyle{fancy}
\begin{abstract}\noindent [[ sec.ABSTRACT ]]\end{abstract}
[% if two_col %]\vspace{0.5em}\noindent\rule{\linewidth}{0.3pt}\vspace{0.2em}[% endif %]
\section{Introduction}
[[ sec.INTRODUCTION ]]
\section{Related Work}
[[ sec.RELATED_WORK ]]
\section{Methodology}
[[ sec.METHODOLOGY ]]
[% if figs %]
[% for f in figs %]
\begin{figure}[ht]\centering
\includegraphics[width=0.95\linewidth]{[[ f.fn ]]}
\caption{[[ f.cap ]]}\label{fig:[[ loop.index ]]}
\end{figure}
[% endfor %]
[% endif %]
\section{Results and Analysis}
[[ sec.RESULTS_AND_ANALYSIS ]]
\section{Discussion}
[[ sec.DISCUSSION ]]
\section{Conclusion}
[[ sec.CONCLUSION ]]
\section*{Acknowledgements}
[[ sec.ACKNOWLEDGEMENTS ]]
[% if bibtex %]
\begin{filecontents*}{\jobname.bib}
[[ bibtex ]]
\end{filecontents*}
\bibliographystyle{unsrtnat}
\bibliography{\jobname}
[% endif %]
\end{document}
"""

def render_latex(
    title: str, authors_df: pd.DataFrame, two_col: bool,
    keywords: str, sections: dict, bibtex: str,
    figs: list | None = None,
) -> str:
    env  = make_jinja_env()
    tmpl = env.from_string(LATEX_TMPL)
    author_list = []
    first_author = ""
    for _, row in authors_df.iterrows():
        name = str(row.get("Name","")).strip()
        if not name:
            continue
        if not first_author:
            first_author = escape_for_latex(name)
        author_list.append({
            "name":  escape_for_latex(name),
            "affil": escape_for_latex(str(row.get("Affiliation",""))),
            "email": escape_for_latex(str(row.get("Email",""))),
        })
    short = (title[:52]+"...") if len(title) > 55 else title
    return tmpl.render(
        escaped_title=escape_for_latex(title),
        short_title=escape_for_latex(short),
        first_author=first_author,
        escaped_kw=escape_for_latex(keywords),
        kw_raw=keywords.strip(),
        authors=author_list,
        two_col=two_col,
        sec=sections,
        bibtex=bibtex.strip(),
        figs=figs or [],
    )

def compile_pdf(
    latex_src: str,
    fig_bins: list[tuple[str, bytes]] | None = None,
) -> tuple[bytes | None, str | None]:
    # NOTE: pdflatex requires TeX Live, which is typically not available in standard Vercel serverless functions.
    # This block is maintained for perfect 1-to-1 functionality preservation.
    with tempfile.TemporaryDirectory() as tmp:
        p = Path(tmp)
        tex = p / "paper.tex"
        tex.write_text(latex_src, encoding="utf-8")
        for fname, fdata in (fig_bins or []):
            (p / fname).write_bytes(fdata)
        cmd = ["pdflatex","-interaction=nonstopmode","-halt-on-error",
               "-output-directory", tmp, str(tex)]
        try:
            for _ in range(2):
                r = subprocess.run(cmd, capture_output=True, text=True,
                                   cwd=tmp, timeout=120)
        except FileNotFoundError:
            return None, "pdflatex_not_found"
        except subprocess.TimeoutExpired:
            return None, "pdflatex timed out (>120s)."
        
        pdf = p / "paper.pdf"
        if pdf.exists() and pdf.stat().st_size > 0:
            return pdf.read_bytes(), None
        log = p / "paper.log"
        if log.exists():
            txt = log.read_text(encoding="utf-8", errors="replace")
            errs = [l for l in txt.splitlines() if l.startswith("!")]
            excerpt = "\n".join(errs[:30]) or txt[-2000:]
        else:
            excerpt = r.stderr[-2000:]
        return None, excerpt

# ─────────────────────────────────────────────────────────────────────────────
# FLASK ROUTES
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/generate", methods=["POST"])
def generate_paper():
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        return jsonify({"error": "Server configuration error: GEMINI_API_KEY environment variable not set."}), 500

    try:
        # Form Data Extraction
        paper_title = request.form.get("title", "")
        paper_keywords = request.form.get("keywords", "")
        authors_json = request.form.get("authors", "[]")
        authors_df = pd.DataFrame(json.loads(authors_json))
        
        abstract_goals = request.form.get("abstract_goals", "")
        methodology_notes = request.form.get("methodology_notes", "")
        intro_bg = request.form.get("intro_bg", "")
        results_data = request.form.get("results_data", "")
        extra_notes = request.form.get("extra_notes", "")
        bibtex_entries = request.form.get("bibtex_entries", "")
        
        layout_opt = request.form.get("layout_opt", "Single Column")
        two_col = layout_opt == "Double Column"
        length_mode = request.form.get("length_mode", "Pages")
        
        if length_mode == "Pages":
            target_pages = int(request.form.get("target_pages", 8))
        else:
            est_words = int(request.form.get("target_words", 5000))
            wpp = WORDS_PER_PAGE_DOUBLE if two_col else WORDS_PER_PAGE_SINGLE
            target_pages = max(4, est_words // wpp)

        uploaded_figs = request.files.getlist("figures")
        if len(uploaded_figs) == 1 and uploaded_figs[0].filename == '':
            uploaded_figs = []

        sec_targets = compute_section_targets(target_pages, two_col)

        # Step 1: ArXiv Literature Grounding
        search_query = f"{paper_title} {paper_keywords}"
        real_bibtex, arxiv_summaries = fetch_arxiv_literature(search_query, max_results=6)
        combined_bibtex = f"{bibtex_entries}\n\n{real_bibtex}".strip()

        # Step 2: Matplotlib Rendering
        generated_charts_list = []
        for idx in range(1, 3):
            chart_asset = generate_academic_chart(api_key, paper_title, results_data, idx)
            if chart_asset:
                generated_charts_list.append(chart_asset)

        # Step 3: Sequential Academic Synthesis
        sections_dict = {}
        compiled_user_notes = f"Abstract: {abstract_goals}\nIntro: {intro_bg}\nMethod: {methodology_notes}\nResults: {results_data}\nExtra: {extra_notes}"
        
        for tag in SECTION_TAGS:
            prompt = generate_section_prompt(
                tag=tag, 
                title=paper_title, 
                arxiv_context=arxiv_summaries, 
                bibtex=combined_bibtex, 
                user_notes=compiled_user_notes, 
                previous_sections=sections_dict, 
                target_words=sec_targets.get(tag, 300)
            )
            section_text = call_gemini_with_retry(api_key, prompt)
            sections_dict[tag] = clean_llm_latex_output(section_text)
            
        wcs = {t: len(v.split()) for t, v in sections_dict.items()}
        total_words = sum(wcs.values())

        # Step 4: Template Render & ZIP Export
        fig_list = []
        for uf in uploaded_figs:
            safe_name = re.sub(r"[^\w.\-]", "_", uf.filename)
            fig_list.append({"fn": safe_name, "cap": f"Empirical field observations: {uf.filename}"})
        for filename, _ in generated_charts_list:
            fig_list.append({"fn": filename, "cap": "Programmatic verification matrix detailing analytical experimental parameters."})
            
        latex_src = render_latex(
            title=paper_title, authors_df=authors_df, two_col=two_col,
            keywords=paper_keywords, sections=sections_dict,
            bibtex=combined_bibtex, figs=fig_list
        )
        
        zip_payload = generate_overleaf_zip(latex_src, combined_bibtex, uploaded_figs, generated_charts_list)
        
        # State encoding for serverless stateless download (JSON Response)
        response_data = {
            "latex_src": latex_src,
            "zip_base64": base64.b64encode(zip_payload).decode("utf-8"),
            "sections": sections_dict,
            "word_counts": wcs,
            "total_words": total_words,
            "targets": sec_targets
        }

        return jsonify(response_data)

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True, port=5000)