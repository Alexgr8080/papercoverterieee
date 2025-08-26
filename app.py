#!/usr/bin/env python3
import os, re, uuid, shutil, subprocess, json
from pathlib import Path
from flask import Flask, render_template, request, redirect, url_for, send_file, flash
from werkzeug.utils import secure_filename

try:
    import docx  # python-docx
except Exception:
    docx = None

APP_ROOT = Path(__file__).resolve().parent
UPLOAD_DIR = APP_ROOT / "uploads"
OUTPUT_DIR = APP_ROOT / "outputs"
TEMPLATE_PATH = APP_ROOT / "ieee_conf_template.tex"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.secret_key = "dev-" + str(uuid.uuid4())

def which(cmd):
    from shutil import which as _which
    return _which(cmd)

def find_pandoc():
    """Return full path to pandoc executable if found, else None.
    Checks env vars, PATH, and common install folders on Windows/macOS/Linux.
    """
    import os, sys
    from pathlib import Path
    # 1) explicit env vars
    for key in ("PANDOC", "PANDOC_PATH"):
        p = os.environ.get(key)
        if p and Path(p).exists():
            return str(Path(p))
    # 2) PATH
    w = which("pandoc")
    if w:
        return w
    # 3) common locations
    candidates = []
    if os.name == "nt":
        candidates += [
            r"C:\Program Files\Pandoc\pandoc.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Pandoc\pandoc.exe"),
            os.path.expanduser(r"~\AppData\Local\Pandoc\pandoc.exe"),
        ]
    else:
        candidates += [
            "/usr/local/bin/pandoc",
            "/usr/bin/pandoc",
            "/opt/homebrew/bin/pandoc",  # Apple Silicon Homebrew
        ]
    for c in candidates:
        if c and Path(c).exists():
            return c
    return None

    from shutil import which as _which
    return _which(cmd)

def run_cmd(cmd, cwd=None):
    p = subprocess.run(cmd, cwd=cwd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    return p.returncode, p.stdout, p.stderr

def docx_to_markdown(docx_path: Path, media_dir: Path) -> str:
    if not which("pandoc"):
        raise RuntimeError("Pandoc is required to convert DOCX. Please install Pandoc.")
    media_dir.mkdir(parents=True, exist_ok=True)
    cmd = ["pandoc", str(docx_path), "-t", "gfm", "--extract-media", str(media_dir)]
    rc, out, err = run_cmd(cmd)
    if rc != 0:
        raise RuntimeError(f"Pandoc failed: {err}")
    return out

def naive_md_to_latex(text: str) -> str:
    out = []
    for line in text.splitlines():
        if line.startswith("# "):
            out.append(r"\section{" + line[2:] + "}")
        elif line.startswith("## "):
            out.append(r"\subsection{" + line[3:] + "}")
        elif line.startswith("### "):
            out.append(r"\subsubsection{" + line[4:] + "}")
        else:
            out.append(line)
    return "\n".join(out)

def markdown_to_latex(md_text: str) -> str:
    pandoc = find_pandoc()
    if not pandoc:
        return naive_md_to_latex(md_text)
    p = subprocess.run([pandoc, "-f", "gfm", "-t", "latex"], input=md_text, text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if p.returncode != 0:
        return naive_md_to_latex(md_text)
    return p.stdout

def parse_sections_from_markdown(md_text: str):
    title = ""
    abstract = ""
    keywords = []
    authors_raw = ""

    lines = md_text.splitlines()
    for ln in lines:
        if ln.startswith("# "):
            title = ln[2:].strip()
            break

    abs_re = re.compile(r"^#{1,6}\s*abstract\s*$", re.IGNORECASE)
    key_re = re.compile(r"^(keywords|index terms)\s*[:—-]\s*(.+)$", re.IGNORECASE)

    in_abs = False
    abs_buf = []
    for ln in lines:
        if abs_re.match(ln.strip()):
            in_abs = True
            continue
        if in_abs and ln.startswith("#"):
            in_abs = False
        if in_abs:
            abs_buf.append(ln.strip())
        m = key_re.match(ln.strip())
        if m and not keywords:
            import re as _re
            ks = _re.split(r"[;,]\s*|,?\s+and\s+", m.group(2))
            keywords = [k.strip(" .") for k in ks if k.strip()]

    abstract = " ".join([s for s in abs_buf]).strip()

    possible = []
    for ln in lines[:20]:
        if ln.startswith("#") or abs_re.match(ln) or key_re.match(ln):
            continue
        if re.search(r"@|University|Institute|College|Department|Dept\.|School|Research|Lab", ln, re.I):
            possible.append(ln.strip())
    authors_raw = "\n".join(possible[:4])

    return {"title": title, "abstract": abstract, "keywords": keywords, "authors_raw": authors_raw}

def split_authors_blocks(authors_raw: str):
    authors = []
    chunks = re.split(r"\n\s*\n|;\s*", authors_raw.strip())
    for ch in chunks:
        if not ch.strip():
            continue
        email = None
        m = re.search(r"([\\w\\.-]+@[\\w\\.-]+\\.\\w+)", ch)
        if m:
            email = m.group(1)
        name = ch.split(",")[0].strip()
        affil = ch.strip()
        authors.append({"name": name, "affiliation": affil, "organization": "", "city_country": "", "email_or_orcid": email or ""})
    if not authors:
        authors = [{"name": "", "affiliation": "", "organization": "", "city_country": "", "email_or_orcid": ""}]
    return authors

from jinja2 import Template
def render_ieee_latex(context: dict) -> str:
    t = Template(TEMPLATE_PATH.read_text(encoding="utf-8"))
    return t.render(**context)

def strip_abstract_keywords_from_latex(latex_body: str) -> str:
    latex_body = re.sub(r"\\section\\*?\\{[Aa]bstract\\}.*?(?=\\section|\\subsection|\\subsubsection|\\paragraph|\\end\\{document\\}|$)", "", latex_body, flags=re.S)
    latex_body = re.sub(r"\\section\\*?\\{[Kk]eywords?.*?\\}.*?(?=\\section|\\subsection|\\subsubsection|\\paragraph|\\end\\{document\\}|$)", "", latex_body, flags=re.S)
    return latex_body

@app.get("/")
def index():
    return render_template("index.html")

@app.get("/diag")
def diag():
    pandoc = find_pandoc()
    return {"pandoc_path": pandoc, "PATH": os.environ.get("PATH","")[:500] + ("..." if len(os.environ.get("PATH",""))>500 else "")}

@app.post("/upload")
def upload():
    f = request.files.get("docx")
    if not f or not f.filename.lower().endswith(".docx"):
        flash("Please upload a .docx file.", "error")
        return redirect(url_for("index"))
    uid = uuid.uuid4().hex[:8]
    up_dir = UPLOAD_DIR / uid
    out_dir = OUTPUT_DIR / uid
    media_dir = out_dir / "media"
    up_dir.mkdir(parents=True, exist_ok=True)
    out_dir.mkdir(parents=True, exist_ok=True)

    filename = secure_filename(f.filename)
    docx_path = up_dir / filename
    f.save(docx_path)

    try:
        md_text = docx_to_markdown(docx_path, media_dir=media_dir)
    except Exception as e:
        flash(str(e), "error")
        shutil.rmtree(up_dir, ignore_errors=True)
        shutil.rmtree(out_dir, ignore_errors=True)
        return redirect(url_for("index"))

    meta_guess = parse_sections_from_markdown(md_text)
    authors = split_authors_blocks(meta_guess.get("authors_raw",""))
    latex_body = markdown_to_latex(md_text)
    latex_body = strip_abstract_keywords_from_latex(latex_body)

    (out_dir / "body.md").write_text(md_text, encoding="utf-8")
    (out_dir / "session.json").write_text(json.dumps({"uid": uid, "docx_name": filename, "body_md": "body.md", "media_dir": "media"}), encoding="utf-8")

    return render_template("review.html",
        uid=uid,
        guessed_title=meta_guess.get("title",""),
        guessed_abstract=meta_guess.get("abstract",""),
        guessed_keywords="; ".join(meta_guess.get("keywords", [])),
        authors=authors,
        body_preview=latex_body[:2000]
    )

@app.post("/generate/<uid>")
def generate(uid):
    title = request.form.get("title","").strip()
    abstract = request.form.get("abstract","").strip()
    import re as _re
    keywords = [k.strip() for k in _re.split(r"[;,]", request.form.get("keywords","")) if k.strip()]

    authors = []
    rows = int(request.form.get("author_rows","1"))
    for i in range(rows):
        prefix = f"author_{i}_"
        authors.append({
            "name": request.form.get(prefix+"name","").strip(),
            "affiliation": request.form.get(prefix+"affiliation","").strip(),
            "organization": request.form.get(prefix+"organization","").strip(),
            "city_country": request.form.get(prefix+"city_country","").strip(),
            "email_or_orcid": request.form.get(prefix+"email","").strip(),
        })

    out_dir = OUTPUT_DIR / uid
    if not out_dir.exists():
        flash("Session expired. Please re-upload your file.", "error")
        return redirect(url_for("index"))
    cache = json.loads((out_dir / "session.json").read_text(encoding="utf-8"))
    body_md = (out_dir / cache["body_md"]).read_text(encoding="utf-8")

    body_latex = markdown_to_latex(body_md)
    body_latex = strip_abstract_keywords_from_latex(body_latex)

    context = {"title": title or "Untitled", "author_blocks": authors, "abstract": abstract, "keywords": keywords, "body_latex": body_latex, "bibfile_base": "refs"}
    paper_tex = render_ieee_latex(context)

    (out_dir / "paper.tex").write_text(paper_tex, encoding="utf-8")
    (out_dir / "refs.bib").write_text("% Add your BibTeX entries here\n", encoding="utf-8")

    media_src = out_dir / cache.get("media_dir","media")
    if media_src.exists():
        figs = out_dir / "figures"
        figs.mkdir(exist_ok=True)
        for p in media_src.rglob("*"):
            if p.suffix.lower() in [".png",".jpg",".jpeg",".pdf",".eps"]:
                import shutil as _sh
                _sh.copy2(p, figs / p.name)

    zip_path = out_dir / "ieee_output.zip"
    import shutil as _sh
    _sh.make_archive(str(zip_path.with_suffix("")), "zip", out_dir)
    return redirect(url_for("download", uid=uid))

@app.get("/download/<uid>")
def download(uid):
    out_dir = OUTPUT_DIR / uid
    if not out_dir.exists():
        flash("Session expired. Please re-upload your file.", "error")
        return redirect(url_for("index"))
    zip_path = out_dir / "ieee_output.zip"
    if not zip_path.exists():
        import shutil as _sh
        _sh.make_archive(str(zip_path.with_suffix("")), "zip", out_dir)
    return send_file(zip_path, as_attachment=True, download_name="ieee_output.zip")

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
