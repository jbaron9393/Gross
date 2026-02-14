#!/usr/bin/env python3
"""
build_hne_viewer.py
- Reads "HNE Grossing Guide TEMPLATES.docx"
- Outputs:
  - hne_grossing_guide.json
  - hne_viewer.html (self-contained)
Usage:
  python build_hne_viewer.py
"""
from docx import Document
import json, re, os, html

DOCX_PATH = "HNE Grossing Guide TEMPLATES.docx"
OUT_JSON  = "hne_grossing_guide.json"
OUT_HTML  = "hne_viewer.html"

def heading_level(style_name: str):
    m = re.match(r"Heading (\d+)", style_name)
    return int(m.group(1)) if m else None

def build_tree(doc):
    root = {"title":"HNE Grossing Guide","level":0,"children":[], "content":[]}
    stack=[root]
    for p in doc.paragraphs:
        txt=(p.text or "").rstrip()
        if not txt.strip():
            if stack[-1]["content"] and stack[-1]["content"][-1] != "":
                stack[-1]["content"].append("")
            continue
        lvl=heading_level(p.style.name)
        if lvl:
            node={"title":txt.strip(), "level":lvl, "children":[], "content":[]}
            while stack and stack[-1]["level"] >= lvl:
                stack.pop()
            stack[-1]["children"].append(node)
            stack.append(node)
        else:
            stack[-1]["content"].append(txt)

    def prune(node):
        while node["content"] and node["content"][-1] == "":
            node["content"].pop()
        for ch in node["children"]:
            prune(ch)
    prune(root)

    counter = 0
    def assign_ids(node, path_titles):
        nonlocal counter
        node["id"] = f"n{counter:05d}"
        counter += 1
        node["path"] = path_titles+[node["title"]] if node["level"]>0 else [node["title"]]
        for ch in node.get("children", []):
            assign_ids(ch, node["path"])
    assign_ids(root, [])
    return {"generated_from": os.path.basename(DOCX_PATH), "root": root}

def escape_js_string(s: str):
    return s.replace("\\","\\\\").replace("</","<\\/")

def main():
    if not os.path.exists(DOCX_PATH):
        raise SystemExit(f"Can't find {DOCX_PATH} in this folder.")
    doc = Document(DOCX_PATH)
    data = build_tree(doc)

    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    # Easiest: keep your own viewer HTML as a template and just inject the JSON.
    # For now, we'll just fail if the viewer isn't present.
    if not os.path.exists(OUT_HTML):
        raise SystemExit(f"Can't find {OUT_HTML}. Copy it into this folder first (or generate once).")

    # Replace JSON inside <script id="hne-data" type="application/json">...</script>
    import re as _re
    with open(OUT_HTML, "r", encoding="utf-8") as f:
        html_txt = f.read()

    new_json = escape_js_string(json.dumps(data, ensure_ascii=False))
    html_txt = _re.sub(
        r'(<script id="hne-data" type="application/json">)(.*?)(</script>)',
        lambda m: m.group(1) + new_json + m.group(3),
        html_txt,
        flags=_re.DOTALL
    )

    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write(html_txt)

    print(f"Wrote {OUT_JSON} and updated {OUT_HTML}")

if __name__ == "__main__":
    main()
