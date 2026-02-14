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
import json
import os
import re
import zipfile
import xml.etree.ElementTree as ET

try:
    from docx import Document as _DocxDocument  # type: ignore
except ModuleNotFoundError:
    _DocxDocument = None

DOCX_PATH = "HNE Grossing Guide TEMPLATES.docx"
OUT_JSON = "hne_grossing_guide.json"
OUT_HTML = "hne_viewer.html"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def heading_level(style_name: str):
    """Extract heading level from names like 'Heading 2' or 'Heading2'."""
    m = re.match(r"Heading\s*(\d+)$", (style_name or "").strip(), flags=re.IGNORECASE)
    return int(m.group(1)) if m else None


def iter_paragraphs_with_styles(docx_path: str):
    """Yield (paragraph_text, style_name) from a .docx file.

    Uses python-docx when installed; otherwise falls back to parsing OOXML directly.
    """
    if _DocxDocument is not None:
        doc = _DocxDocument(docx_path)
        for p in doc.paragraphs:
            yield (p.text or ""), getattr(p.style, "name", "") or ""
        return

    ns = {"w": W_NS}
    with zipfile.ZipFile(docx_path) as zf:
        styles_by_id = {}
        try:
            styles_xml = zf.read("word/styles.xml")
            styles_root = ET.fromstring(styles_xml)
            for style in styles_root.findall("w:style", ns):
                sid = style.attrib.get(f"{{{W_NS}}}styleId", "")
                name_el = style.find("w:name", ns)
                name = name_el.attrib.get(f"{{{W_NS}}}val", sid) if name_el is not None else sid
                if sid:
                    styles_by_id[sid] = name
        except KeyError:
            pass

        doc_xml = zf.read("word/document.xml")
        root = ET.fromstring(doc_xml)

        for para in root.findall(".//w:p", ns):
            text_chunks = []
            for t in para.findall(".//w:t", ns):
                text_chunks.append(t.text or "")
            text = "".join(text_chunks)

            style_id = ""
            pstyle = para.find("w:pPr/w:pStyle", ns)
            if pstyle is not None:
                style_id = pstyle.attrib.get(f"{{{W_NS}}}val", "")

            style_name = styles_by_id.get(style_id, style_id)
            yield text, style_name


def build_tree(docx_path: str):
    root = {"title": "HNE Grossing Guide", "level": 0, "children": [], "content": []}
    stack = [root]

    for paragraph_text, style_name in iter_paragraphs_with_styles(docx_path):
        txt = (paragraph_text or "").rstrip()
        if not txt.strip():
            if stack[-1]["content"] and stack[-1]["content"][-1] != "":
                stack[-1]["content"].append("")
            continue

        lvl = heading_level(style_name)
        if lvl:
            node = {"title": txt.strip(), "level": lvl, "children": [], "content": []}
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
        node["path"] = path_titles + [node["title"]] if node["level"] > 0 else [node["title"]]
        for ch in node.get("children", []):
            assign_ids(ch, node["path"])

    assign_ids(root, [])
    return {"generated_from": os.path.basename(DOCX_PATH), "root": root}


def escape_js_string(s: str):
    return s.replace("\\", "\\\\").replace("</", "<\\/")


def main():
    if not os.path.exists(DOCX_PATH):
        raise SystemExit(f"Can't find {DOCX_PATH} in this folder.")

    data = build_tree(DOCX_PATH)

    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

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
        flags=_re.DOTALL,
    )

    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write(html_txt)

    print(f"Wrote {OUT_JSON} and updated {OUT_HTML}")


if __name__ == "__main__":
    main()
