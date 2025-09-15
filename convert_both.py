import os
import re
import zipfile
import subprocess
from pathlib import Path
from typing import List, Tuple
from xml.etree import ElementTree as ET

from bs4 import BeautifulSoup
import pandas as pd

# --------------------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------------------
# If Poppler is on PATH, leave as "pdftohtml".
# Otherwise put the full path to pdftohtml.exe, e.g.:
# POPPLER = r"C:\Users\YOU\poppler\Library\bin\pdftohtml.exe"
POPPLER = os.getenv("POPPLER_EXE", "pdftohtml")

IN_DIR  = Path("input")
OUT_DIR = Path("output")
OUT_DIR.mkdir(parents=True, exist_ok=True)

# DocBook DOCTYPE identifiers (swap to Kevin's exact values if needed)
DOCBOOK_PUBLIC = '-//RIS Dev//DTD DocBook V4.3 -Based Variant V1.1//EN'
DOCBOOK_SYSTEM = 'http://LOCALHOST/dtd/V1.1/RittDocBook.dtd'

# Heading threshold for PDF → you can tweak this per corpus
PDF_HEADING_SIZE_THRESHOLD = 14.0


# --------------------------------------------------------------------------------------
# UTILS
# --------------------------------------------------------------------------------------
def run_args(args: List[str]) -> subprocess.CompletedProcess:
    """Run a command by passing a list of args (Windows-safe)."""
    return subprocess.run(args, capture_output=True, text=True)

def escape_xml(t: str) -> str:
    """Minimal XML escaping for text nodes."""
    return (
        t.replace("&", "&amp;")
         .replace("<", "&lt;")
         .replace(">", "&gt;")
    )

def slug(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip()


# --------------------------------------------------------------------------------------
# PDF → Poppler XML → Clean XML (sections, titles, paragraphs)
# --------------------------------------------------------------------------------------
def looks_scanned_pdf(pdf_path: Path) -> bool:
    """
    Quick heuristic: attempt a tiny Poppler export; if it fails or is tiny,
    assume scanned (no extractable text).
    """
    probe_xml = OUT_DIR / (pdf_path.stem + "._probe.xml")
    res = run_args([POPPLER, "-xml", "-enc", "UTF-8", str(pdf_path), str(probe_xml)])
    tiny_or_missing = (not probe_xml.exists()) or probe_xml.stat().st_size < 5000
    if probe_xml.exists():
        probe_xml.unlink(missing_ok=True)
    return res.returncode != 0 or tiny_or_missing

def pdf_to_poppler_xml(pdf_path: Path) -> Path:
    out_xml = OUT_DIR / (pdf_path.stem + ".poppler.xml")
    res = run_args([POPPLER, "-xml", "-enc", "UTF-8", str(pdf_path), str(out_xml)])
    if res.returncode != 0 or not out_xml.exists():
        raise RuntimeError(res.stderr or "pdftohtml failed")
    return out_xml

def normalize_poppler(poppler_xml: Path) -> Path:
    """
    Convert Poppler layout XML into a simple content XML:
      <document>
        <section><title>…</title><p>…</p>…</section>
        ...
      </document>
    Sections are created when we hit a 'heading' (by font size).
    """
    tree = ET.parse(poppler_xml)
    root = tree.getroot()

    # Collect runs with font sizes so we can chunk by headings
    runs = []
    for page in root.findall("page"):
        for t in page.findall("text"):
            txt = "".join(t.itertext()).strip()
            if not txt:
                continue
            size = float(t.attrib.get("size", "0"))
            runs.append({"text": txt, "size": size})

    # Group into sections
    sections: List[Tuple[str, List[str]]] = []  # (title, [paras])
    current_title = None
    current_paras: List[str] = []

    for r in runs:
        if r["size"] >= PDF_HEADING_SIZE_THRESHOLD:
            # Commit previous section
            if current_title is not None or current_paras:
                sections.append((slug(current_title) if current_title else "", current_paras))
                current_paras = []
            current_title = r["text"]
        else:
            current_paras.append(r["text"])

    # Final section commit
    if current_title is not None or current_paras:
        sections.append((slug(current_title) if current_title else "", current_paras))

    # Build cleaned XML
    soup = BeautifulSoup(features="xml")
    doc = soup.new_tag("document")
    soup.append(doc)

    for title, paras in sections:
        sec = soup.new_tag("section")
        t = soup.new_tag("title")
        t.string = title if title else "Section"
        sec.append(t)
        for p in paras:
            if p.strip():
                np = soup.new_tag("p")
                np.string = p.strip()
                sec.append(np)
        doc.append(sec)

    out = OUT_DIR / (poppler_xml.stem.replace(".poppler", "") + ".xml")
    out.write_text(soup.prettify(), encoding="utf-8")
    return out


# --------------------------------------------------------------------------------------
# EPUB → Combined Clean XML (sections, titles, paragraphs)
# --------------------------------------------------------------------------------------
def read_epub(epub_path: Path) -> dict:
    """
    Return: {
        "meta": {title, creator, publisher, identifier, language, date, subject},
        "html_parts": [html_string_in_spine_order, ...]
    }
    """
    with zipfile.ZipFile(epub_path, "r") as z:
        with z.open("META-INF/container.xml") as f:
            c = ET.parse(f).getroot()
            ns = {"c": "urn:oasis:names:tc:opendocument:xmlns:container"}
            rootfile = c.find(".//c:rootfile", ns).attrib["full-path"]

        with z.open(rootfile) as f:
            opf_root = ET.parse(f).getroot()
        opf_dir = Path(rootfile).parent

        ns = {"dc": "http://purl.org/dc/elements/1.1/"}
        meta = {}
        for k in ["title", "creator", "publisher", "identifier", "language", "date", "subject"]:
            el = opf_root.find(f".//dc:{k}", ns)
            if el is not None and el.text:
                meta[k] = el.text

        id_href = {item.attrib["id"]: item.attrib["href"] for item in opf_root.findall(".//{*}item")}
        spine_ids = [it.attrib["idref"] for it in opf_root.findall(".//{*}spine/{*}itemref")]
        hrefs = [id_href.get(i) for i in spine_ids if id_href.get(i)]

        html_parts = []
        with zipfile.ZipFile(epub_path, "r") as z2:
            for href in hrefs:
                p = str(opf_dir / href).replace("\\", "/")
                if p in z2.namelist():
                    with z2.open(p) as f:
                        html_parts.append(f.read().decode("utf-8", errors="ignore"))
    return {"meta": meta, "html_parts": html_parts}

def normalize_epub_to_xml(epub_path: Path) -> Path:
    """
    Combine EPUB spine content into:
      <document>
        <meta>...</meta>
        <section><title>...</title><p>...</p>...</section>
        ...
      </document>
    """
    data = read_epub(epub_path)
    soup = BeautifulSoup(features="xml")
    doc = soup.new_tag("document")
    soup.append(doc)

    # Meta block (optional)
    meta = soup.new_tag("meta")
    for k, v in data["meta"].items():
        tag = soup.new_tag(k)
        tag.string = v
        meta.append(tag)
    doc.append(meta)

    # One section per spine file; title = first h1/h2 or "Section"
    for html in data["html_parts"]:
        hs = BeautifulSoup(html, "html.parser")
        sec = soup.new_tag("section")
        title_el = hs.find(["h1", "h2"])
        t = soup.new_tag("title")
        t.string = title_el.get_text(strip=True) if title_el else "Section"
        sec.append(t)
        for p in hs.find_all("p"):
            txt = p.get_text(" ", strip=True)
            if txt:
                np = soup.new_tag("p")
                np.string = txt
                sec.append(np)
        doc.append(sec)

    out = OUT_DIR / (epub_path.stem + ".xml")
    out.write_text(soup.prettify(), encoding="utf-8")
    return out


# --------------------------------------------------------------------------------------
# DocBook assembly: write chapter files + master book.xml w/ DOCTYPE + entities
# --------------------------------------------------------------------------------------
def write_chapter_file(out_dir: Path, idx: int, title: str, paragraphs: List[str]) -> Path:
    """
    Write one DocBook chapter file:
      <chapter>
        <title>...</title>
        <para>...</para>
      </chapter>
    Returns path to chNNNN.xml
    """
    ch_name = f"ch{idx:04d}.xml"
    ch_path = out_dir / ch_name
    lines = ['<?xml version="1.0" encoding="UTF-8"?>', '<chapter>']
    if title:
        lines.append(f'  <title>{escape_xml(title)}</title>')
    for p in paragraphs:
        p = p.strip()
        if p:
            lines.append(f'  <para>{escape_xml(p)}</para>')
    lines.append('</chapter>')
    ch_path.write_text("\n".join(lines), encoding="utf-8")
    return ch_path

def write_book_master(out_dir: Path, chapters: List[Tuple[str, Path]], meta: dict) -> Path:
    """
    chapters: list of (entity_name, chapter_path)
    meta keys (optional): title, author, isbn, publisher
    """
    book_path = out_dir / "book.xml"
    entity_lines = [f'  <!ENTITY {name} SYSTEM "{path.name}">' for name, path in chapters]

    # bookinfo (fill more fields as needed)
    bi = ["  <bookinfo>"]
    if meta.get("title"):     bi.append(f'    <title>{escape_xml(meta["title"])}</title>')
    if meta.get("author"):    bi.append(f'    <author>{escape_xml(meta["author"])}</author>')
    if meta.get("isbn"):      bi.append(f'    <isbn>{escape_xml(meta["isbn"])}</isbn>')
    if meta.get("publisher"): bi.append(f'    <publisher>{escape_xml(meta["publisher"])}</publisher>')
    bi.append("  </bookinfo>")

    body = [f'  &{name};' for name, _ in chapters]

    content = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        f'<!DOCTYPE book PUBLIC "{DOCBOOK_PUBLIC}"',
        f'  "{DOCBOOK_SYSTEM}" [',
        *entity_lines,
        ']>',
        '<book>',
        *bi,
        *body,
        '</book>'
    ]
    book_path.write_text("\n".join(content), encoding="utf-8")
    return book_path

def docbook_from_epub_combined(combined_xml_path: Path, meta: dict):
    """
    Takes combined EPUB XML (<document><section>...</section>...) and writes:
      ch0000.xml, ch0001.xml, ... + book.xml with entities.
    """
    soup = BeautifulSoup(combined_xml_path.read_text(encoding="utf-8"), "xml")
    out_dir = combined_xml_path.parent

    chapters: List[Tuple[str, Path]] = []
    idx = 0
    for sec in soup.find_all("section"):
        title_el = sec.find("title")
        title = slug(title_el.get_text(" ", strip=True)) if title_el else f"Section {idx}"
        paras = [p.get_text(" ", strip=True) for p in sec.find_all("p")]
        ch_path = write_chapter_file(out_dir, idx, title, paras)
        chapters.append((f"ch{idx:04d}", ch_path))
        idx += 1

    write_book_master(out_dir, chapters, meta)

def docbook_from_pdf_clean(clean_xml_path: Path, meta: dict):
    """
    Takes cleaned PDF XML (<document><section>...</section>...) and writes:
      ch0000.xml, ch0001.xml, ... + book.xml with entities.
    """
    soup = BeautifulSoup(clean_xml_path.read_text(encoding="utf-8"), "xml")
    out_dir = clean_xml_path.parent

    chapters: List[Tuple[str, Path]] = []
    sections = soup.find_all("section")
    if sections:
        for idx, sec in enumerate(sections):
            title_el = sec.find("title")
            title = slug(title_el.get_text(" ", strip=True)) if title_el else f"Chapter {idx}"
            paras = [p.get_text(" ", strip=True) for p in sec.find_all("p")]
            ch_path = write_chapter_file(out_dir, idx, title, paras)
            chapters.append((f"ch{idx:04d}", ch_path))
    else:
        # Fallback: all <p> as one chapter
        paras = [p.get_text(" ", strip=True) for p in soup.find_all("p")]
        ch_path = write_chapter_file(out_dir, 0, "Chapter 0", paras)
        chapters.append(("ch0000", ch_path))

    write_book_master(out_dir, chapters, meta)


# --------------------------------------------------------------------------------------
# BATCH DRIVER
# --------------------------------------------------------------------------------------
def convert_all():
    rows = []
    IN_DIR.mkdir(exist_ok=True)

    for f in sorted(IN_DIR.glob("*")):
        ext = f.suffix.lower()

        # ---------------- PDF ----------------
        if ext == ".pdf":
            try:
                if looks_scanned_pdf(f):
                    rows.append({"file": f.name, "type": "pdf", "status": "NEEDS_OCR"})
                    continue

                pop = pdf_to_poppler_xml(f)
                cleaned = normalize_poppler(pop)

                # Minimal metadata placeholders; adjust/derive per project
                meta = {
                    "title":     f.stem,    # or parse from PDF metadata if available
                    "author":    "",
                    "isbn":      "",
                    "publisher": "Rittenhouse"
                }
                docbook_from_pdf_clean(cleaned, meta)

                rows.append({"file": f.name, "type": "pdf", "status": "ok",
                             "clean_xml": cleaned.name, "book": "book.xml"})
            except Exception as e:
                rows.append({"file": f.name, "type": "pdf", "status": f"error: {e}"})

        # ---------------- EPUB ----------------
        elif ext == ".epub":
            try:
                combined = normalize_epub_to_xml(f)

                # Pull some meta from EPUB (already embedded in combined file if present)
                # You can re-parse combined for meta, or leave placeholders:
                meta = {
                    "title":     f.stem,
                    "author":    "",
                    "isbn":      "",
                    "publisher": "Rittenhouse"
                }
                docbook_from_epub_combined(combined, meta)

                rows.append({"file": f.name, "type": "epub", "status": "ok",
                             "clean_xml": combined.name, "book": "book.xml"})
            except Exception as e:
                rows.append({"file": f.name, "type": "epub", "status": f"error: {e}"})

        # ------------- skip others ------------
        else:
            rows.append({"file": f.name, "type": ext, "status": "skipped"})

    pd.DataFrame(rows).to_csv(OUT_DIR / "batch_report.csv", index=False)
    print("Done → see output/ and batch_report.csv")


if __name__ == "__main__":
    convert_all()
