import os, zipfile, subprocess, shlex
from pathlib import Path
from xml.etree import ElementTree as ET
from bs4 import BeautifulSoup
import pandas as pd

POPPLER = r"C:\Users\geldy\Release-25.07.0-0\poppler-25.07.0\Library\bin\pdftohtml.exe"
IN_DIR  = Path("input")
OUT_DIR = Path("output")
OUT_DIR.mkdir(parents=True, exist_ok=True)

def run_args(args: list[str]):
    """Run a command by passing a list of args. Works reliably on Windows paths."""
    return subprocess.run(args, capture_output=True, text=True)

def looks_scanned_pdf(pdf_path: Path) -> bool:
    probe_xml = OUT_DIR / (pdf_path.stem + "._probe.xml")
    res = run_args([
    POPPLER, "-xml", "-enc", "UTF-8",
    str(pdf_path), str(probe_xml)
    ])
    tiny_or_missing = (not probe_xml.exists()) or probe_xml.stat().st_size < 5000
    if probe_xml.exists():
        probe_xml.unlink(missing_ok=True)
    return res.returncode != 0 or tiny_or_missing

def pdf_to_poppler_xml(pdf_path: Path) -> Path:
    out_xml = OUT_DIR / (pdf_path.stem + ".poppler.xml")
    res = run_args([
    POPPLER, "-xml", "-enc", "UTF-8",
    str(pdf_path), str(out_xml)
    ])
    if res.returncode != 0 or not out_xml.exists():
        raise RuntimeError(res.stderr or "pdftohtml failed")
    return out_xml


def normalize_poppler(poppler_xml: Path) -> Path:
    tree = ET.parse(poppler_xml)
    root = tree.getroot()
    items = []
    for page in root.findall("page"):
        for t in page.findall("text"):
            txt = "".join(t.itertext()).strip()
            if not txt:
                continue
            size = float(t.attrib.get("size", "0"))
            items.append(("heading" if size >= 14 else "para", txt))
    soup = BeautifulSoup(features="xml")
    doc  = soup.new_tag("document")
    soup.append(doc)
    section = None
    for kind, content in items:
        if kind == "heading":
            section = soup.new_tag("section")
            title   = soup.new_tag("title"); title.string = content
            section.append(title); doc.append(section)
        else:
            p = soup.new_tag("p"); p.string = content
            (section or doc).append(p)
    out = OUT_DIR / (poppler_xml.stem.replace(".poppler", "") + ".xml")
    out.write_text(soup.prettify(), encoding="utf-8")
    return out

def read_epub(epub_path: Path) -> dict:
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
        for k in ["title","creator","publisher","identifier","language","date","subject"]:
            el = opf_root.find(f".//dc:{k}", ns)
            if el is not None and el.text: meta[k] = el.text
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
    data = read_epub(epub_path)
    soup = BeautifulSoup(features="xml")
    doc  = soup.new_tag("document")
    meta = soup.new_tag("meta")
    for k, v in data["meta"].items():
        tag = soup.new_tag(k); tag.string = v
        meta.append(tag)
    doc.append(meta)
    for html in data["html_parts"]:
        hs = BeautifulSoup(html, "html.parser")
        sec = soup.new_tag("section")
        title_el = hs.find(["h1","h2"])
        title = soup.new_tag("title"); title.string = (title_el.get_text(strip=True) if title_el else "Section")
        sec.append(title)
        for p in hs.find_all("p"):
            text = p.get_text(" ", strip=True)
            if text:
                np = soup.new_tag("p"); np.string = text
                sec.append(np)
        doc.append(sec)
    soup.append(doc)
    out = OUT_DIR / (epub_path.stem + ".xml")
    out.write_text(soup.prettify(), encoding="utf-8")
    return out

def convert_all():
    rows = []
    IN_DIR.mkdir(exist_ok=True)
    for f in sorted(IN_DIR.glob("*")):
        if f.suffix.lower() == ".pdf":
            try:
                if looks_scanned_pdf(f):
                    rows.append({"file": f.name, "type": "pdf", "status": "NEEDS_OCR"})
                else:
                    pop = pdf_to_poppler_xml(f)
                    cleaned = normalize_poppler(pop)
                    rows.append({"file": f.name, "type": "pdf", "status": "ok", "out": cleaned.name})
            except Exception as e:
                rows.append({"file": f.name, "type": "pdf", "status": f"error: {e}"})
        elif f.suffix.lower() == ".epub":
            try:
                out = normalize_epub_to_xml(f)
                rows.append({"file": f.name, "type": "epub", "status": "ok", "out": out.name})
            except Exception as e:
                rows.append({"file": f.name, "type": "epub", "status": f"error: {e}"})
    pd.DataFrame(rows).to_csv(OUT_DIR / "batch_report.csv", index=False)
    print("Done → see output/ and batch_report.csv")

if __name__ == "__main__":
    convert_all()
