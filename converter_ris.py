#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Minimal RIS-ready converter for PDF/EPUB → book.xml + chXXXX.xml + images + package.zip

Usage (Windows PowerShell):
  .venv\Scripts\activate
  python converter_ris.py --input "input/LDDB 2025-2027 paperback edition.05.26.25.final.pdf" ^
                          --isbn 9781234567890 --title "Little Dental Drug Booklet 2025–2027" ^
                          --author "Stanley F. Malamed" --publisher "Rittenhouse"

  python converter_ris.py --input input/book.epub --isbn 9781234567890 --type epub ...

Notes:
- If pdftohtml is not on PATH, set POPPLER_BIN below.
- This outputs to out/<ISBN>/ and creates package-<ISBN>.zip
"""

import os, re, json, zipfile, shutil, subprocess, argparse
from pathlib import Path
from datetime import datetime
from bs4 import BeautifulSoup
from lxml import etree
from ebooklib import epub

POPPLER_BIN = r"C:\Users\YOURUSER\Release-25.07.0-0\poppler-25.07.0\Library\bin\pdftohtml.exe"  # update if needed

RIS_NS = "urn:ris:r2"  # temporary namespace until final schema is supplied
NSMAP = {None: RIS_NS}

def safe_stem(name: str) -> str:
    return re.sub(r"[^a-zA-Z0-9._-]+", "_", name)

def run_pdftohtml(pdf, out_dir):
    out_xml = out_dir / (Path(pdf).stem + ".poppler.xml")
    cmd = [POPPLER_BIN, "-xml", "-nodrm", "-zoom", "1.5", str(pdf), str(out_xml)]
    subprocess.run(cmd, check=True)
    return out_xml

def ensure_dir(p: Path):
    p.mkdir(parents=True, exist_ok=True)

def new_elem(tag, text=None, attrib=None):
    el = etree.Element(tag, nsmap=NSMAP)
    if attrib:
        for k, v in attrib.items():
            el.set(k, str(v))
    if text:
        el.text = text
    return el

def write_xml(root: etree._Element, path: Path):
    xml_bytes = etree.tostring(root, xml_declaration=True, encoding="UTF-8", pretty_print=True)
    path.write_bytes(xml_bytes)

def parse_epub_to_chapters(epub_path: Path):
    book = epub.read_epub(str(epub_path))
    chapters = []
    idx = 0
    for item in book.get_items_of_type(ebooklib.ITEM_DOCUMENT):
        idx += 1
        soup = BeautifulSoup(item.get_content(), "lxml")
        # Very light mapping: h1/h2 → <title>, p/li → <p>
        ch = new_elem("chapter", attrib={"id": f"ch{idx:04d}"})
        title = soup.find(["h1", "h2"])
        if title and title.get_text(strip=True):
            ch.append(new_elem("title", title.get_text(strip=True)))
        for tag in soup.find_all(["p", "li"]):
            txt = tag.get_text(" ", strip=True)
            if txt:
                ch.append(new_elem("p", txt))
        chapters.append(ch)
    return chapters

def parse_poppler_xml_to_chapters(poppler_xml: Path):
    """Heuristic: split chapter when we see a centered ALL-CAPS line or 'Chapter ' pattern."""
    soup = BeautifulSoup(poppler_xml.read_text(encoding="utf-8", errors="ignore"), "lxml-xml")
    pages = soup.find_all("page")
    chunks = []
    buf = []
    def flush():
        nonlocal buf
        if buf:
            chunks.append(buf); buf = []
    for pg in pages:
        for t in pg.find_all("text"):
            txt = t.get_text(" ", strip=True)
            if not txt: 
                continue
            # crude chapter boundary detection
            if re.match(r"^(chapter\s+\d+|[A-Z0-9 ,.'/-]{8,})$", txt.strip(), re.IGNORECASE):
                flush()
                buf.append(txt)  # keep heading
            else:
                buf.append(txt)
    flush()
    chapters = []
    for i, chunk in enumerate(chunks, 1):
        ch = new_elem("chapter", attrib={"id": f"ch{i:04d}"})
        if chunk:
            ch.append(new_elem("title", chunk[0][:200]))
            for ptxt in chunk[1:]:
                ch.append(new_elem("p", ptxt))
        chapters.append(ch)
    if not chapters:  # fallback: single chapter
        ch = new_elem("chapter", attrib={"id": "ch0001"})
        for pg in pages:
            for t in pg.find_all("text"):
                txt = t.get_text(" ", strip=True)
                if txt: ch.append(new_elem("p", txt))
        chapters = [ch]
    return chapters

def build_book_xml(isbn, title, author, publisher, chapters):
    book = new_elem("book")
    info = new_elem("bookinfo")
    info.append(new_elem("isbn", isbn))
    info.append(new_elem("title", title))
    if author: info.append(new_elem("author", author))
    if publisher: info.append(new_elem("publisher", publisher))
    info.append(new_elem("created", datetime.utcnow().isoformat() + "Z"))
    book.append(info)
    # externalize chapters to chNNNN.xml files; book.xml keeps <chapterref href="..."/>
    refs = new_elem("contents")
    for i, _ in enumerate(chapters, 1):
        refs.append(new_elem("chapterref", attrib={"href": f"ch{i:04d}.xml"}))
    book.append(refs)
    return book

def validate_wellformed(path: Path):
    try:
        etree.parse(str(path))
        return True, ""
    except Exception as e:
        return False, str(e)

def package_folder(book_dir: Path, isbn: str):
    zpath = book_dir.parent / f"package-{isbn}.zip"
    with zipfile.ZipFile(zpath, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for p in book_dir.rglob("*"):
            if p.is_file():
                z.write(p, p.relative_to(book_dir))
    return zpath

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--isbn", required=True)
    ap.add_argument("--title", default="")
    ap.add_argument("--author", default="")
    ap.add_argument("--publisher", default="Rittenhouse")
    ap.add_argument("--type", choices=["pdf","epub"], help="Force type (else inferred by extension)")
    args = ap.parse_args()

    src = Path(args.input)
    in_type = args.type or src.suffix.lower().lstrip(".")
    out_root = Path("out")
    book_dir = out_root / args.isbn
    images_dir = book_dir / "images"
    ensure_dir(book_dir); ensure_dir(images_dir)

    # 1) Build chapters
    if in_type == "epub":
        chapters = parse_epub_to_chapters(src)
    elif in_type == "pdf":
        poppler_xml = run_pdftohtml(src, book_dir)
        chapters = parse_poppler_xml_to_chapters(poppler_xml)
    else:
        raise SystemExit(f"Unsupported type: {in_type}")

    # 2) Write chapter files
    ch_paths = []
    for i, ch in enumerate(chapters, 1):
        ch_path = book_dir / f"ch{i:04d}.xml"
        root = new_elem("chapterfile")
        root.append(ch)
        write_xml(root, ch_path)
        ch_paths.append(str(ch_path.name))

    # 3) Write book.xml
    book_xml = build_book_xml(args.isbn, args.title, args.author, args.publisher, chapters)
    write_xml(book_xml, book_dir / "book.xml")

    # 4) Minimal manifest + batch report
    manifest = {
        "isbn": args.isbn, "title": args.title, "author": args.author,
        "chapters": ch_paths, "images": [],
        "created": datetime.utcnow().isoformat()+"Z",
        "source": src.name, "source_type": in_type
    }
    (book_dir / "manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    (book_dir / "batch_report.csv").write_text(
        "file,type,status\n" + f"{src.name},{in_type},ok\n", encoding="utf-8"
    )

    # 5) Validate well-formedness (schema validation can be added when available)
    ok1, e1 = validate_wellformed(book_dir / "book.xml")
    errors = []
    if not ok1: errors.append(("book.xml", e1))
    for chp in ch_paths:
        ok, e = validate_wellformed(book_dir / chp)
        if not ok: errors.append((chp, e))
    if errors:
        (book_dir / "validation.log").write_text("\n".join([f"{a}: {b}" for a,b in errors]), encoding="utf-8")

    # 6) Package
    zpath = package_folder(book_dir, args.isbn)
    print(f"Done → {book_dir}\nPackage: {zpath}")

if __name__ == "__main__":
    main()
