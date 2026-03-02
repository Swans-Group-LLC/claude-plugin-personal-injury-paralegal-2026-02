#!/usr/bin/env python3
"""Accept all tracked changes in a .docx file to produce a clean document.

Processes the raw Word XML to:
- Unwrap <w:ins> elements (keep inserted content, remove wrapper)
- Remove <w:del> elements entirely (discard deleted content)
- Remove <w:rPrChange> elements (accept formatting changes)
- Remove <w:pPrChange> elements (accept paragraph formatting changes)
- Remove <w:del> markers inside <w:rPr> of <w:pPr> (paragraph-mark deletions)

Usage:
    python3 accept-tracked-changes.py <input.docx> <output.docx>
"""

import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def accept_changes(xml_path: str) -> None:
    """Accept all tracked changes in a Word document.xml file in place."""
    parser = etree.XMLParser(remove_blank_text=False)
    tree = etree.parse(xml_path, parser)
    root = tree.getroot()

    # 1. Unwrap <w:ins> — keep child runs, remove the <w:ins> wrapper
    for ins in root.iter(f"{W}ins"):
        parent = ins.getparent()
        if parent is None:
            continue
        idx = list(parent).index(ins)
        children = list(ins)
        for i, child in enumerate(children):
            parent.insert(idx + i, child)
        parent.remove(ins)

    # 2. Remove <w:del> elements entirely (discard deleted content)
    for del_elem in list(root.iter(f"{W}del")):
        parent = del_elem.getparent()
        if parent is None:
            continue
        # Only remove block-level <w:del> (those containing <w:r> with <w:delText>)
        # and paragraph-mark <w:del> markers inside <w:rPr>
        parent.remove(del_elem)

    # 3. Remove <w:rPrChange> elements (accept run formatting changes)
    for rpr_change in list(root.iter(f"{W}rPrChange")):
        parent = rpr_change.getparent()
        if parent is not None:
            parent.remove(rpr_change)

    # 4. Remove <w:pPrChange> elements (accept paragraph formatting changes)
    for ppr_change in list(root.iter(f"{W}pPrChange")):
        parent = ppr_change.getparent()
        if parent is not None:
            parent.remove(ppr_change)

    # 5. Remove <w:sectPrChange> elements (accept section formatting changes)
    for sect_change in list(root.iter(f"{W}sectPrChange")):
        parent = sect_change.getparent()
        if parent is not None:
            parent.remove(sect_change)

    # 6. Remove empty paragraphs left behind by deletions.
    #    After accepting changes, some <w:p> elements contain no text — only
    #    empty <w:pPr> and possibly <w:numPr> (list formatting). These render
    #    as blank bullets, blank numbered items, or blank lines in the PDF.
    body = root.find(f"{W}body")
    if body is not None:
        def _get_text(p):
            return "".join(t.text for t in p.iter(f"{W}t") if t.text)

        def _has_numpr(p):
            ppr = p.find(f"{W}pPr")
            return ppr is not None and ppr.find(f"{W}numPr") is not None

        # Pass 1: Remove all empty paragraphs that have list formatting
        for p in list(body.findall(f"{W}p")):
            if not _get_text(p).strip() and _has_numpr(p):
                body.remove(p)

        # Pass 2: Collapse runs of consecutive empty paragraphs.
        # Keep at most one empty paragraph between any two non-empty paragraphs.
        prev_was_empty = False
        for p in list(body.findall(f"{W}p")):
            if not _get_text(p).strip():
                if prev_was_empty:
                    body.remove(p)
                else:
                    prev_was_empty = True
            else:
                prev_was_empty = False

    tree.write(xml_path, xml_declaration=True, encoding="UTF-8", standalone=True)


def main():
    if len(sys.argv) != 3:
        print(f"Usage: {sys.argv[0]} <input.docx> <output.docx>", file=sys.stderr)
        sys.exit(1)

    input_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])

    if not input_path.is_file():
        print(f"Error: input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    # Work in a temp directory: unpack, process, repack
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        unpack_dir = tmp / "unpacked"

        # Unpack the .docx (it's a zip)
        with zipfile.ZipFile(input_path, "r") as zf:
            zf.extractall(unpack_dir)

        # Accept changes in document.xml
        doc_xml = unpack_dir / "word" / "document.xml"
        if not doc_xml.is_file():
            print("Error: word/document.xml not found in .docx", file=sys.stderr)
            sys.exit(1)

        accept_changes(str(doc_xml))

        # Also process headers and footers if present
        word_dir = unpack_dir / "word"
        for header_footer in word_dir.glob("header*.xml"):
            accept_changes(str(header_footer))
        for header_footer in word_dir.glob("footer*.xml"):
            accept_changes(str(header_footer))

        # Repack into a new .docx
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for file_path in sorted(unpack_dir.rglob("*")):
                if file_path.is_file():
                    arcname = file_path.relative_to(unpack_dir)
                    zf.write(file_path, arcname)

    print(f"Clean document written to: {output_path}")


if __name__ == "__main__":
    main()
