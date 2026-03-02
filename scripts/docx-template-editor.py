#!/usr/bin/env python3
"""
Reusable DOCX template editor with tracked changes.

Populates a .docx template that uses {PLACEHOLDER} fields and {{AI SECTION}} blocks
with actual data, wrapping all changes in Word tracked changes so reviewers can see
exactly what was filled in.

USAGE:
  1. Unpack the .docx:
       python3 scripts/office/unpack.py template.docx unpacked/

  2. Run this script (after configuring — see CONFIGURATION section at bottom):
       python3 docx_template_editor.py

  3. Repack the .docx:
       python3 scripts/office/pack.py unpacked/ output.docx --original template.docx --validate false

See DOCX-TEMPLATE-EDITING-REFERENCE.md for full documentation and pitfall guide.
"""

import copy
from lxml import etree

# ============================================================================
# CONSTANTS
# ============================================================================

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

_next_id = [100]


def _get_id():
    val = _next_id[0]
    _next_id[0] += 1
    return str(val)


# ============================================================================
# LOW-LEVEL XML BUILDERS
# ============================================================================

def make_rpr_from_xml(rpr_xml_str):
    """Parse an rPr XML string into an lxml element."""
    return etree.fromstring(rpr_xml_str)


def make_ins_run(text, rpr_elem, author, date_str):
    """Create a <w:ins> element containing a single run with the given text.

    Args:
        text: The text content to insert.
        rpr_elem: An lxml element for <w:rPr> to apply to the run.
        author: The tracked change author name (e.g., "Claude").
        date_str: ISO date string for the tracked change (e.g., "2026-01-01T00:00:00Z").
    """
    ins = etree.SubElement(etree.Element('dummy'), f'{W}ins')
    ins.set(f'{W}id', _get_id())
    ins.set(f'{W}author', author)
    ins.set(f'{W}date', date_str)
    r = etree.SubElement(ins, f'{W}r')
    r.append(copy.deepcopy(rpr_elem))
    t = etree.SubElement(r, f'{W}t')
    t.text = text
    if text and (text[0] == ' ' or text[-1] == ' '):
        t.set(XML_SPACE, 'preserve')
    return ins


def make_del_run(text, rpr_elem, author, date_str):
    """Create a <w:del> element containing a single run with the given text.

    Args:
        text: The text content being deleted.
        rpr_elem: An lxml element for <w:rPr> (usually copied from the original run).
        author: The tracked change author name.
        date_str: ISO date string for the tracked change.
    """
    del_elem = etree.SubElement(etree.Element('dummy'), f'{W}del')
    del_elem.set(f'{W}id', _get_id())
    del_elem.set(f'{W}author', author)
    del_elem.set(f'{W}date', date_str)
    r = etree.SubElement(del_elem, f'{W}r')
    r.append(copy.deepcopy(rpr_elem))
    dt = etree.SubElement(r, f'{W}delText')
    dt.text = text
    if text and (text[0] == ' ' or text[-1] == ' '):
        dt.set(XML_SPACE, 'preserve')
    return del_elem


def make_clean_ppr(spacing_before, spacing_after, justification, rpr_xml_str):
    """Create a clean <w:pPr> element with specified formatting.

    This builds a fresh pPr from scratch — NEVER copy pPr from deleted paragraphs,
    as the deletion markers will carry over and cause your inserted paragraphs to
    also be marked as deleted.

    Args:
        spacing_before: Paragraph spacing before in twips (e.g., 240 = ~12pt).
        spacing_after: Paragraph spacing after in twips.
        justification: Paragraph justification ("both", "left", "center", "right").
        rpr_xml_str: XML string for the standard run properties.
    """
    ppr = etree.SubElement(etree.Element('dummy'), f'{W}pPr')
    if spacing_before > 0 or spacing_after > 0:
        spacing = etree.SubElement(ppr, f'{W}spacing')
        if spacing_before > 0:
            spacing.set(f'{W}before', str(spacing_before))
        if spacing_after > 0:
            spacing.set(f'{W}after', str(spacing_after))
    jc = etree.SubElement(ppr, f'{W}jc')
    jc.set(f'{W}val', justification)
    rpr_in_ppr = etree.SubElement(ppr, f'{W}rPr')
    for child in make_rpr_from_xml(rpr_xml_str):
        rpr_in_ppr.append(copy.deepcopy(child))
    return ppr


def make_paragraph(text, spacing_before, spacing_after, justification, rpr_xml_str, author, date_str):
    """Create a new <w:p> element with tracked-change insertion.

    Args:
        text: The paragraph text.
        spacing_before: Paragraph spacing before in twips.
        spacing_after: Paragraph spacing after in twips.
        justification: Paragraph justification.
        rpr_xml_str: XML string for the standard run properties.
        author: Tracked change author name.
        date_str: ISO date for tracked change.
    """
    p = etree.SubElement(etree.Element('dummy'), f'{W}p')
    p.append(copy.deepcopy(make_clean_ppr(spacing_before, spacing_after, justification, rpr_xml_str)))
    ins_run = make_ins_run(text, make_rpr_from_xml(rpr_xml_str), author, date_str)
    p.append(ins_run)
    return p


# ============================================================================
# PARAGRAPH UTILITIES
# ============================================================================

def get_paragraph_text(p):
    """Get the full text content of a paragraph by joining all <w:t> elements."""
    return ''.join(t.text for t in p.iter(f'{W}t') if t.text)


def _mark_paragraph_mark_deleted(p, author, date_str):
    """Mark a paragraph's paragraph mark as deleted so it collapses when changes are accepted."""
    ppr = p.find(f'{W}pPr')
    if ppr is not None:
        rpr_in_ppr = ppr.find(f'{W}rPr')
        if rpr_in_ppr is None:
            rpr_in_ppr = etree.SubElement(ppr, f'{W}rPr')
        del_mark = etree.SubElement(rpr_in_ppr, f'{W}del')
        del_mark.set(f'{W}id', _get_id())
        del_mark.set(f'{W}author', author)
        del_mark.set(f'{W}date', date_str)


# ============================================================================
# HIGH-LEVEL OPERATIONS
# ============================================================================

def replace_simple_placeholders(body, replacements, rpr_xml_str, author, date_str):
    """Replace {PLACEHOLDER} text in <w:t> elements with tracked changes.

    For each placeholder found:
    - The original text is wrapped in <w:del> (deletion)
    - The replacement text is wrapped in <w:ins> (insertion)
    - Any <w:highlight> formatting is removed from the replacement

    Args:
        body: The <w:body> lxml element.
        replacements: Dict mapping placeholder strings to replacement values.
        rpr_xml_str: XML string for standard run properties (used as fallback).
        author: Tracked change author name.
        date_str: ISO date for tracked change.
    """
    std_rpr = make_rpr_from_xml(rpr_xml_str)

    for p in body.iter(f'{W}p'):
        full_text = get_paragraph_text(p)
        for placeholder, replacement in replacements.items():
            if placeholder not in full_text:
                continue

            for r in p.findall(f'.//{W}r'):
                t = r.find(f'{W}t')
                if t is None or t.text is None or placeholder not in t.text:
                    continue

                parent = r.getparent()
                idx = list(parent).index(r)

                # Get rPr and make a clean copy (remove highlight)
                rpr = r.find(f'{W}rPr')
                clean_rpr = copy.deepcopy(rpr) if rpr is not None else copy.deepcopy(std_rpr)
                hl = clean_rpr.find(f'{W}highlight')
                if hl is not None:
                    clean_rpr.remove(hl)

                original_text = t.text

                if original_text == placeholder:
                    # Entire run is the placeholder
                    del_elem = make_del_run(original_text, rpr if rpr is not None else std_rpr, author, date_str)
                    ins_elem = make_ins_run(replacement, clean_rpr, author, date_str)
                    parent.remove(r)
                    parent.insert(idx, ins_elem)
                    parent.insert(idx, del_elem)
                else:
                    # Placeholder is part of larger text — split
                    before, after = original_text.split(placeholder, 1)
                    elements = []

                    if before:
                        r_before = etree.SubElement(etree.Element('dummy'), f'{W}r')
                        r_before.append(copy.deepcopy(clean_rpr))
                        t_before = etree.SubElement(r_before, f'{W}t')
                        t_before.text = before
                        if before[0] == ' ' or before[-1] == ' ':
                            t_before.set(XML_SPACE, 'preserve')
                        elements.append(r_before)

                    elements.append(make_del_run(placeholder, rpr if rpr is not None else std_rpr, author, date_str))
                    elements.append(make_ins_run(replacement, clean_rpr, author, date_str))

                    if after:
                        r_after = etree.SubElement(etree.Element('dummy'), f'{W}r')
                        r_after.append(copy.deepcopy(clean_rpr))
                        t_after = etree.SubElement(r_after, f'{W}t')
                        t_after.text = after
                        if after[0] == ' ' or after[-1] == ' ':
                            t_after.set(XML_SPACE, 'preserve')
                        elements.append(r_after)

                    parent.remove(r)
                    for i, elem in enumerate(elements):
                        parent.insert(idx + i, elem)


def find_ai_section(body, section_name):
    """Find start and end paragraph indices for an {{AI SECTION}} block.

    Uses fuzzy matching to handle smart quotes (the unpack script converts
    straight quotes to curly/smart quotes, so exact string matching fails).

    Args:
        body: The <w:body> lxml element.
        section_name: The section name (e.g., "Statement of Liability").

    Returns:
        Tuple of (start_idx, end_idx, paragraphs_list) or (None, None, list) if not found.
    """
    paragraphs = list(body.findall(f'{W}p'))
    start_idx = None
    end_idx = None

    for i, p in enumerate(paragraphs):
        text = get_paragraph_text(p)
        if '{{AI SECTION START' in text and section_name in text:
            start_idx = i
        # Fuzzy match — don't try to match exact quotes around section_name
        if 'AI SECTION END' in text and section_name in text:
            end_idx = i
            break

    return start_idx, end_idx, paragraphs


def replace_ai_section(body, section_name, replacement_texts,
                        spacing_before, spacing_after, justification,
                        rpr_xml_str, author, date_str):
    """Replace an {{AI SECTION}} block with new paragraphs using tracked changes.

    Deletes all paragraphs from START to END (inclusive), marking both content
    and paragraph marks as deleted so the section fully collapses. Then inserts
    new paragraphs with the replacement text after the deleted block.

    Args:
        body: The <w:body> lxml element.
        section_name: The section name to find.
        replacement_texts: List of strings, one per paragraph to insert.
        spacing_before: Paragraph spacing before in twips.
        spacing_after: Paragraph spacing after in twips.
        justification: Paragraph justification.
        rpr_xml_str: XML string for standard run properties.
        author: Tracked change author name.
        date_str: ISO date for tracked change.
    """
    start_idx, end_idx, paragraphs = find_ai_section(body, section_name)

    if start_idx is None or end_idx is None:
        print(f"WARNING: Could not find AI section '{section_name}'")
        return

    print(f"Found AI section '{section_name}' from paragraph {start_idx} to {end_idx}")

    std_rpr = make_rpr_from_xml(rpr_xml_str)

    # Mark all paragraphs in the section as deleted
    for i in range(start_idx, end_idx + 1):
        p = paragraphs[i]
        text = get_paragraph_text(p)

        if not text.strip():
            # Empty paragraph — just mark paragraph mark as deleted
            _mark_paragraph_mark_deleted(p, author, date_str)
            continue

        # Replace runs with deletion markup
        runs_to_remove = list(p.findall(f'{W}r'))
        ins_to_remove = list(p.findall(f'{W}ins'))

        first_rpr = None
        if runs_to_remove:
            first_rpr = runs_to_remove[0].find(f'{W}rPr')

        for r in runs_to_remove:
            p.remove(r)
        for ins in ins_to_remove:
            p.remove(ins)

        del_elem = make_del_run(text, first_rpr if first_rpr is not None else std_rpr, author, date_str)
        p.append(del_elem)

        # Mark paragraph mark as deleted on ALL paragraphs (including last!)
        # so the deleted section fully collapses with no leftover empty space
        _mark_paragraph_mark_deleted(p, author, date_str)

    # Insert new paragraphs after the deleted section
    ref_parent = paragraphs[end_idx].getparent()
    body_children = list(ref_parent)
    body_insert_idx = body_children.index(paragraphs[end_idx]) + 1

    for j, text in enumerate(replacement_texts):
        new_p = make_paragraph(text, spacing_before, spacing_after, justification,
                               rpr_xml_str, author, date_str)
        ref_parent.insert(body_insert_idx + j, new_p)


def delete_paragraph_by_text(body, match_text, rpr_xml_str, author, date_str):
    """Find a paragraph containing match_text and mark it entirely as deleted.

    Args:
        body: The <w:body> lxml element.
        match_text: Text to search for in paragraph content.
        rpr_xml_str: XML string for standard run properties.
        author: Tracked change author name.
        date_str: ISO date for tracked change.
    """
    std_rpr = make_rpr_from_xml(rpr_xml_str)

    for p in body.findall(f'{W}p'):
        text = get_paragraph_text(p)
        if match_text not in text:
            continue

        runs = list(p.findall(f'{W}r'))
        first_rpr = runs[0].find(f'{W}rPr') if runs else None

        for r in runs:
            p.remove(r)

        if text.strip():
            del_elem = make_del_run(text, first_rpr if first_rpr is not None else std_rpr, author, date_str)
            p.append(del_elem)

        _mark_paragraph_mark_deleted(p, author, date_str)
        print(f"Deleted paragraph containing '{match_text[:50]}...'")
        return

    print(f"WARNING: Could not find paragraph containing '{match_text[:50]}...'")


def edit_template(doc_path, simple_replacements, ai_sections, paragraphs_to_delete,
                  rpr_xml_str, spacing_before=240, spacing_after=240,
                  justification='both', author='Claude', date_str='2026-01-01T00:00:00Z'):
    """Main entry point: edit an unpacked document.xml with all replacements.

    Args:
        doc_path: Path to the unpacked word/document.xml.
        simple_replacements: Dict of {'{Placeholder}': 'value'}.
        ai_sections: Dict of {'Section Name': ['para1 text', 'para2 text', ...]}.
        paragraphs_to_delete: List of match strings to identify paragraphs to delete entirely.
        rpr_xml_str: XML string for the standard body text run properties.
        spacing_before: Paragraph spacing before in twips (default 240).
        spacing_after: Paragraph spacing after in twips (default 240).
        justification: Paragraph justification (default 'both').
        author: Tracked change author name (default 'Claude').
        date_str: ISO date string for tracked changes.
    """
    parser = etree.XMLParser(remove_blank_text=False)
    tree = etree.parse(doc_path, parser)
    body = tree.getroot().find(f'{W}body')

    # 0. Normalize AI sections: split any entry containing newlines into separate paragraphs.
    #    Word ignores \n inside <w:t> — they render as spaces, not line breaks.
    #    Each string in the list must be a single paragraph.
    for section_name in list(ai_sections.keys()):
        expanded = []
        for text in ai_sections[section_name]:
            if '\n' in text:
                expanded.extend(line for line in text.split('\n') if line.strip())
            else:
                expanded.append(text)
        ai_sections[section_name] = expanded

    # 1. AI sections FIRST (before simple replacements modify the tree)
    print("Replacing AI sections...")
    for section_name, texts in ai_sections.items():
        replace_ai_section(body, section_name, texts,
                           spacing_before, spacing_after, justification,
                           rpr_xml_str, author, date_str)

    # 2. Simple placeholder replacements
    print("Processing simple replacements...")
    replace_simple_placeholders(body, simple_replacements, rpr_xml_str, author, date_str)

    # 3. Paragraph deletions
    print("Processing paragraph deletions...")
    for match_text in paragraphs_to_delete:
        delete_paragraph_by_text(body, match_text, rpr_xml_str, author, date_str)

    # Write back
    tree.write(doc_path, xml_declaration=True, encoding='UTF-8')
    print(f"Done! Written to {doc_path}")


# ============================================================================
# EXAMPLE USAGE (replace with your actual configuration)
# ============================================================================

if __name__ == '__main__':
    # This is an EXAMPLE showing how to call edit_template().
    # Replace everything below with your actual template data.

    W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    EXAMPLE_RPR = f'''<w:rPr xmlns:w="{W_NS}">
      <w:rFonts w:ascii="Times New Roman" w:eastAsia="Arial" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
      <w:sz w:val="22"/>
      <w:szCs w:val="22"/>
    </w:rPr>'''

    edit_template(
        doc_path='./unpacked/word/document.xml',
        simple_replacements={
            '{Client Name}': 'Jane Doe',
            '{Date}': 'January 1, 2026',
        },
        ai_sections={
            'Section Name': [
                'First paragraph of the AI-drafted section.',
                'Second paragraph of the AI-drafted section.',
            ],
        },
        paragraphs_to_delete=[
            'Optional Line To Delete',
        ],
        rpr_xml_str=EXAMPLE_RPR,
        spacing_before=240,
        spacing_after=240,
        justification='both',
        author='Claude',
        date_str='2026-01-01T00:00:00Z',
    )
