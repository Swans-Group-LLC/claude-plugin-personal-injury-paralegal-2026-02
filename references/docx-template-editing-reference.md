# DOCX Template Editing Reference Guide

This guide tells you exactly how to populate a .docx template that uses `{PLACEHOLDER}` fields and `{{AI SECTION}}` blocks with actual case data, using tracked changes so attorneys can review what was filled in. It applies to demand letters, settlement agreements, or any similar document drafting skill.

## Overview of the Approach

You will write a Python script that uses `lxml` to edit the raw XML inside the .docx file. The workflow is: **unpack the .docx → edit document.xml → repack the .docx**. Do NOT use `python-docx` for this — it cannot produce tracked changes. Do NOT use the Edit tool on the XML — there are too many changes and multi-paragraph AI sections to handle that way.

## Step-by-Step Process

### 1. Unpack the Template

```bash
cp "path/to/template.docx" ./draft.docx
python3 /path/to/skills/docx/scripts/office/unpack.py ./draft.docx ./unpacked/
```

The unpack script merges adjacent runs (so a placeholder like `{Plaintiff Name}` that Word may have split across multiple `<w:r>` elements will be in a single run) and converts smart quotes to XML entities. This is critical — without run merging, placeholder text can be split across runs and become impossible to find.

### 2. Analyze the Template BEFORE Writing Your Script

Before writing any editing code, extract two things from the unpacked `word/document.xml`:

**A. Paragraph spacing.** Run this to see what spacing the template's body paragraphs use:

```python
from lxml import etree
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
tree = etree.parse('./unpacked/word/document.xml')
body = tree.getroot().find(f'{W}body')
for i, p in enumerate(body.findall(f'{W}p')):
    texts = [t.text for t in p.iter(f'{W}t') if t.text]
    text = ''.join(texts)[:80]
    ppr = p.find(f'{W}pPr')
    spacing = ppr.find(f'{W}spacing') if ppr is not None else None
    if spacing is not None:
        before = spacing.get(f'{W}before', '-')
        after = spacing.get(f'{W}after', '-')
        print(f'P{i}: spacing(before={before}, after={after}): {text}')
```

You MUST match this spacing on all inserted paragraphs. If the template uses `before=240, after=240`, your inserted paragraphs must use those same values. Getting this wrong produces visible spacing mismatches that attorneys will notice immediately.

**B. Run properties (rPr).** Check what font, size, and other formatting the body text uses. Look at a normal body paragraph's `<w:rPr>`:

```xml
<w:rPr>
  <w:rFonts w:ascii="Times New Roman" w:eastAsia="Arial" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
  <w:sz w:val="22"/>
  <w:szCs w:val="22"/>
</w:rPr>
```

Store this as your `STD_RPR_XML` constant. All inserted text must use these exact properties.

### 3. Detect Smart Quotes in AI Section Markers

**CRITICAL PITFALL:** The unpack script converts straight quotes to smart quotes (Unicode curly quotes). So the template text `{{AI SECTION END "Liability"}}` becomes `{{AI SECTION END \u201CLiability\u201D}}` in the parsed XML.

Do NOT try to match the exact end marker string with an f-string. Instead, use a fuzzy match:

```python
# WRONG - will fail because of smart quotes:
if f'{{{{AI SECTION END "{section_name}"}}}}' in text:

# CORRECT - match on content, not exact punctuation:
if 'AI SECTION END' in text and section_name in text:
```

This bit us hard and is easy to miss because `repr()` output looks identical for smart and straight quotes.

### 4. Process AI Sections BEFORE Simple Placeholders

**CRITICAL ORDER OF OPERATIONS:**

1. Replace AI sections first
2. Then do simple placeholder replacements
3. Then handle any paragraph deletions

If you do simple replacements first, they will modify placeholder text inside AI section example content (e.g., replacing `{Defendant Name}` inside the example text), which corrupts the text but doesn't cause the AI section markers themselves to break. However, modifying the tree structure during simple replacements can cause index shifts that make AI section paragraph indices wrong. Always do AI sections first.

### 5. How to Delete Paragraphs (AI Sections and Unwanted Lines)

When deleting a range of paragraphs (like an entire AI section), you must do THREE things for each paragraph:

**A. Replace the run content with `<w:del>`:**
```xml
<w:del w:id="1" w:author="Claude" w:date="2026-01-01T00:00:00Z">
  <w:r><w:rPr>...</w:rPr><w:delText>original text here</w:delText></w:r>
</w:del>
```

**B. Mark the paragraph mark as deleted** by adding `<w:del/>` inside `<w:pPr><w:rPr>`:
```xml
<w:pPr>
  <w:rPr>
    <w:del w:id="2" w:author="Claude" w:date="2026-01-01T00:00:00Z"/>
  </w:rPr>
</w:pPr>
```

**C. Do this on ALL paragraphs including the LAST one in the range.** This is critical. If you skip the last paragraph's paragraph-mark deletion, the deleted section will leave an empty gap/line between the preceding heading and the new inserted content. This gap is visible even after accepting all changes and looks wrong.

### 6. How to Insert New Paragraphs

**NEVER copy `<w:pPr>` from deleted paragraphs.** This was a major bug — if you copy the pPr from a paragraph you just marked as deleted, the copied pPr carries the `<w:del>` marker, which causes your new inserted paragraph to also be marked for deletion. Always build a fresh, clean pPr:

```python
def make_clean_ppr(spacing_before=240, spacing_after=240):
    ppr = etree.SubElement(etree.Element('dummy'), f'{W}pPr')
    spacing = etree.SubElement(ppr, f'{W}spacing')
    spacing.set(f'{W}before', str(spacing_before))
    spacing.set(f'{W}after', str(spacing_after))
    jc = etree.SubElement(ppr, f'{W}jc')
    jc.set(f'{W}val', 'both')  # or whatever justification the template uses
    rpr_in_ppr = etree.SubElement(ppr, f'{W}rPr')
    # Add standard font/size properties here
    return ppr
```

Wrap inserted text in `<w:ins>`:
```xml
<w:ins w:id="3" w:author="Claude" w:date="2026-01-01T00:00:00Z">
  <w:r><w:rPr>STANDARD_RPR</w:rPr><w:t>inserted text</w:t></w:r>
</w:ins>
```

Insert new paragraphs AFTER the last deleted paragraph of the AI section:
```python
body_insert_idx = body_children.index(last_deleted_paragraph) + 1
for j, text in enumerate(new_paragraphs):
    new_p = make_paragraph(text)
    body.insert(body_insert_idx + j, new_p)
```

### 7. How to Replace Simple Placeholders

For each `{PLACEHOLDER}` in a `<w:t>` element:

1. Find the run containing it
2. Copy the run's `<w:rPr>` and REMOVE any `<w:highlight>` element (template placeholders often have cyan highlighting)
3. If the entire run text equals the placeholder, replace it with a `<w:del>` + `<w:ins>` pair
4. If the placeholder is part of larger text, split into: text-before + `<w:del>` + `<w:ins>` + text-after

Always apply the cleaned rPr (highlight removed) to the inserted run so the replacement text has the same font/size as surrounding text but without the placeholder highlighting.

### 8. Repack and Validate

```bash
rm -f ./draft.docx
python3 /path/to/skills/docx/scripts/office/pack.py ./unpacked/ ./draft.docx \
  --original "path/to/template.docx" --validate false
```

Use `--validate false` because tracked changes with `<w:del>` inside `<w:pPr><w:rPr>` may trigger a schema validation error about expecting `<w:rPrChange>`. This is a false positive — the document renders correctly in Word and LibreOffice despite this warning.

### 9. Verify the Output (Internal Only)

For your own verification during drafting, you may convert to PDF and inspect it. However:

- **Do NOT share the PDF with the user during the drafting/revision phase.** The PDF will render tracked changes as visible markup, which is confusing.
- Only share the .docx with the user during review. They can open it in Word to accept/reject changes.
- The PDF should only be generated and shared once as the **final clean version** after the user approves.

For internal verification:
```bash
python3 /path/to/skills/docx/scripts/office/soffice.py --headless --convert-to pdf ./draft.docx
```

Also generate a markdown preview with tracked changes visible:
```bash
pandoc ./draft.docx --track-changes=all -o preview.md
```

Check the markdown for: all placeholders replaced (no `{...}` remaining), AI sections showing deletion of template text and insertion of new text, no `[]{.paragraph-deletion}` markers on inserted paragraphs (only on deleted ones).

### 10. PDF Conversion — LibreOffice Crash Handling

**Known issue:** On certain Linux environments (Ubuntu 22 VMs/containers), LibreOffice exits with code 134 and a `malloc_consolidate(): unaligned fastbin chunk detected` error when converting .docx to PDF. The PDF is typically written successfully before the crash, but the non-zero exit code can cause scripts to treat it as a failure.

**Do NOT call soffice directly.** Use the provided wrapper script instead:

```bash
python3 ${CLAUDE_PLUGIN_ROOT}/scripts/docx-convert-to-pdf.py ./input.docx ./output.pdf
```

The wrapper script:
- Tolerates exit code 134 (the known malloc crash-on-exit)
- Validates the output file exists, has a reasonable size (> 1KB), and has a valid PDF header
- Retries once automatically if the first attempt produces no output
- Exits 0 on success, 1 on failure — so you can trust the exit code

If you must call soffice directly for any reason, ALWAYS validate the output afterward:
1. Confirm the PDF file exists
2. Confirm the file size is > 1KB
3. Read the first 5 bytes and verify they are `%PDF-`

Never assume a non-zero exit code means the conversion failed, and never assume a zero exit code (or the existence of a file) means the conversion succeeded.

## Common Pitfalls Summary

| Pitfall | Symptom | Fix |
|---------|---------|-----|
| Smart quotes in AI markers | `find_ai_section_paragraphs` returns `None` for start/end | Use fuzzy matching: `'AI SECTION END' in text and section_name in text` |
| Simple replacements before AI sections | AI section indices shift or content is corrupted | Always process AI sections first |
| Copying pPr from deleted paragraphs | Inserted paragraphs disappear or collapse into headings | Build fresh pPr with `make_clean_ppr()`, never `copy.deepcopy(ref_ppr)` |
| Missing paragraph-mark deletion on last AI section paragraph | Empty gap/extra line between heading and new content after accepting changes | Mark paragraph mark as deleted on ALL paragraphs including the last one |
| Wrong spacing on inserted paragraphs | Inserted text looks too tight or too loose vs template | Analyze template spacing first, then match exactly in `make_clean_ppr()` |
| Highlight not removed from replaced placeholders | Replacement text shows cyan/yellow background | Strip `<w:highlight>` from copied rPr before applying to inserted runs |
| Using `python-docx` | No tracked changes in output | Use `lxml` on unpacked XML directly |
| Using Edit tool on XML | Dozens of edits needed, very slow, error-prone | Write a Python script instead |
| Not using `--validate false` on pack | Pack fails with schema error about `w:rPrChange` | The error is a false positive; use `--validate false` |
| Editing stale unpacked XML after user modified the .docx | User's manual edits (accepted changes, formatting, content) are silently overwritten | Always re-unpack the .docx from disk before every revision round |

## Revision Rounds — Re-read Before Every Edit

When the user requests changes after the initial draft:

1. The user may have opened the .docx in Word and accepted tracked changes, made manual edits, or reformatted content. The unpacked XML from the initial edit is now **stale and must not be reused**.
2. Before every revision, re-unpack the current .docx from disk:

   ```bash
   rm -rf ./unpacked/
   python3 .../unpack.py ./demand-letter-draft.docx ./unpacked/
   ```

3. Re-analyze the XML (spacing, rPr) since accepting tracked changes alters the XML structure.
4. Apply the edit to the freshly unpacked XML.
5. Repack.

## Multi-Line Content in AI Sections

When drafting AI section replacement text, NEVER put multiple list items or bullet points into a single string with `\n` newlines. Word ignores newline characters inside `<w:t>` — they render as spaces, not line breaks.

Instead, make each list item a separate entry in the `replacement_texts` list:

**WRONG:**
```python
'Section Name': [
    '• Item 1\n• Item 2\n• Item 3',
]
```

**CORRECT:**
```python
'Section Name': [
    '• Item 1',
    '• Item 2',
    '• Item 3',
]
```

Each string in the list becomes its own `<w:p>` paragraph in the document.

Note: The `docx-template-editor.py` script includes a defensive check that auto-splits entries containing `\n`, but you should still follow this convention for clarity.

## Template Placeholder Conventions

For skill authors creating new templates, use these conventions so the editing script can work generically:

- **Simple placeholders:** `{FIELD_NAME}` — single curly braces, ALL_CAPS with underscores. Apply cyan highlight in Word so they're visually obvious. These get simple find-and-replace.
- **AI-drafted sections:** `{{AI SECTION START "Section Name": instructions for the AI}}` and `{{AI SECTION END "Section Name"}}` — double curly braces. Everything between START and END (inclusive) gets deleted and replaced with AI-generated content.
- **Deletable lines:** For optional content that may be removed entirely (like "Future Medical Costs" when not applicable), just use a standard `{PLACEHOLDER}` — the script can find and delete the entire paragraph.

## Script Architecture

Structure your editing script like this:

```python
#!/usr/bin/env python3
"""Generic template editor with tracked changes."""
from lxml import etree
import copy

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# === CONFIGURATION (change per-template) ===
DOC_PATH = './unpacked/word/document.xml'
DATE_STR = '2026-02-13T00:00:00Z'
STD_RPR_XML = '...'  # Extract from template analysis
SPACING_BEFORE = 240  # Extract from template analysis
SPACING_AFTER = 240   # Extract from template analysis

SIMPLE_REPLACEMENTS = { '{Placeholder}': 'Value', ... }
AI_SECTIONS = {
    'Section Name': ['paragraph 1 text', 'paragraph 2 text', ...],
}
PARAGRAPHS_TO_DELETE = ['text that identifies the paragraph to delete']

# === REUSABLE FUNCTIONS (copy as-is) ===
# make_std_rpr(), make_ins_run(), make_del_run(),
# make_clean_ppr(), make_paragraph(), get_paragraph_text(),
# process_simple_replacements(), find_ai_section_paragraphs(),
# replace_ai_section(), delete_paragraph_by_text()

# === EXECUTION ===
def main():
    tree = etree.parse(DOC_PATH, etree.XMLParser(remove_blank_text=False))
    body = tree.getroot().find(f'{W}body')

    # 1. AI sections first
    for name, paragraphs in AI_SECTIONS.items():
        replace_ai_section(body, name, paragraphs)

    # 2. Simple replacements
    process_simple_replacements(body)

    # 3. Paragraph deletions
    for text_match in PARAGRAPHS_TO_DELETE:
        delete_paragraph_by_text(body, text_match)

    tree.write(DOC_PATH, xml_declaration=True, encoding='UTF-8')
```

This separation means each new template/skill only needs to change the CONFIGURATION section. The reusable functions stay the same.
