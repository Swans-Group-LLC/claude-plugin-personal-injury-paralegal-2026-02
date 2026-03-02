#!/usr/bin/env python3
"""Convert a .docx file to PDF using LibreOffice, with crash-tolerant validation.

LibreOffice on certain Linux environments (Ubuntu 22 containers/VMs) crashes with
exit code 134 and 'malloc_consolidate(): unaligned fastbin chunk detected' after
writing the PDF. This script tolerates that specific crash, validates the output,
and retries once on failure.

Usage:
    python3 docx-convert-to-pdf.py <input.docx> [output.pdf]

If output.pdf is not specified, the PDF is written to the same directory as the
input with the same basename (e.g., draft.docx -> draft.pdf).
"""

import os
import subprocess
import sys
from pathlib import Path


def convert(input_path: str, output_path: str, attempt: int = 1) -> bool:
    """Run LibreOffice conversion and validate the result.

    Returns True if a valid PDF was produced, False otherwise.
    """
    input_p = Path(input_path)
    output_p = Path(output_path)
    outdir = str(output_p.parent)

    # Remove any stale output from a previous attempt
    if output_p.exists():
        output_p.unlink()

    result = subprocess.run(
        [
            "soffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", outdir,
            str(input_p),
        ],
        capture_output=True,
        timeout=120,
    )

    # LibreOffice writes the PDF with the input's basename
    produced = Path(outdir) / (input_p.stem + ".pdf")

    # Rename to desired output path if different
    if produced.exists() and produced != output_p:
        produced.rename(output_p)

    # Validate
    if not output_p.exists():
        print(f"[Attempt {attempt}] PDF file was not created.", file=sys.stderr)
        return False

    size = output_p.stat().st_size
    if size < 1000:
        print(
            f"[Attempt {attempt}] PDF file is suspiciously small ({size} bytes).",
            file=sys.stderr,
        )
        return False

    # Check PDF header magic bytes
    with open(output_p, "rb") as f:
        header = f.read(5)
    if header != b"%PDF-":
        print(
            f"[Attempt {attempt}] Output file is not a valid PDF (bad header).",
            file=sys.stderr,
        )
        return False

    # Success — even if LibreOffice exited with 134
    if result.returncode == 134:
        print(
            f"[Attempt {attempt}] LibreOffice exited with code 134 "
            "(known malloc_consolidate crash) but PDF is valid "
            f"({size:,} bytes)."
        )
    elif result.returncode != 0:
        print(
            f"[Attempt {attempt}] LibreOffice exited with code {result.returncode} "
            f"but PDF is valid ({size:,} bytes)."
        )
    else:
        print(f"[Attempt {attempt}] PDF generated successfully ({size:,} bytes).")

    return True


def main():
    if len(sys.argv) < 2 or len(sys.argv) > 3:
        print(f"Usage: {sys.argv[0]} <input.docx> [output.pdf]", file=sys.stderr)
        sys.exit(1)

    input_path = sys.argv[1]
    if not os.path.isfile(input_path):
        print(f"Error: input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    if len(sys.argv) == 3:
        output_path = sys.argv[2]
    else:
        output_path = str(Path(input_path).with_suffix(".pdf"))

    # Attempt 1
    if convert(input_path, output_path, attempt=1):
        sys.exit(0)

    # Attempt 2 (retry)
    print("Retrying conversion...", file=sys.stderr)
    if convert(input_path, output_path, attempt=2):
        sys.exit(0)

    print("PDF conversion failed after 2 attempts.", file=sys.stderr)
    sys.exit(1)


if __name__ == "__main__":
    main()
