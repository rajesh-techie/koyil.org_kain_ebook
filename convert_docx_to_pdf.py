#!/usr/bin/env python3
"""
Convert DOCX to PDF using Microsoft Word.

No extra Python packages required - uses cscript.exe (built into Windows)
to drive Word's COM automation.

Usage:
    python convert_docx_to_pdf.py <input.docx> [output.pdf] [--open]

Examples:
    python convert_docx_to_pdf.py Thiruviruththam_FINAL.docx
    python convert_docx_to_pdf.py Thiruviruththam_FINAL.docx output.pdf
    python convert_docx_to_pdf.py Thiruviruththam_FINAL.docx output.pdf --open
"""

import sys
import subprocess
import os
import tempfile
from pathlib import Path
from datetime import datetime


def convert_docx_to_pdf(input_path, output_path=None, open_pdf=False):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Validate and resolve input
    input_path = Path(input_path).resolve()
    if not input_path.exists():
        print(f"[{timestamp}] Error: Input file does not exist: {input_path}")
        return False
    if input_path.suffix.lower() != ".docx":
        print(f"[{timestamp}] Error: Input must be a .docx file, got: {input_path.suffix}")
        return False

    # Resolve output path
    if output_path is None:
        output_path = input_path.with_suffix(".pdf")
    else:
        output_path = Path(output_path).resolve()

    output_path.parent.mkdir(parents=True, exist_ok=True)

    input_mb = round(input_path.stat().st_size / (1024 * 1024), 2)
    print(f"[{timestamp}] Input:  {input_path} ({input_mb} MB)")
    print(f"[{timestamp}] Output: {output_path}")

    # Build a temporary VBScript that drives Word COM (no extra packages needed)
    vbs = f"""
Dim word, doc
Set word = CreateObject("Word.Application")
word.Visible = False
Set doc = word.Documents.Open("{str(input_path)}", False, False, False)
doc.SaveAs "{str(output_path)}", 17
doc.Close False
word.Quit
WScript.Quit 0
"""

    vbs_file = None
    try:
        # Write VBScript to a temp file
        fd, vbs_path = tempfile.mkstemp(suffix=".vbs")
        with os.fdopen(fd, "w") as f:
            f.write(vbs)
        vbs_file = vbs_path

        print(f"[{timestamp}] Converting with Microsoft Word...")
        start = datetime.now()

        result = subprocess.run(
            ["cscript.exe", "//NoLogo", vbs_path],
            capture_output=True,
            text=True,
            timeout=300
        )

        if result.returncode != 0:
            print(f"[{timestamp}] Error: {result.stderr.strip() or result.stdout.strip()}")
            return False

        if not output_path.exists():
            print(f"[{timestamp}] Error: PDF was not created")
            return False

        duration = (datetime.now() - start).total_seconds()
        output_mb = round(output_path.stat().st_size / (1024 * 1024), 2)
        print(f"[{timestamp}] Done: {output_path} ({output_mb} MB) in {duration:.1f}s")

        if open_pdf:
            os.startfile(str(output_path))

        return True

    except subprocess.TimeoutExpired:
        print(f"[{timestamp}] Error: Conversion timed out")
        return False
    except FileNotFoundError:
        print(f"[{timestamp}] Error: cscript.exe not found - are you on Windows?")
        return False
    except Exception as e:
        print(f"[{timestamp}] Error: {e}")
        return False
    finally:
        if vbs_file and os.path.exists(vbs_file):
            os.remove(vbs_file)


def main():
    if len(sys.argv) < 2:
        print("Usage: python convert_docx_to_pdf.py <input.docx> [output.pdf] [--open]")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = None
    open_pdf = False

    for arg in sys.argv[2:]:
        if arg == "--open":
            open_pdf = True
        else:
            output_path = arg

    success = convert_docx_to_pdf(input_path, output_path, open_pdf)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
