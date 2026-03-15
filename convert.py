"""
convert.py - Mass convert .docx documents in a directory to a single JSON file.

Usage:
    python convert.py [input_dir] [output_file]

Arguments:
    input_dir   Directory containing .docx files (default: current directory)
    output_file Path for the output JSON file (default: output.json)
"""

import argparse
import json
import os
import sys

from docx import Document


def extract_text_from_docx(filepath: str) -> str:
    """Extract plain text content from a .docx file.

    Args:
        filepath: Absolute or relative path to the .docx file.

    Returns:
        A single string with all paragraph text joined by newlines.
    """
    doc = Document(filepath)
    return "\n".join(paragraph.text for paragraph in doc.paragraphs)


def collect_docx_files(directory: str) -> list[str]:
    """Return a sorted list of .docx file paths in the given directory.

    Args:
        directory: Path to the directory to scan.

    Returns:
        Sorted list of absolute paths to .docx files.
    """
    if not os.path.isdir(directory):
        raise NotADirectoryError(f"'{directory}' is not a valid directory.")
    return sorted(
        os.path.join(directory, f)
        for f in os.listdir(directory)
        if f.lower().endswith(".docx")
    )


def convert_docs_to_json(input_dir: str, output_file: str) -> list[dict]:
    """Convert all .docx files in input_dir into a single JSON file.

    Each entry in the JSON array has the shape:
        {
            "filename": "<basename of the .docx file>",
            "content":  "<extracted plain-text content>"
        }

    Args:
        input_dir:   Directory containing .docx files.
        output_file: Destination path for the output JSON file.

    Returns:
        The list of document dictionaries that was written to the JSON file.

    Raises:
        NotADirectoryError: If input_dir is not a valid directory.
        FileNotFoundError:  If no .docx files are found in input_dir.
    """
    docx_files = collect_docx_files(input_dir)
    if not docx_files:
        raise FileNotFoundError(f"No .docx files found in '{input_dir}'.")

    documents = []
    for filepath in docx_files:
        filename = os.path.basename(filepath)
        content = extract_text_from_docx(filepath)
        documents.append({"filename": filename, "content": content})
        print(f"  Converted: {filename}")

    output_dir = os.path.dirname(os.path.abspath(output_file))
    os.makedirs(output_dir, exist_ok=True)

    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(documents, f, indent=2, ensure_ascii=False)

    return documents


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Mass convert .docx files in a directory to a single JSON file."
    )
    parser.add_argument(
        "input_dir",
        nargs="?",
        default=".",
        help="Directory containing .docx files (default: current directory)",
    )
    parser.add_argument(
        "output_file",
        nargs="?",
        default="output.json",
        help="Output JSON file path (default: output.json)",
    )
    args = parser.parse_args()

    print(f"Scanning '{args.input_dir}' for .docx files...")
    try:
        documents = convert_docs_to_json(args.input_dir, args.output_file)
        print(f"\nDone! {len(documents)} document(s) written to '{args.output_file}'.")
    except (NotADirectoryError, FileNotFoundError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
