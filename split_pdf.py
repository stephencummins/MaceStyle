#!/usr/bin/env python3
"""
Split a PDF file in half.
Creates two output files: one with the first half of pages, one with the second half.
"""

import sys
from pathlib import Path

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    print("Error: pypdf library not found. Install it with: pip install pypdf")
    sys.exit(1)


def split_pdf_in_half(input_path, output_dir=None):
    """
    Split a PDF file into two halves.

    Args:
        input_path: Path to the input PDF file
        output_dir: Directory for output files (defaults to same as input)

    Returns:
        Tuple of (first_half_path, second_half_path)
    """
    input_path = Path(input_path)

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    if not input_path.suffix.lower() == '.pdf':
        raise ValueError("Input file must be a PDF")

    # Set output directory
    if output_dir is None:
        output_dir = input_path.parent
    else:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

    # Read the PDF
    reader = PdfReader(input_path)
    total_pages = len(reader.pages)

    if total_pages < 2:
        raise ValueError("PDF must have at least 2 pages to split")

    # Calculate split point
    midpoint = total_pages // 2

    # Create output filenames
    base_name = input_path.stem
    first_half_path = output_dir / f"{base_name}_part1.pdf"
    second_half_path = output_dir / f"{base_name}_part2.pdf"

    # Create first half
    writer1 = PdfWriter()
    for i in range(midpoint):
        writer1.add_page(reader.pages[i])

    with open(first_half_path, 'wb') as f:
        writer1.write(f)

    # Create second half
    writer2 = PdfWriter()
    for i in range(midpoint, total_pages):
        writer2.add_page(reader.pages[i])

    with open(second_half_path, 'wb') as f:
        writer2.write(f)

    return first_half_path, second_half_path, midpoint, total_pages


def main():
    if len(sys.argv) < 2:
        print("Usage: python split_pdf.py <input_pdf> [output_directory]")
        print("\nExample:")
        print("  python split_pdf.py document.pdf")
        print("  python split_pdf.py document.pdf ./output")
        sys.exit(1)

    input_file = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None

    try:
        first_half, second_half, midpoint, total = split_pdf_in_half(input_file, output_dir)
        print(f"Successfully split PDF ({total} pages total):")
        print(f"  First half (pages 1-{midpoint}): {first_half}")
        print(f"  Second half (pages {midpoint + 1}-{total}): {second_half}")
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
