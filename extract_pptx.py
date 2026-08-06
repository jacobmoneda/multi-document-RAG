"""
Extracts text, speaker notes, and table content from .pptx lecture files
into a structured JSON format ready for chunking/embedding in a RAG pipeline.

Usage:
    python extract_pptx.py <input_folder> <output_json>

Example:
    python extract_pptx.py data/raw/ data/extracted_slides.json
"""

import sys
import json
from pathlib import Path
from pptx import Presentation


def extract_shape_text(shape):
    """Extract text from a single shape (text box, title, placeholder, etc.)."""
    if not shape.has_text_frame:
        return ""
    lines = []
    for para in shape.text_frame.paragraphs:
        line = "".join(run.text for run in para.runs)
        if line.strip():
            lines.append(line)
    return "\n".join(lines)


def extract_table_text(shape):
    """Extract text from a table shape as pipe-separated rows."""
    rows_text = []
    table = shape.table
    for row in table.rows:
        cells = [cell.text.strip() for cell in row.cells]
        rows_text.append(" | ".join(cells))
    return "\n".join(rows_text)


def extract_slide(slide, slide_num, source_file, course):
    """Pull all text content + metadata from a single slide."""
    text_parts = []
    has_image_only_content = False

    for shape in slide.shapes:
        if shape.has_text_frame:
            text = extract_shape_text(shape)
            if text:
                text_parts.append(text)
        elif shape.has_table:
            text_parts.append(extract_table_text(shape))
        elif shape.shape_type == 13:  # MSO_SHAPE_TYPE.PICTURE
            has_image_only_content = True

    slide_text = "\n".join(text_parts).strip()

    # Speaker notes, if present
    notes_text = ""
    if slide.has_notes_slide:
        notes_frame = slide.notes_slide.notes_text_frame
        notes_text = notes_frame.text.strip() if notes_frame else ""

    return {
        "text": slide_text,
        "notes": notes_text,
        "source_file": source_file,
        "slide_number": slide_num,
        "course": course,
        "chunk_id": f"{Path(source_file).stem}_s{slide_num}".lower().replace(" ", "_"),
        "has_image_only_content": has_image_only_content,
    }


def extract_pptx_file(filepath, course=None):
    """Extract all slides from one .pptx file."""
    prs = Presentation(filepath)
    source_file = Path(filepath).name
    course = course or Path(filepath).stem.split("_")[0]  # e.g. "COMP717" from "COMP717_Lecture4.pptx"

    slides_data = []
    for i, slide in enumerate(prs.slides, start=1):
        slide_data = extract_slide(slide, i, source_file, course)
        # Skip completely empty slides (e.g. section dividers with no text)
        if slide_data["text"] or slide_data["notes"]:
            slides_data.append(slide_data)

    return slides_data


def extract_folder(folder_path, output_path):
    """Walk a folder of .pptx files and dump all extracted slides to JSON."""
    folder = Path(folder_path)
    pptx_files = sorted(folder.glob("*.pptx"))

    if not pptx_files:
        print(f"No .pptx files found in {folder_path}")
        return

    all_slides = []
    for pptx_file in pptx_files:
        print(f"Extracting: {pptx_file.name}")
        try:
            slides = extract_pptx_file(pptx_file)
            all_slides.extend(slides)
            print(f"  -> {len(slides)} slides extracted")
        except Exception as e:
            print(f"  !! Failed to extract {pptx_file.name}: {e}")

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(all_slides, f, indent=2, ensure_ascii=False)

    print(f"\nDone. {len(all_slides)} total slides written to {output_path}")

    # Quick sanity flags worth checking manually
    image_only = [s for s in all_slides if s["has_image_only_content"] and not s["text"]]
    if image_only:
        print(f"Note: {len(image_only)} slides have image-only content with no extracted text "
              f"(diagrams/screenshots) — these won't be searchable by text retrieval yet.")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python extract_pptx.py <input_folder> <output_json>")
        sys.exit(1)

    input_folder = sys.argv[1]
    output_json = sys.argv[2]
    extract_folder(input_folder, output_json)