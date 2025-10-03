# extractor.py
# Standalone PDF PO extractor (Streamlit-friendly)
# - Scans 'input_pdf_folder' for PDFs
# - Extracts PO header info and item blocks using improved parsing logic
# - Saves output directly to 'Extracted.csv' (overwrites if exists)
#
# Requirements:
#   pip install PyMuPDF pandas

import fitz  # PyMuPDF
import re
import os
import shutil
import pandas as pd
from pathlib import Path

# -----------------------------
# Helper: clean_description
# -----------------------------
def clean_description(desc: str) -> str:
    """
    Trim and remove repeated prefix/suffix patterns.
    If the start and end repeated and the middle contains 'Each' or floats,
    prefer the longer repeated substring.
    """
    desc = (desc or "").strip()
    if not desc:
        return ""
    words = desc.split()
    # Check for repeated substring (try longer first)
    for length in range(len(words)//2, 0, -1):
        start_sub = " ".join(words[:length])
        end_sub = " ".join(words[-length:])
        if start_sub.lower() == end_sub.lower():
            middle = " ".join(words[length:-length])
            if re.search(r'\bEach\b', middle, flags=re.IGNORECASE) or re.search(r'\d+\.\d+', middle):
                return max(start_sub, end_sub, key=len).strip()
    return desc

# -----------------------------
# Header extraction
# -----------------------------
def extract_po_info(pdf_path: str) -> dict:
    """Extract purchase order header info from a single PDF."""
    try:
        doc = fitz.open(pdf_path)
        full_text = ""
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            text_blocks = page.get_text("blocks")
            # sort by y (top) then x (left) for deterministic ordering
            text_blocks.sort(key=lambda block: (block[1], block[0]))
            for block in text_blocks:
                full_text += (block[4] or "") + "\n"
        doc.close()

        # Document Number (e.g., PO12345)
        doc_number_match = re.search(r"\bPO\d+\b", full_text, flags=re.IGNORECASE)
        document_number = doc_number_match.group(0) if doc_number_match else ""

        # Reference
        reference_match = re.search(r"Your Reference\s+([\w\-\./]*)", full_text, flags=re.IGNORECASE)
        reference = reference_match.group(1).strip() if reference_match else ""
        if reference.lower() == "your":
            reference = ""

        # Payment Term
        payment_match = re.search(r"Payment Term:\s*(.+)", full_text, flags=re.IGNORECASE)
        payment_term = payment_match.group(1).strip() if payment_match else ""

        # Document Date (first occurrence of dd/mm/yyyy)
        doc_date_match = re.search(r"\b\d{1,2}/\d{1,2}/\d{4}\b", full_text)
        document_date = doc_date_match.group(0) if doc_date_match else ""

        return {
            "Document Number": document_number,
            "Reference": reference,
            "Payment Term": payment_term,
            "Document Date": document_date,
        }

    except Exception as e:
        return {"error": str(e)}

# -----------------------------
# Parse one item block
# -----------------------------
def parse_block(block_lines: list) -> dict:
    """
    Parse an individual item block (list of stripped lines) and return a dictionary.
    """
    data = {
        "Item_Code": "",
        "Delivery Date": "",
        "Description": "",
        "Item Details": "",
        "UoM(optional)": "",
        "Quantity": "",
        "Price": "",
        "Total": ""
    }

    # Delivery Date (line after 'Delivery Date:')
    for i, line in enumerate(block_lines):
        if line.startswith("Delivery Date:") and i + 1 < len(block_lines):
            data["Delivery Date"] = block_lines[i+1].strip()
            break

    # Item Code (line after 'Item Code:')
    item_code_index = -1
    for i, line in enumerate(block_lines):
        if line.startswith("Item Code:") and i + 1 < len(block_lines):
            data["Item_Code"] = block_lines[i+1].strip()
            item_code_index = i + 1
            break

    # Price & Total: find float-like lines (e.g., 123.45), exclude "0.0000"
    float_lines = [ln for ln in block_lines if re.match(r"^\d+\.\d+$", ln) and ln != "0.0000"]
    if float_lines:
        data["Price"] = float_lines[0]
        # Total is the maximum numeric value
        try:
            data["Total"] = max(float_lines, key=lambda x: float(x))
        except Exception:
            data["Total"] = float_lines[-1]

    # Quantity calculation
    try:
        if data["Price"] and data["Total"]:
            qty = float(data["Total"]) / float(data["Price"])
            data["Quantity"] = str(round(qty, 4))
    except Exception:
        data["Quantity"] = ""

    # Detect UoM
    for ln in block_lines:
        if ln == "Each":
            data["UoM(optional)"] = "Each"
            break

    # Item Details: look for lines starting with Rev/REV
    rev_index = -1
    for i, ln in enumerate(block_lines):
        if re.match(r"^(Rev|REV)\s+\w+", ln):
            data["Item Details"] = ln
            rev_index = i
            break

    # Description: try to locate slice between detected indices
    if item_code_index != -1:
        if rev_index != -1:
            desc_lines = block_lines[rev_index+1:item_code_index]
        else:
            # find nearest integer-only line before item_code_index (like an index)
            int_index = -1
            for j in range(item_code_index-1, -1, -1):
                if re.match(r"^\d+$", block_lines[j].strip()):
                    int_index = j
                    break
            if int_index != -1:
                desc_lines = block_lines[int_index+1:item_code_index]
            else:
                desc_lines = block_lines[:item_code_index]

        # Remove 'Item Code:' if present and empty lines
        desc_lines = [l for l in desc_lines if not l.strip().startswith("Item Code:") and l.strip()]
        combined_desc = " ".join(desc_lines).strip()
        data["Description"] = clean_description(combined_desc)

    return data

# -----------------------------
# Extract item blocks (improved logic)
# -----------------------------
def extract_item_blocks(pdf_path: str):
    """
    Extract item details from PDF table blocks.
    This looks for a start marker like 'Tax %'/'Tax%Tax%' to identify table area,
    collects blocks until a stop marker, then parses by Delivery Date grouping.
    """
    try:
        doc = fitz.open(pdf_path)
        full_text = ""
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            text_blocks = page.get_text("blocks")
            text_blocks.sort(key=lambda block: (block[1], block[0]))
            for block in text_blocks:
                full_text += (block[4] or "") + "\n"
        doc.close()

        # Normalize numbers (remove thousands separators)
        full_text = re.sub(r'(?<=\d),(?=\d)', '', full_text)
        lines = full_text.splitlines()
        stop_marker = "▌Tax Details"

        inside_table = False
        blocks = []
        current_block = []
        waiting_for_date = False
        prev_line = ""  # carry previous line to detect split 'Tax %' markers
        start_marker_found = False

        for line in lines:
            line_strip = line.strip()

            # --------------------
            # Find start marker (handles 'Tax %' split across lines)
            # --------------------
            if not inside_table:
                combined = (prev_line + line_strip).replace(" ", "")
                if combined == "Tax%Tax%":
                    inside_table = True
                    prev_line = ""
                    start_marker_found = True
                    continue
                if line_strip == "Tax %":
                    prev_line = line_strip
                    continue
                prev_line = line_strip
                continue

            # --------------------
            # Stop marker
            # --------------------
            if line_strip.startswith(stop_marker):
                inside_table = False
                continue

            # --------------------
            # Collect block lines
            # --------------------
            current_block.append(line_strip)

            if line_strip.startswith("Delivery Date:"):
                waiting_for_date = True
                continue

            if waiting_for_date:
                # expecting a date like d/m/yyyy or dd/mm/yyyy
                if re.match(r"\d{1,2}/\d{1,2}/\d{4}", line_strip):
                    # attach the date line
                    current_block.append(line_strip)
                    # If block contains 'page' (page footer/header), skip and reset
                    if any(re.search(r'page', l, re.IGNORECASE) for l in current_block):
                        current_block = []
                        waiting_for_date = False
                        inside_table = False
                        prev_line = ""
                        continue
                    # Otherwise parse this block
                    data = parse_block(current_block)
                    blocks.append(data)
                    current_block = []
                    waiting_for_date = False

        if not start_marker_found:
            return {"error": f"❌ '{os.path.basename(pdf_path)}' cannot convert to CSV, maybe format mismatch."}

        return blocks

    except Exception as e:
        return {"error": str(e)}

# -----------------------------
# Main script
# -----------------------------
def main():
    input_pdf_folder = Path("input_pdf_folder")
    # Ensure input folder exists
    if not input_pdf_folder.exists():
        print(f"⚠️ Input folder '{input_pdf_folder}' does not exist. Create it and add PDFs.")
        return

    # Collect pdf files recursively
    pdf_files = []
    for root, _, files_in_dir in os.walk(input_pdf_folder):
        for f in files_in_dir:
            if f.lower().endswith(".pdf") and not f.startswith("._") and "__MACOSX" not in root:
                pdf_files.append(os.path.join(root, f))

    if not pdf_files:
        print("⚠️ No PDF files found in 'input_pdf_folder'.")
        return

    all_data = []
    for pdf_file in pdf_files:
        po_info = extract_po_info(pdf_file)
        if isinstance(po_info, dict) and "error" in po_info:
            print(f"Error reading header from '{os.path.basename(pdf_file)}': {po_info['error']}")
            continue

        item_blocks = extract_item_blocks(pdf_file)

        if isinstance(item_blocks, dict) and "error" in item_blocks:
            print(item_blocks["error"])
            continue
        elif not isinstance(item_blocks, list):
            print(f"⚠️ Unexpected item_blocks result for '{os.path.basename(pdf_file)}'. Skipping.")
            continue

        for block in item_blocks:
            # Skip items if PO issue date not found
            if not po_info.get("Document Date", ""):
                continue

            row = {
                "Document Number": po_info.get("Document Number", ""),
                "Reference": po_info.get("Reference", ""),
                "Item_Code": block.get("Item_Code", ""),
                "PO_issue Date": po_info.get("Document Date", ""),
                "Delivery Date": block.get("Delivery Date", ""),
                "Description": block.get("Description", ""),
                "Item Details": block.get("Item Details", ""),
                "UoM(optional)": block.get("UoM(optional)", ""),
                "Quantity": block.get("Quantity", ""),
                "Price": block.get("Price", ""),
                "Total": block.get("Total", ""),
                "Payment Term": po_info.get("Payment Term", "")
            }
            all_data.append(row)

    if all_data:
        df = pd.DataFrame(all_data)
        expected_cols = [
            "Document Number","Reference","Item_Code",
            "PO_issue Date","Delivery Date",
            "Description","Item Details","UoM(optional)",
            "Quantity","Price","Total","Payment Term"
        ]
        # Drop rows with missing PO_issue Date
        df = df.dropna(subset=["PO_issue Date"])
        existing_cols = [c for c in expected_cols if c in df.columns]
        df = df[existing_cols]

        # Sort by numeric part of Document Number if available
        if "Document Number" in df.columns:
            try:
                df["Document Number Sort"] = df["Document Number"].str.extract(r"(\d+)").astype(float)
                df = df.sort_values(by="Document Number Sort").drop(columns=["Document Number Sort"])
            except Exception:
                # If extraction fails, skip sorting
                pass

        out_file = Path("Extracted.csv")
        try:
            df.to_csv(out_file, index=False)
            print(f"✅ Extraction complete → {out_file.resolve()}")
        except Exception as e:
            print(f"❌ Failed to write CSV: {e}")
    else:
        print("⚠️ No data extracted.")

if __name__ == "__main__":
    main()
