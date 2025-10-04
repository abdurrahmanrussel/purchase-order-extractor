# extractor.py
import fitz  # PyMuPDF
import re
import zipfile
import os
import shutil
import pandas as pd
from pathlib import Path
from google.colab import files


# -----------------------------
# Functions
# -----------------------------
def extract_po_info(pdf_path: str) -> dict:
    """Extract purchase order header info from a single PDF."""
    try:
        doc = fitz.open(pdf_path)
        full_text = ""
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            text_blocks = page.get_text("blocks")
            text_blocks.sort(key=lambda block: (block[1], block[0]))
            for block in text_blocks:
                full_text += block[4] + "\n"
        doc.close()

        doc_number_match = re.search(r"\bPO\d+\b", full_text)
        document_number = doc_number_match.group(0) if doc_number_match else ""

        reference_match = re.search(r"Your Reference\s+([\w\-\/]*)", full_text)
        reference = reference_match.group(1).strip() if reference_match else ""
        if reference.lower() == "your":
            reference = ""

        payment_match = re.search(r"Payment Term:\s*(.+)", full_text)
        payment_term = payment_match.group(1).strip() if payment_match else ""

        # Document Date (PO_Issue Date)
        doc_date_match = re.search(r"\b\d{2}/\d{2}/\d{4}\b", full_text)
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
# Clean description logic
# -----------------------------
def clean_description(desc):
    desc = desc.strip()
    words = desc.split()
    for length in range(len(words)//2, 0, -1):  # check longest possible repeated substring
        start_sub = " ".join(words[:length])
        end_sub = " ".join(words[-length:])
        if start_sub.lower() == end_sub.lower():  # repeated substring detected
            middle = " ".join(words[length:-length])
            # If middle contains "Each" or numbers, keep longest repeated part
            if re.search(r'\bEach\b', middle, flags=re.IGNORECASE) or re.search(r'\d+\.\d+', middle):
                return max(start_sub, end_sub, key=len).strip()
    return desc


# -----------------------------
# Extract item blocks
# -----------------------------
def extract_item_blocks(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        full_text = ""
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            text_blocks = page.get_text("blocks")
            text_blocks.sort(key=lambda block: (block[1], block[0]))
            for block in text_blocks:
                full_text += block[4] + "\n"
        doc.close()

        # Remove commas in numbers
        full_text = re.sub(r'(?<=\d),(?=\d)', '', full_text)
        lines = full_text.splitlines()
        stop_marker = "▌Tax Details"

        inside_table = False
        blocks = []
        current_block = []
        waiting_for_date = False
        prev_line = ""  # for multi-line start marker
        start_marker_found = False

        for line in lines:
            line_strip = line.strip()

            # --------------------
            # Find start marker
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
                if re.match(r"\d{1,2}/\d{1,2}/\d{4}", line_strip):
                    current_block.append(line_strip)
                    # If block contains "page", ignore and restart searching for start marker
                    if any(re.search(r'page', l, re.IGNORECASE) for l in current_block):
                        current_block = []
                        waiting_for_date = False
                        inside_table = False
                        prev_line = ""
                        continue
                    # Otherwise, parse normally
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
# Parse block and clean description
# -----------------------------
def parse_block(block_lines):
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

    # Delivery Date
    for i, line in enumerate(block_lines):
        if line.strip().startswith("Delivery Date:") and i + 1 < len(block_lines):
            data["Delivery Date"] = block_lines[i+1].strip()
            break

    # Item Code
    item_code_index = -1
    for i, line in enumerate(block_lines):
        if line.strip().startswith("Item Code:") and i + 1 < len(block_lines):
            data["Item_Code"] = block_lines[i+1].strip()
            item_code_index = i+1
            break

    # Price & Total
    float_lines = [line.strip() for line in block_lines if re.match(r"^\d+\.\d+$", line.strip()) and line.strip() != "0.0000"]
    if float_lines:
        data["Price"] = float_lines[0]
        data["Total"] = max(float_lines, key=lambda x: float(x))

    # Quantity
    try:
        if data["Price"] and data["Total"]:
            data["Quantity"] = str(round(float(data["Total"]) / float(data["Price"]), 4))
    except:
        data["Quantity"] = ""

    # Detect UoM
    for line in block_lines:
        if line.strip() == "Each":
            data["UoM(optional)"] = "Each"
            break

    # Item Details (new logic)
    rev_index = -1
    for i, line in enumerate(block_lines):
        line_strip = line.strip()
        if re.match(r"^(Rev|REV)\s+\w+", line_strip):
            data["Item Details"] = line_strip
            rev_index = i
            break

    # Description (cleaned)
    if item_code_index != -1:
        if rev_index != -1:
            desc_lines = block_lines[rev_index+1:item_code_index]
        else:
            int_index = -1
            for j in range(item_code_index-1, -1, -1):
                if re.match(r"^\d+$", block_lines[j].strip()):
                    int_index = j
                    break
            if int_index != -1:
                desc_lines = block_lines[int_index+1:item_code_index]
            else:
                desc_lines = block_lines[:item_code_index]
        desc_lines = [l for l in desc_lines if not l.strip().startswith("Item Code:")]
        combined_desc = " ".join([l.strip() for l in desc_lines if l.strip()])
        data["Description"] = clean_description(combined_desc)

    return data


# -----------------------------
# 2️⃣ Clear previous uploads / folders
# -----------------------------
if os.path.exists("pdf_folder"):
    shutil.rmtree("pdf_folder")
os.makedirs("pdf_folder", exist_ok=True)

pdf_files = []
seen_filenames = set()


# -----------------------------
# 3️⃣ Upload ZIP or single PDF
# -----------------------------
uploaded = files.upload()

for file_name in uploaded.keys():
    if file_name.lower().endswith(".zip"):
        print(f"Processing ZIP: {file_name}")
        with zipfile.ZipFile(file_name, 'r') as zip_ref:
            zip_ref.extractall("pdf_folder")
    elif file_name.lower().endswith(".pdf"):
        if not file_name.startswith("._"):
            print(f"Processing single PDF: {file_name}")
            if file_name not in seen_filenames:
                pdf_files.append(file_name)
                seen_filenames.add(file_name)

# Recursively add PDFs from extracted ZIP folder
for root, dirs, files_in_dir in os.walk("pdf_folder"):
    for f in files_in_dir:
        if f.lower().endswith(".pdf") and not f.startswith("._") and "__MACOSX" not in root:
            full_path = os.path.join(root, f)
            if full_path not in seen_filenames:
                pdf_files.append(full_path)
                seen_filenames.add(full_path)


# -----------------------------
# 4️⃣ Process all PDFs
# -----------------------------
all_data = []

for pdf_file in pdf_files:
    po_info = extract_po_info(pdf_file)
    item_blocks = extract_item_blocks(pdf_file)

    if isinstance(item_blocks, dict) and "error" in item_blocks:
        print(item_blocks["error"])
        continue
    elif not isinstance(item_blocks, list):
        continue

    for block in item_blocks:
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


# -----------------------------
# 5️⃣ Create DataFrame if data exists
# -----------------------------
if all_data:
    df = pd.DataFrame(all_data)

    expected_cols = [
        "Document Number","Reference","Item_Code",
        "PO_issue Date","Delivery Date",
        "Description","Item Details","UoM(optional)",
        "Quantity","Price","Total","Payment Term"
    ]

    df = df.dropna(subset=["PO_issue Date"])
    existing_cols = [c for c in expected_cols if c in df.columns]
    df = df[existing_cols]

    if "Document Number" in df.columns:
        df["Document Number Sort"] = df["Document Number"].str.extract(r"(\d+)").astype(int)
        df = df.sort_values(by="Document Number Sort").drop(columns=["Document Number Sort"])

    # -----------------------------
    # 6️⃣ Download CSV
    # -----------------------------
    if not df.empty:
        df.to_csv("Extracted.csv", index=False)
        files.download("Extracted.csv")
    else:
        print("⚠️ No valid data extracted, CSV not created.")
else:
    print("⚠️ No data extracted from any PDF.")
