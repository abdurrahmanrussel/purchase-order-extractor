# 📄 Purchase Order Extractor

Extract structured data from **Purchase Order PDFs** into a clean CSV file. Works via **CLI**, **local Streamlit**, or **online Streamlit**.

---

## ⚡ Usage

### 1. Online (Streamlit Cloud)
Visit: [https://industrial-laser-po-extractor.streamlit.app/](https://industrial-laser-po-extractor.streamlit.app/)  
Upload PDFs/ZIPs, reorder/remove columns, sort, search, select rows, and download CSV.

### 2. Offline Streamlit: Open the web app in your browser and use the sidebar controls as above.
```bash
git clone https://github.com/abdurrahmanrussel/purchase-order-extractor.git
cd purchase-order-extractor
pip install -r requirements.txt
streamlit run streamlit_app.py
```


3. Offline CLI
Place PDFs/folder (not Zip) in input_pdf_folder/ and run:

```bash
python extractor.py
The extracted CSV will be saved in output/.
```







Ask ChatGPT
