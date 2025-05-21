import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import os
import io

st.set_page_config(page_title="Excel File Analyzer", layout="wide")
st.title("📊 Excel File Analyzer from Web URL")

# Step 1: Get the user input URL
url = st.text_input("Enter a webpage URL with downloadable .xlsx files:", "https://www.epa.gov/lmop/project-and-landfill-data-state")

# Step 2: Extract links from the URL
def get_download_links(base_url, file_types=[".xlsx"]):
    try:
        response = requests.get(base_url)
        soup = BeautifulSoup(response.content, "html.parser")
        links = soup.find_all("a", href=True)
        return [urljoin(base_url, link['href']) for link in links if any(link['href'].lower().endswith(ft) for ft in file_types)]
    except Exception as e:
        st.error(f"Failed to extract links: {e}")
        return []

if st.button("Fetch Excel File Links"):
    file_links = get_download_links(url)
    if file_links:
        st.success(f"Found {len(file_links)} Excel files.")
        st.write(file_links)
        st.session_state["excel_links"] = file_links
    else:
        st.warning("No Excel files found.")

# Step 3: Sheet and Column Input
sheet_name = st.text_input("Enter the sheet name to analyze:", "LMOP Database")
columns = st.text_input("Enter column names to analyze (comma-separated):", "Actual MW Generation, Rated MW Capacity")

# Step 4: Analyze the Excel files
def download_excel_files(links):
    os.makedirs("downloads", exist_ok=True)
    local_paths = []
    for link in links:
        try:
            filename = os.path.basename(link)
            local_path = os.path.join("downloads", filename)
            r = requests.get(link)
            with open(local_path, 'wb') as f:
                f.write(r.content)
            local_paths.append(local_path)
        except Exception as e:
            st.warning(f"Failed to download {link}: {e}")
    return local_paths

def analyze_files(filepaths, sheet, columns):
    results = []
    col_list = [c.strip() for c in columns.split(',') if c.strip()]
    for path in filepaths:
        try:
            df = pd.read_excel(path, sheet_name=sheet)
            row = {"File": os.path.basename(path)}
            for col in col_list:
                if col in df.columns:
                    clean = pd.to_numeric(df[col], errors='coerce').dropna()
                    row[f"{col} Min"] = clean.min()
                    row[f"{col} Max"] = clean.max()
                else:
                    row[f"{col} Min"] = "Not found"
                    row[f"{col} Max"] = "Not found"
            results.append(row)
        except Exception as e:
            results.append({"File": os.path.basename(path), "Error": str(e)})
    return pd.DataFrame(results)

if st.button("Download & Analyze Excel Files") and "excel_links" in st.session_state:
    filepaths = download_excel_files(st.session_state["excel_links"])
    if filepaths:
        st.info(f"Downloaded {len(filepaths)} files. Starting analysis...")
        df_result = analyze_files(filepaths, sheet_name, columns)
        st.dataframe(df_result)

        # Prepare Excel download
        output = io.BytesIO()
        df_result.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        st.download_button(
            label="📥 Download Result as Excel",
            data=output,
            file_name="Excel_Analysis_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
