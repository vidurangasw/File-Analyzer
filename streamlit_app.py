import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import os
import io
import matplotlib.pyplot as plt
import seaborn as sns
import docx
import PyPDF2
import openai
from transformers import pipeline

st.set_page_config(page_title="Excel & Document Analyzer", layout="wide")
st.title("📊 File Analyzer: Excel, PDF, DOC with Visualization & AI Insights")

# Optional: Add your OpenAI API key securely in Streamlit secrets
api_key = st.secrets.get("openai_api_key", os.getenv("OPENAI_API_KEY"))

# Initialize OpenAI client only if API key is available
client = openai.OpenAI(api_key=api_key) if api_key else None

# Load HuggingFace summarizer
hf_summarizer = pipeline("summarization", model="google/flan-t5-small")

def generate_ai_summary(text):
    if not text.strip():
        return "No input provided for summarization."
    try:
        if client:
            response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are a data analyst. Summarize the following data insights in natural language."},
                    {"role": "user", "content": text}
                ]
            )
            return response.choices[0].message.content
        else:
            raise Exception("OpenAI client not initialized.")
    except Exception as e:
        st.warning(f"Falling back to HuggingFace model due to: {e}")
        try:
            result = hf_summarizer(text[:1000])  # Truncate for token limits
            return result[0]['summary_text']
        except Exception as he:
            return f"HuggingFace summarization failed: {he}"

def summarize_graph_insights(df, col_list):
    summaries = []
    for col in col_list:
        vals = pd.to_numeric(df[col], errors='coerce').dropna()
        if not vals.empty:
            description = vals.describe()
            summaries.append(
                f"The distribution of **{col}** has a mean of {description['mean']:.2f}, a median of {vals.median():.2f}, "
                f"and standard deviation of {description['std']:.2f}. The minimum is {description['min']:.2f} and maximum is {description['max']:.2f}."
            )
    return generate_ai_summary("\n".join(summaries))

url = st.text_input("Enter a webpage URL with downloadable .xlsx files:", "https://www.epa.gov/lmop/project-and-landfill-data-state")

# Extract downloadable links
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

columns = st.text_input("Enter column names to analyze (comma-separated):", "Actual MW Generation, Rated MW Capacity")

uploaded_file = st.file_uploader("Upload Excel, PDF, or DOC/DOCX file", type=["xlsx", "pdf", "doc", "docx"])
selected_sheet_name = None

# Analyze Excel manually downloaded from web URL
def analyze_df(df, columns):
    col_list = [c.strip() for c in columns.split(',') if c.strip()]
    row = {}
    for col in col_list:
        if col in df.columns:
            clean = pd.to_numeric(df[col], errors='coerce').dropna()
            row[f"{col} Min"] = clean.min()
            row[f"{col} Max"] = clean.max()
    return row

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
    for path in filepaths:
        try:
            df = pd.read_excel(path, sheet_name=sheet)
            link = next((url for url in st.session_state["excel_links"] if os.path.basename(url) == os.path.basename(path)), None)
            display_link = f"[Open File]({link})" if link else os.path.basename(path)
            row = {"File": display_link}
            row.update(analyze_df(df, columns))
            results.append(row)
        except Exception as e:
            results.append({"File": os.path.basename(path), "Error": str(e)})
    return pd.DataFrame(results)

if st.button("Download & Analyze Excel Files") and "excel_links" in st.session_state:
    sheet_name = st.text_input("Enter the sheet name to analyze:", "LMOP Database")
    filepaths = download_excel_files(st.session_state["excel_links"])
    if filepaths:
        st.info(f"Downloaded {len(filepaths)} files. Starting analysis...")

        def analyze_files_with_links(filepaths, sheet, columns):
            results = []
            for path in filepaths:
                try:
                    df = pd.read_excel(path, sheet_name=sheet)
                    link = next((url for url in st.session_state["excel_links"] if os.path.basename(url) == os.path.basename(path)), None)
                    display_link = f'<a href="{link}" target="_blank">{os.path.basename(path)}</a>' if link else os.path.basename(path)
                    row = {"File": display_link}
                    row.update(analyze_df(df, columns))
                    results.append(row)
                except Exception as e:
                    results.append({"File": os.path.basename(path), "Error": str(e)})
            return pd.DataFrame(results)

        df_result = analyze_files_with_links(filepaths, sheet_name, columns)
        st.markdown("### 🔗 Analysis Results with File Links", unsafe_allow_html=True)
        st.write(df_result.to_html(escape=False, index=False), unsafe_allow_html=True)

        output = io.BytesIO()
        df_plain = df_result.copy()
        df_plain['File'] = df_plain['File'].apply(lambda x: BeautifulSoup(x, "html.parser").text if isinstance(x, str) else x)
        df_plain.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)
        st.download_button("📥 Download Result as Excel", data=output, file_name="Excel_Analysis_Summary.xlsx")


# Uploaded file section
if uploaded_file is not None:
    file_type = uploaded_file.name.split('.')[-1].lower()

    if file_type == "xlsx":
        try:
            excel_preview = pd.ExcelFile(uploaded_file)
            sheet_names = excel_preview.sheet_names
            selected_sheet_name = st.selectbox("Select sheet to analyze", sheet_names)

            df_uploaded = pd.read_excel(uploaded_file, sheet_name=selected_sheet_name)
            st.dataframe(df_uploaded.head())
            analysis = analyze_df(df_uploaded, columns)
            st.json(analysis)

            col_list = [c.strip() for c in columns.split(',') if c.strip() and c.strip() in df_uploaded.columns]
            for col in col_list:
                st.write(f"### Histogram for {col}")
                fig, ax = plt.subplots()
                sns.histplot(df_uploaded[col], bins=20, kde=True, ax=ax)
                st.pyplot(fig)

            if len(col_list) >= 2:
                st.write("### Correlation Heatmap")
                fig, ax = plt.subplots()
                sns.heatmap(df_uploaded[col_list].corr(), annot=True, cmap='coolwarm', ax=ax)
                st.pyplot(fig)

                st.write("### Boxplot & Scatterplot")
                fig, axes = plt.subplots(1, 2, figsize=(14, 5))
                sns.boxplot(data=df_uploaded[col_list], ax=axes[0])
                sns.scatterplot(x=col_list[0], y=col_list[1], data=df_uploaded, ax=axes[1])
                st.pyplot(fig)

            # AI summary
            # AI summary
            if not df_uploaded.empty:
                st.write("### AI Summary")
                summary = []
                for col in col_list:
                    vals = pd.to_numeric(df_uploaded[col], errors='coerce').dropna()
                    if not vals.empty:
                        summary.append(f"Column {col}: mean={vals.mean():.2f}, median={vals.median():.2f}, std={vals.std():.2f}")
                ai_response = generate_ai_summary("\n".join(summary))
                st.markdown("#### 🤖 AI-Powered Summary:")
                st.success(ai_response)

                st.write("### 📉 AI-Powered Graph Summary")
                graph_insight = summarize_graph_insights(df_uploaded, col_list)
                st.info(graph_insight)

            # Map visualization if coordinates available
            if 'Latitude' in df_uploaded.columns and 'Longitude' in df_uploaded.columns:
                try:
                    lat_lon_df = df_uploaded[['Latitude', 'Longitude']].dropna()
                    lat_lon_df = lat_lon_df.rename(columns={"Latitude": "lat", "Longitude": "lon"})
                    st.write("### 🗺️ Map of Locations")
                    st.map(lat_lon_df)
                except Exception as e:
                    st.warning(f"Could not display map: {e}")

        except Exception as e:
            st.error(f"Error reading Excel file: {e}")

    elif file_type == "pdf":
        reader = PyPDF2.PdfReader(uploaded_file)
        text = "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])
        st.text_area("PDF Text Content", text, height=300)

    elif file_type in ["doc", "docx"]:
        doc = docx.Document(uploaded_file)
        text = "\n".join([para.text for para in doc.paragraphs if para.text.strip()])
        st.text_area("Word Document Content", text, height=300)
