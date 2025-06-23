# File Analyzer

A Streamlit web app for analyzing Excel, PDF, and Word documents with data visualization and AI-powered summaries.

## Prerequisites
- Python 3.8 or newer
- Install dependencies:
  ```bash
  pip install -r requirements.txt
  ```

## Setting the OpenAI API Key
The app uses OpenAI for generating summaries. Add your key to `.streamlit/secrets.toml`:

```toml
openai_api_key = "sk-..."
```

Alternatively, export the variable before running the app:

```bash
export OPENAI_API_KEY=sk-...
```

## Running the Streamlit App
Launch the interface with:

```bash
streamlit run streamlit_app.py
```

This command starts a local server and opens the app in your browser.

## Features
- Fetch Excel download links from a webpage
- Download and analyze multiple Excel files
- Upload Excel, PDF, or DOC/DOCX files for inspection
- Display statistics such as min and max values for chosen columns
- Plot histograms, correlation heatmaps, and other charts
- AI-generated summaries with a HuggingFace fallback when OpenAI isn't available
- Map visualization when latitude and longitude are present

## Example
Run the app and follow the prompts to analyze your files:

```bash
streamlit run streamlit_app.py
```

After processing, the app displays charts and an AI-generated summary of your data.
