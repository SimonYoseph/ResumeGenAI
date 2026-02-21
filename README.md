# ResumeGenAI

AI-powered ATS-style resume analysis app built with Streamlit.

Upload a resume, paste a job description, and get:
- ATS score breakdown
- matched and missing skills
- AI-generated fit insights and recommendations
- visual charts for quick review

## Features

- Supports resume upload in PDF and DOCX
- Extracts and parses resume text into key sections
- Compares resume content to job description keywords
- Generates AI feedback using Hugging Face Inference API
- Visualizes results with Plotly charts
- Exports analysis output (PDF generation support included)

## Tech Stack

- Python 3.11+
- Streamlit
- Hugging Face Hub Inference Client
- pdfplumber / python-docx
- Plotly / pandas

## Project Structure

```
ResumeGenAI/
├── app.py
├── requirements.txt
├── README.md
├── run.sh
├── run_history.json
└── path/
```

## Quick Start

From the project root:

```bash
cd /Users/simonyoseph/Documents/Projects/Gen-AI-Resume-Scanner/ResumeGenAI
./run.sh
```

`run.sh` will create `.venv` if needed, install dependencies, and launch Streamlit.

## Run Locally (macOS)

### 1) Open the project folder

```bash
cd /Users/simonyoseph/Documents/Projects/Gen-AI-Resume-Scanner/ResumeGenAI
```

### 2) Create and activate a virtual environment (recommended)

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3) Install dependencies

```bash
pip install -r requirements.txt
```

### 4) Start the Streamlit app

```bash
streamlit run app.py
```

Streamlit will print a local URL (usually `http://localhost:8501`). Open it in your browser.

## Hugging Face Token

You need a Hugging Face token with inference access.

1. Create a token at: https://huggingface.co/settings/tokens
2. Grant permission to call inference providers
3. Use the token in the app when prompted

Optional `.env` setup:

```bash
cp .env.example .env
```

Then add your token value in `.env` if your app version reads it from environment variables.

## Usage

1. Upload your resume (PDF/DOCX)
2. Paste the target job description
3. Provide your Hugging Face API token
4. Click the analysis button
5. Review score, gaps, and recommendations

## Troubleshooting

- If `streamlit` is not found, activate your virtual environment again.
- If package install fails, run `python3 -m pip install --upgrade pip` then retry.
- If AI analysis fails, verify your Hugging Face token permissions and quota.
- If parsing fails on a resume, try a text-based PDF (not scanned image-only PDF).

## License

This project is for educational and portfolio use unless you add a specific license file.

