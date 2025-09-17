# Transcript Summarizer & QA Generator

A Streamlit application that:
1. “Pulls” a Microsoft Teams call transcript from a local `.docx` (as if from a source),
2. Cleans and summarizes it using OpenAI’s GPT‐4-mini,
3. Displays both original and summarized text side by side, and
4. Allows comparison between two `.docx` transcripts—computing BLEU, WER, Levenshtein, and semantic‐embedding similarity.

## Folder Structure

transcript_summarizer/
├── .gitignore
├── README.md
├── requirements.txt
├── .env
└── src/
├── init.py
├── file_io.py
├── openai_client.py
├── text_processing.py
├── comparators.py
└── app.py


## Setup

1. Clone this repo.
2. Copy `.env.example` to `.env` and fill in your `OPENAI_API_KEY`.
3. Install dependencies:

   ```bash
   pip install -r requirements.txt
   python -m nltk.downloader punkt

