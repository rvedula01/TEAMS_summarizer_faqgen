# Meeting Summarizer & Knowledge Base Generator

A Streamlit application that transforms meeting transcripts and chat logs into structured, actionable summaries and knowledge bases. The application:

1. **Ingests Multiple Inputs**:
   - Microsoft Teams meeting transcripts (`.docx`)
   - Chat logs from meetings (`.docx` or `.txt`)
   - Supports single or multiple file uploads

2. **Data Processing**:
   - Merges transcript and chat data chronologically
   - Cleans and normalizes text while preserving technical details
   - Extracts key information including action items, decisions, and technical discussions

3. **AI-Powered Analysis**:
   - Uses OpenAI's GPT-4o for intelligent summarization
   - Extracts technical Q&A pairs for knowledge base creation
   - Identifies action items with responsible parties
   - Highlights key observations and decisions

4. **Output Generation**:
   - Generates comprehensive meeting summaries
   - Creates structured FAQ documents
   - Produces action item trackers
   - Exports to multiple formats (DOCX, PDF)

## Key Features

### 1. Transcript & Chat Merger
- Combines meeting transcripts and chat logs into a single timeline
- Preserves chronological order of all discussions
- Maintains speaker attribution and timestamps

### 2. Intelligent Summarization
- Generates concise, action-oriented summaries
- Focuses on technical content and decisions
- Preserves important context and details

### 3. Knowledge Extraction
- Automatically identifies and extracts technical Q&A pairs
- Filters out non-technical content and casual conversation
- Creates structured knowledge base entries

### 4. Action Item Tracking
- Identifies action items with clear ownership
- Tracks status and due dates
- Highlights priority items

## Folder Structure
meeting_summarizer/ ├── .gitignore ├── README.md ├── requirements.txt ├── .env └── src/ ├── init.py ├── app.py # Main Streamlit application ├── openai_client.py # OpenAI API integration ├── text_processing.py # Text cleaning and processing └── file_io.py # File handling utilities


## Setup

1. **Clone the repository**:
   ```bash
   git clone [https://github.com/yourusername/meeting-summarizer.git](https://github.com/yourusername/meeting-summarizer.git)
   cd meeting-summarizer

2. **Install dependencies**:
   pip install -r requirements.txt