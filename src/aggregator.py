# -*- coding: utf-8 -*-
"""
Created on Mon Jun  9 15:14:48 2025

@author: ShivakrishnaBoora
"""


import os
from dotenv import load_dotenv
from openai_client import call_openai_chat
from file_io import save_text_to_docx
from text_processing import chunked_clean_and_summarize
import re
# Load environment variables
load_dotenv()

# Folder containing chunked summaries
SOURCE_FOLDER = "debug_chunks"
# Output files
COMBINED_TEXT_FILE = "combined_chunks.txt"
FINAL_SUMMARY_DOCX = "final_summary.docx"

import nltk

# Approximate character limit per API call
MAX_CHARS = 3000  

def _split_meta_chunks(text: str, max_chars: int = MAX_CHARS) -> list[str]:
    """
    Split `text` into pieces ≤ max_chars, cutting on sentence boundaries.
    """
    sents = nltk.sent_tokenize(text)
    chunks, cur = [], ""
    for s in sents:
        if len(cur) + len(s) + 1 > max_chars:
            chunks.append(cur.strip())
            cur = s
        else:
            cur = f"{cur} {s}".strip()
    if cur:
        chunks.append(cur)
    return chunks


def read_chunk_summaries(folder: str) -> str:
    """
    Read all summary_chunk_*.txt in `folder`, sorted by filename,
    and concatenate them into one large string.
    """
    parts = []
    for fname in sorted(os.listdir(folder)):
        if fname.startswith("summary_chunk_") and fname.endswith(".txt"):
            path = os.path.join(folder, fname)
            with open(path, "r", encoding="utf-8") as f:
                parts.append(f.read().strip())
    combined = "\n\n".join(parts)
    # Optionally save the raw combined text for inspection
    with open(COMBINED_TEXT_FILE, "w", encoding="utf-8") as out:
        out.write(combined)
    return combined


def summarize_timeline(combined_text: str) -> str:
    """
    If combined_text is small, call the API directly.
    Otherwise, split into meta-chunks, summarize each, then concatenate.
    """
    def _call(chunk):
        prompt = (
            "You are an expert at summarizing technical call transcripts into accurate lines. "
            "Follow these rules:\n"
            " 1. Keep each original timestamp verbatim (e.g., “05:42”, “06:00 AM IST”). "
            "Do not convert or recalculate times.\n"
            " 2. Use Active Voice only.\n"
            " 3. Preserve all numeric values (e.g., “500+ users”) and all team/role mentions (e.g., “admin team”, “Cloud OPS Team”).\n"
            " 4. Write each line as “HH:MM Speaker Text” with minimal words, but do not drop critical context.\n"
            " 5. Do not invent or infer missing details; only condense what’s in the cleaned transcript.\n\n"
            "Example of correct format:\n"
            "05:41 XXXXX Was there any upgrade or change on the environment?\n"
            "05:41 XXXXX No.\n"
            "05:41 XXXXX When was the first issue reported?\n"
            "6. **If the same speaker speaks in consecutive timestamps, merge their content into one summary line**—"
            "use the earliest timestamp, and separate clauses with semicolons.\n\n"
            "Example:\n"
            "05:10 Alice Checked the server status; reported CPU at 80%.\n"
            "05:12 Alice Noted that memory usage remained stable.\n"
            "→ Becomes:\n"
            "05:10 Alice Checked the server status and reported CPU at 80%; noted memory usage remained stable.\n\n"
            "Now, summarize the CLEANED TRANSCRIPT below into that format:\n\n"
            "Now, given the CLEANED TRANSCRIPT below, produce a similarly concise summary.\n"
            "Return only lines of the form “HH:MM Speaker Text” without adding or removing time details.\n\n"
            " 1. Uses Active Voice only.\n"
            " 2. Matches the exact “HH:MM Speaker Text” format—no extra commentary.\n"
            " 3. Never remove any numeric values or team names; they must appear exactly as in the transcript.\n\n"
            
            f"{chunk}"
        )
        return call_openai_chat(prompt)

    if len(combined_text) <= MAX_CHARS:
        return _call(combined_text)

    # meta-chunk & summarize
    parts = []
    for mc in _split_meta_chunks(combined_text):
        parts.append(_call(mc))
    return "\n".join(parts)



def summarize_table(combined_text: str) -> str:
    """
    Always return exactly one Markdown table.
    If the text is too large, first condense it, then extract the table.
    """
    # 1) If it's too large, get a condensed version first:
    if len(combined_text) > MAX_CHARS:
        # condense into a single smaller text
        condense_prompt = (
            "You are an expert at condensing transcript summaries. "
            "Preserve timestamps, names, teams, numbers, and table-worthy items. "
            "Produce a shorter unified summary.\n\n"
            f"{combined_text}"
        )
        combined_text = call_openai_chat(condense_prompt)
        # now combined_text is small enough to extract the table from

    # 2) Now do exactly one table extraction
    table_prompt = (
        
        "You are an expert assistant. Please provide ONLY a single Markdown table with the columns: "
        "| Action Item | Responsible Team | Notes |"
        
        "No additional text, bullet points, or explanations should be included before or after the table."
        
        "Format it exactly as a Markdown table. Do not add anything else. Keep 2-5 points only and merge the multiple action points per owner."
        "Follwed by Free text as given below"
        "**Observations/Follow-Up**\n"
        "• The VM system processes initiated the restart as it was a planned maintenance \"Operating system service plan\". "
        "No users restarted it; the process did. 06:00 AM IST was the restart time.<br>"
        "• IAS service restart helped in application functionality.\n\n"
        "TRANSCRIPT/SUMMARY:\n"
        f"{combined_text}"
    )
    return call_openai_chat(table_prompt)


def aggregate_and_finalize(
    raw_transcript: str,
    debug: bool = False
    ) -> tuple[str, str]:
    """
    Full pipeline:
      1. Chunk clean & summarize → partial timeline summaries.
      2. Concatenate partial timeline summaries → full timeline.
      3. (Optional) Meta-summarize full timeline if too long.
      4. ONE table extraction pass on the *full* timeline.
    Returns (full_timeline, table_md).
    """
    # 1+2) chunked clean & summarize as before
    full_timeline = chunked_clean_and_summarize(raw_transcript, debug=debug)

    # 3) final merge for timeline if needed (already inside chunked fn)

    # 4) Table extraction *only once* on the full timeline
    table_md = summarize_table(full_timeline)

    return full_timeline, table_md

def extract_last_markdown_table(md_text: str) -> str:
    """
    Extract only the last contiguous Markdown table block from md_text.
    A Markdown table block consists of lines starting and ending with '|'.
    """
    lines = md_text.splitlines()
    last_table_start = None
    last_table_end = None

    for i in range(len(lines) - 1, -1, -1):
        if re.match(r"^\|.*\|$", lines[i].strip()):
            if last_table_end is None:
                last_table_end = i
            last_table_start = i
        elif last_table_end is not None:
            break  # stop when no longer in table lines

    if last_table_start is not None and last_table_end is not None:
        # Expand upwards for contiguous lines
        start = last_table_start
        while start > 0 and re.match(r"^\|.*\|$", lines[start - 1].strip()):
            start -= 1
        return "\n".join(lines[start : last_table_end + 1])
    else:
        return ""  # no table found

import re

def extract_single_markdown_table(md_text: str) -> str:
    """
    Extract exactly one contiguous Markdown table block from the input text.
    Returns the last found Markdown table block as a string, or empty string if none.
    """
    lines = md_text.splitlines()
    table_blocks = []
    current_block = []
    inside_table = False

    for line in lines:
        if re.match(r'^\|.*\|$', line.strip()):
            current_block.append(line)
            inside_table = True
        else:
            if inside_table:
                table_blocks.append(current_block)
                current_block = []
                inside_table = False
    # if ended inside a table block
    if inside_table and current_block:
        table_blocks.append(current_block)

    if table_blocks:
        # Return last table block joined by newline
        return '\n'.join(table_blocks[-1])
    else:
        return ""

def main():
    combined = read_chunk_summaries(SOURCE_FOLDER)
    timeline = summarize_timeline(combined)
    table_md = summarize_table(combined)
    final_output = f"{timeline}\n\n{table_md}"
    save_text_to_docx(final_output, FINAL_SUMMARY_DOCX)
    print(f"Final summary written to {FINAL_SUMMARY_DOCX}")

    
if __name__ == "__main__":
    main()
