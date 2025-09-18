# -*- coding: utf-8 -*-
"""
Created on Fri Jun  6 15:17:13 2025

@author: ShivakrishnaBoora
"""


from typing import Any
from typing import List, Generator
import nltk
import os

# Ensure all required NLTK data is downloaded
try:
    nltk.data.find('tokenizers/punkt')
except LookupError:
    nltk.download('punkt', quiet=True)

try:
    nltk.data.find('tokenizers/punkt_tab')
except LookupError:
    nltk.download('punkt_tab', quiet=True)

from openai_client import call_openai_chat

def clean_transcript(raw_transcript: str) -> str:
    """
    Clean the transcript by:
        1. Pre-processing the input to handle timestamp and speaker format
        2. Maintaining original question structure and tone
        3. Removing unnecessary filler words and duplicate words
        4. Handling questions and confirmations properly
        5. Preserving numeric values and team mentions
        6. Fixing grammar and punctuation
    """
    # Pre-process the input to handle the timestamp and speaker format
    lines = raw_transcript.split('\n')
    processed_lines = []
    i = 0
    
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue
            
        # Check if line starts with a date (MM/DD/YYYY or YYYY-MM-DD)
        if any(line.startswith(fmt) for fmt in ['01/', '02/', '03/', '04/', '05/', '06/', '07/', '08/', '09/', '10/', '11/', '12/']):
            # This is a timestamp line, next line is the speaker and content
            if i + 1 < len(lines):
                timestamp = line
                speaker_content = lines[i+1].strip()
                # Split speaker and content (first word is speaker, rest is content)
                parts = speaker_content.split(' ', 1)
                if len(parts) == 2:
                    speaker, content = parts
                    # Format as MM:SS Speaker Content
                    processed_line = f"{timestamp.split(', ')[1]} {speaker} {content}"
                    processed_lines.append(processed_line)
                i += 2  # Skip the next line as we've processed it
                continue
        i += 1
    
    # Join the processed lines with newlines
    pre_processed = '\n\n'.join(processed_lines)
    
    prompt = (
        "You are a professional transcript cleaner. Follow these rules EXACTLY:\n\n"
        "## INPUT FORMAT:\n"
        "- Each line contains: Timestamp Speaker Content\n\n"
        "## CRITICAL OUTPUT FORMAT REQUIREMENTS:\n"
        "- Each line MUST follow this EXACT format: Timestamp Speaker Content\n"
        "- Preserve the original timestamp format exactly as provided\n"
        "- Use EXACTLY ONE space between timestamp and speaker name\n"
        "- Use EXACTLY ONE space between speaker name and content\n"
        "- Add EXACTLY ONE blank line between each entry\n"
        "- NO headers, footers, explanations, or additional text\n"
        "- DO NOT end any line with a speaker name\n\n"
        "## TEXT CLEANING RULES:\n"
        "1. Remove duplicate words or phrases (e.g., 'but but but' → 'but')\n"
        "2. Remove filler words: um, uh, hmm, ah, you know, like, so\n"
        "3. Remove any single-word sentences, except 'Yes' or 'No'\n"
        "4. Keep all original questions and confirmation phrases\n"
        "5. Correct grammar, capitalization, and punctuation\n"
        "6. Remove redundant or repeated information\n"
        "7. Keep all numbers and team/role mentions unchanged\n"
        "8. If a line contains '[IMAGE:...]', keep it exactly as is\n\n"
        "## EXAMPLES:\n\n"
        "Input:\n"
        "15:21:23 Aryan I'll do that. OK, OK. Tell me what is the issue? Neerav ABC\n"
        "15:21:27 Neerav ABC OK.\n\n"
        "Output:\n"
        "15:21:23 Aryan I'll do that. Please tell me what is the issue?\n\n"
        "15:21:27 Neerav ABC OK.\n\n"

        "***## MANDATORY RULE FOR IMAGE PLACEHOLDERS:\n***"
        "***- If a line contains an image placeholder like [IMAGE:temp_images\\image_x.png], DO NOT modify the content in any way.\n***"
        "***- Retain such lines EXACTLY as they appear, including speaker, timestamp, and image text.\n***"
        "***- DO NOT clean, merge, format, or edit these lines.Ad there should be only one line for each image placeholder\n ***"
        "***- Example: '03:15 Speaker content [IMAGE:temp_images\\image_x.png]' → keep exactly like that strictly without any modification.\n\n***"
        
        "Now clean this transcript:\n"
        f"{pre_processed}"
    )
    
    result = call_openai_chat(prompt)
    return result


def summarize_transcript(cleaned_transcript: str, include_header: bool = False) -> str:
    """
    Summarize a cleaned transcript into professional, active-voice statements
    with consistent formatting and proper grammar.
    """
    
    prompt = (
        "You are a professional transcript summarizer. Create summaries following these EXACT rules:\n\n"
        
        "## CRITICAL OUTPUT FORMAT REQUIREMENTS:\n"
        "- Each line MUST follow this EXACT format: MM:SS Speaker Content\n"
        "- Use EXACTLY ONE space between timestamp and speaker name\n"
        "- Use EXACTLY ONE space between speaker name and content\n"
        "- Add EXACTLY ONE blank line between each entry\n"
        "- NO headers, footers, explanations, or additional text\n"
        "- Preserve timestamp format exactly as provided (0:08, 02:44, 1:23:45, etc.)\n\n"
        
        "***## MANDATORY RULE FOR IMAGE PLACEHOLDERS:\n***"
        "***- If a line contains an image placeholder like [IMAGE:temp_images\\image_x.png], DO NOT modify the content in any way.\n***"
        "***- Retain such lines EXACTLY as they appear, including speaker, timestamp, and image text.\n***"
        "***- DO NOT clean, merge, format, or edit these lines.Ad there should be only one line for each image placeholder\n ***"
        "***- Example: '03:15 Speaker content [IMAGE:temp_images\\image_x.png]' → keep exactly like that strictly without any modification.\n\n***"
        
        "## CONTENT RULES:\n"
        "1. **Grammar and Voice:**\n"
        "   - Use ACTIVE VOICE exclusively\n"
        "   - Write complete, grammatically correct sentences\n"
        "   - Proper punctuation and capitalization\n"
        "   - Example: 'System was restarted' → 'Team restarted the system'\n\n"
        
        "2. **Consecutive speaker handling:**\n"
        "   - Merge consecutive same-speaker entries\n"
        "   - Use EARLIEST timestamp from the sequence\n"
        "   - Separate merged content with semicolons\n"
        "   - Example: 'Chin 0:16 First point; Second point; Third point'\n\n"
        
        "3. **Content preservation:**\n"
        "   - Keep ALL numeric values exactly (500+, P1, P2, server names)\n"
        "   - Preserve team/role mentions exactly (Admin team, Cloud OPS Team)\n"
        "   - Maintain technical terminology and identifiers\n"
        "   - Keep important process steps and troubleshooting actions\n\n"
        
        "4. **Semantic accuracy:**\n"
        "   - Fix illogical statements\n"
        "   - Correct country/region references\n"
        "   - Clarify unclear technical references\n"
        "   - Example: 'all countries in India' → 'all regions in India'\n\n"
        
        "5. **Professional summarization:**\n"
        "   - Convert lengthy explanations into clear statements\n"
        "   - Remove redundant information while preserving key details\n"
        "   - Maintain chronological flow of events\n"
        "   - Transform unclear questions into professional inquiries\n\n"
        
        "## MANDATORY OUTPUT FORMAT:\n"
        "MM:SS Speaker Content\n"
        "\n"
        "MM:SS Speaker Content\n"
        "\n"
        "MM:SS Speaker Content\n\n"
        
        "## EXAMPLE OUTPUT:\n"
        "0:00 Speaker1 Team identified API latency spike; logs show increased Server A load.\n"
        "\n"
        "0:08 Speaker2 We restarted backend service to reduce response times.\n"
        "\n"
        "1:23:45 Speaker3 Can you confirm cache node impact status?\n\n"
        
        "Summarize this transcript:\n"
        f"{cleaned_transcript}"
    )
    
    summarized = call_openai_chat(prompt)
    # print("summarized",summarized)
    summarized = summarized.replace("```", "")
    summarized = summarized.replace("plaintext", "")
    summarized = summarized.replace("PLAINTEXT", "")
    summarized = summarized.replace("plaintex", "")
    
    # if include_header:
    #     return f"**Timelines (Times are in Eastern time (GMT-5) unless otherwise noted):**\n\n{summarized}"
    return summarized

# Approximate max characters per chunk (~token limit proxy)
MAX_CHARS_PER_CHUNK = 3000

def _split_raw_into_chunks(raw: str, max_chars: int = MAX_CHARS_PER_CHUNK) -> List[str]:
    """
    Split raw transcript into chunks under max_chars, cutting only at sentence boundaries.
    """
    sentences = nltk.sent_tokenize(raw)
    chunks, current = [], ""
    for sent in sentences:
        if len(current) + len(sent) + 1 > max_chars:
            chunks.append(current.strip())
            current = sent
        else:
            current = f"{current} {sent}".strip()
    if current:
        chunks.append(current.strip())
    return chunks

def process_chunk_with_images(chunk: str) -> str:
    """Process a chunk of text, handling image placeholders."""
    # Extract image placeholders and their positions
    import re
    image_placeholders = []
    
    def replace_image(match):
        image_path = match.group(1)
        image_placeholders.append((match.start(), image_path))
        return f"[IMAGE:{image_path}]"  # Keep a placeholder that won't be modified by cleaning
    
    # Replace image placeholders with a temporary marker
    processed_chunk = re.sub(r'\[IMAGE:(.*?)\]', replace_image, chunk)
    
    # Clean the text (without affecting image placeholders)
    cleaned = clean_transcript(processed_chunk)
    
    # Only summarize if there's actual text content (not just images)
    text_content = re.sub(r'\[IMAGE:.*?\]', '', cleaned).strip()
    if text_content:
        summary = summarize_transcript(cleaned, include_header=False)
        summary = summary.replace("```", "").strip()
    else:
        summary = ""
    
    # Re-insert image placeholders
    if image_placeholders:
        for pos, img_path in image_placeholders:
            # Insert images after the summary
            if summary:
                summary = f"{summary}\n\n[IMAGE:{img_path}]"
            else:
                summary = f"[IMAGE:{img_path}]"
    
    return summary

def chunked_clean_and_summarize(
    raw_transcript: str,
    debug: bool = True,
    final_merge: bool = True
    ) -> Generator[str, None, str]:
    """
    Generator version:
      - Yields each summary_part for a raw_chunk as soon as ready.
      - Preserves image placeholders in the output.
      - After all chunks, if final_merge is True and combined is too large,
        yields the final merged summary as one last item.
    Returns:
      The final combined summary (same as before), but through StopIteration.value.
    """
    os.makedirs("debug_chunks", exist_ok=True)

    raw_chunks = _split_raw_into_chunks(raw_transcript)
    full_summary_parts: List[str] = []

    # Process each chunk, preserving image placeholders
    for idx, raw_chunk in enumerate(raw_chunks, start=1):
        if debug:
            with open(f"debug_chunks/raw_chunk_{idx}.txt", "w", encoding="utf-8") as f:
                f.write(raw_chunk)

        # Process the chunk, handling images and text
        summary_part = process_chunk_with_images(raw_chunk)
        
        if debug:
            with open(f"debug_chunks/summary_chunk_{idx}.txt", "w", encoding="utf-8") as f:
                f.write(summary_part)

        full_summary_parts.append(summary_part)
        # Yield this chunk's summary immediately:
        yield summary_part

    # 4: Concatenate all parts
    combined = "\n".join(full_summary_parts)

    # 5: Optionally do a final merge and yield that as well
    if final_merge and len(combined) > MAX_CHARS_PER_CHUNK:
        merge_prompt = (
            "You are an expert at merging multiple transcript summaries into one concise summary. "
            "Preserve all timestamps, NER, numbers, and Active Voice format.\n\n"
            f"PARTIAL_SUMMARIES:\n{combined}"
        )
        combined = call_openai_chat(merge_prompt)
        # Clean up any triple backticks from the final merged response
        combined = combined.replace("```", "").strip()
        # Clean up any markdown formatting from the final merged response
        combined = combined.replace("```", "").strip()
        
        # Add the timeline header to the final merged output
        combined = f"<b>Timelines (Times are in Eastern time (GMT-5) unless otherwise noted):</b>\n\n{combined}"
        
        if debug:
            with open("debug_chunks/final_merged.txt", "w", encoding="utf-8") as f:
                f.write(combined)
                
        yield combined

    # Return the full combined summary to consumer via StopIteration.value
    return combined, 

def recursive_summarize(
    text: str,
    chunk_size: int = 5000,
    max_rounds: int = 2
    ) -> str:
    """
    If `text` is too large, split into sub‐chunks of ~chunk_size chars,
    summarize each, and concatenate. Repeat up to max_rounds times.
    """
    from .openai_client import call_openai_chat

    def _chunk_text(long_text: str) -> List[str]:
        """Split on sentence boundaries keeping each piece ≤ chunk_size."""
        sents = nltk.sent_tokenize(long_text)
        chunks, cur = [], ""
        for s in sents:
            if len(cur) + len(s) + 1 > chunk_size:
                chunks.append(cur.strip())
                cur = s
            else:
                cur = f"{cur} {s}".strip()
        if cur:
            chunks.append(cur.strip())

        
        return chunks

    # Base case: small enough
    if len(text) <= chunk_size or max_rounds <= 0:
        return text

    # Otherwise, split & summarize each piece
    pieces = _chunk_text(text)
    meta_summaries = []
    for piece in pieces:
        prompt = (
            "You are condensing a transcript summary. "
            "Keep Active Voice, timestamps, names, and numbers. "
            "Summarize this chunk briefly:\n\n" + piece
        )
        meta_summaries.append(call_openai_chat(prompt))
        
    combined_meta = "\n".join(meta_summaries)
    combined_meta = combined_meta.replace("`", "")
    # Recurse one fewer round
    return recursive_summarize(combined_meta, chunk_size=chunk_size, max_rounds=max_rounds-1)
