"""
Extract FAQs (questions and answers) from documents using OpenAI's GPT.
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openai
from typing import List, Dict, Optional
from dotenv import load_dotenv
from src.text_processing import _split_raw_into_chunks
import ast

load_dotenv()

api_key = os.getenv('OPENAI_API_KEY')
if not api_key:
    raise ValueError("OPENAI_API_KEY environment variable not set. Please set it in a .env file or your environment.")
_client = openai.OpenAI(api_key=api_key)


def extract_faqs(text: str, max_chunk_size: int = 3000) -> List[Dict[str, str]]:
    """
    Extract FAQs (questions and answers) from the given text using OpenAI with chunking.
    
    Args:
        text: The text to extract FAQs from
        max_chunk_size: Maximum size of each chunk in characters
        
    Returns:
        List of dictionaries containing questions and their corresponding answers
    """
    try:
        # Split text into chunks using the existing text_processing function
        chunks = _split_raw_into_chunks(text, max_chunk_size)
        print(f"\nTotal chunks to process: {len(chunks)}")
        print(f"First chunk preview: {chunks[0][:100]}...")
        
        # Process each chunk and collect results
        faqs = []
        for idx, chunk in enumerate(chunks, 1):
            print(f"\nProcessing chunk {idx}/{len(chunks)} for FAQs...")
            print(f"Chunk size: {len(chunk)} characters")
            print(f"Chunk preview: {chunk[:50]}...")
            chunk_faqs = _process_chunk_for_faqs(chunk)
            print(f"Found {len(chunk_faqs)} FAQs in this chunk")
            if chunk_faqs:
                print(f"First FAQ in chunk: {chunk_faqs[0]}")
            faqs.extend(chunk_faqs)
        
        print(f"\nTotal FAQs found: {len(faqs)}")
        if faqs:
            print(f"First FAQ overall: {faqs[0]}")
        return faqs
    except Exception as e:
        print(f"Error in extract_faqs: {e}")
        raise

def _process_chunk_for_faqs(chunk: str) -> List[Dict[str, str]]:
    """
    Helper function to process a single chunk of text for FAQs using OpenAI.
    """
    try:
        response = _client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "system",
                    "content": "You are a helpful assistant that extracts technical FAQs from text."
                },
                {
                    "role": "user",
                    "content": f"""Extract only technical questions and their corresponding answers from the following text. 
                    Format your response as a JSON array of objects with 'question' and 'answer' keys.
                    Only include complete question-answer pairs that are clearly related.
                    
                    Text: {chunk}"""
                }
            ],
            temperature=0,
            max_tokens=2048
        )
        
        # Get the response content and clean it up
        response_content = response.choices[0].message.content
        
        # Print raw response for debugging
        print(f"\nRaw response content: {response_content}")
        
        # Remove any code block markers and trim whitespace
        # Try multiple patterns to clean the response
        patterns = [
            (r'^```json\s*', ''),  # Remove ```json at start
            (r'^```\s*', ''),      # Remove ``` at start
            (r'```\s*$', ''),      # Remove ``` at end
            (r'\s+', ' '),         # Replace multiple spaces with single space
            (r'^\s+', ''),         # Remove leading spaces
            (r'\s+$', '')          # Remove trailing spaces
        ]
        
        cleaned_content = response_content
        for pattern, replacement in patterns:
            cleaned_content = re.sub(pattern, replacement, cleaned_content, flags=re.MULTILINE)
        
        # Try multiple parsing methods
        try:
            # Try ast.literal_eval first
            try:
                faqs = ast.literal_eval(cleaned_content)
                if isinstance(faqs, list) and all(isinstance(item, dict) and 'question' in item and 'answer' in item for item in faqs):
                    return faqs
                print(f"Parsed content doesn't match expected format: {faqs}")
            except Exception as e:
                print(f"ast.literal_eval failed: {e}")
                
            # Try json.loads
            try:
                import json
                faqs = json.loads(cleaned_content)
                if isinstance(faqs, list) and all(isinstance(item, dict) and 'question' in item and 'answer' in item for item in faqs):
                    return faqs
                print(f"JSON parsed content doesn't match expected format: {faqs}")
            except Exception as e:
                print(f"json.loads failed: {e}")
                
            # Try to fix JSON formatting issues
            try:
                # Add missing quotes around keys
                fixed_content = re.sub(r'([{,]\s*)(\w+)(\s*:)', r'\1"\2"\3', cleaned_content)
                # Add missing quotes around values if needed
                fixed_content = re.sub(r'([":]\s*)([^\[\]{}"\',\n]+)([\s\],}])', r'\1"\2"\3', fixed_content)
                
                faqs = json.loads(fixed_content)
                if isinstance(faqs, list) and all(isinstance(item, dict) and 'question' in item and 'answer' in item for item in faqs):
                    print(f"Successfully parsed after fixing JSON format")
                    return faqs
                print(f"Fixed JSON parsed content doesn't match expected format: {faqs}")
            except Exception as e:
                print(f"Failed to parse after fixing JSON format: {e}")
                
            # If all parsing attempts fail, return empty list
            print(f"\nFailed to parse FAQ response after all attempts:")
            print(f"Original content: {response_content}")
            print(f"Cleaned content: {cleaned_content}")
            return []
            
        except Exception as e:
            print(f"Error parsing FAQ response: {e}")
            print(f"Final content after all cleanup: {cleaned_content}")
            return []
            
    except Exception as e:
        print(f"Error processing chunk for FAQs: {e}")
        return []