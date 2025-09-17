# -*- coding: utf-8 -*-
"""
Created on Fri Jun  6 15:16:46 2025

@author: ShivakrishnaBoora
"""



import os
import openai
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

def call_openai_chat(prompt: str, model: str = "gpt-3.5-turbo") -> str:
    """
    Send `prompt` to OpenAI's chat completions (v1+ style).
    Returns only the generated text.
    Raises RuntimeError if there's an error with the API call.
    """
    print(f"\n=== Sending to OpenAI (model: {model}) ===")
    print(f"Prompt length: {len(prompt)} characters")
    print(f"Prompt preview: {prompt[:200]}..." if len(prompt) > 200 else f"Prompt: {prompt}")
    
    try:
        # Get API key from environment variable
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            raise ValueError("OPENAI_API_KEY environment variable not set. Please set it in a .env file or your environment.")
        _client = openai.OpenAI(api_key=api_key)
        response = _client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.0,
            max_tokens=2048,
        )
        
        # # Debug print the raw response structure
        # print("\n=== Raw API Response ===")
        # print(f"Response type: {type(response)}")
        # print(f"Response object: {response}")
        
        if not response.choices or len(response.choices) == 0:
            error_msg = "API returned no choices in response"
            print(f"ERROR: {error_msg}")
            raise RuntimeError(error_msg)
            
        message_content = response.choices[0].message.content
        if message_content is None:
            error_msg = "API returned None as message content"
            print(f"ERROR: {error_msg}")
            print(f"Full response object: {response}")
            raise RuntimeError(error_msg)
            
        # print(f"Response content length: {len(message_content)} characters")
        # print(f"Response preview: {message_content[:200]}..." if len(message_content) > 200 else f"Response: {message_content}")
        
        return message_content.strip()
        
    except Exception as e:
        print(f"API Key: {api_key}")
        error_msg = f"OpenAI API Error: {str(e)}"
        print(f"\n=== ERROR ===\n{error_msg}")
        print(f"Error type: {type(e).__name__}")
        import traceback
        traceback.print_exc()
        raise RuntimeError(error_msg) from e


def get_embedding_similarity(text1: str, text2: str, model: str = "text-embedding-3-small") -> float:
    """
    Compute cosine similarity between two texts using OpenAI embeddings.
    Returns a float between 0.0 and 1.0.
    
    Args:
        text1: First text to compare
        text2: Second text to compare
        model: Embedding model to use (default: text-embedding-3-small)
        
    Returns:
        float: Cosine similarity between the two texts (0.0 to 1.0)
        
    Note:
        - Automatically truncates text to fit within token limits
        - Uses a more efficient model by default
    """
    import numpy as np
    from sklearn.metrics.pairwise import cosine_similarity
    import tiktoken
    
    # Initialize tokenizer for the specified model
    try:
        encoding = tiktoken.encoding_for_model(model)
    except KeyError:
        # Fallback to cl100k_base which works with most models
        encoding = tiktoken.get_encoding("cl100k_base")
    
    # Define token limits for different models
    model_max_tokens = {
        "text-embedding-3-small": 8191,
        "text-embedding-3-large": 8191,
        "text-embedding-ada-002": 8191
    }
    
    max_tokens = model_max_tokens.get(model, 8191)
    
    def truncate_text(text: str, max_tokens: int) -> str:
        """Truncate text to fit within token limit."""
        tokens = encoding.encode(text)
        if len(tokens) <= max_tokens:
            return text
        truncated_tokens = tokens[:max_tokens]
        return encoding.decode(truncated_tokens)
    
    try:
        # Truncate texts if needed
        text1 = truncate_text(text1, max_tokens)
        text2 = truncate_text(text2, max_tokens)
        
        # Get embeddings
        response = _client.embeddings.create(
            model=model,
            input=[text1, text2]
        )
        
        # Extract embeddings
        emb1 = np.array(response.data[0].embedding).reshape(1, -1)
        emb2 = np.array(response.data[1].embedding).reshape(1, -1)
        
        # Calculate similarity
        sim = cosine_similarity(emb1, emb2)[0][0]
        return float(sim)
        
    except Exception as e:
        # Fallback to simpler similarity measure if embedding fails
        print(f"Error computing embeddings: {str(e)}")
        # Use SequenceMatcher as fallback
        from difflib import SequenceMatcher
        return SequenceMatcher(None, text1[:1000], text2[:1000]).ratio()
