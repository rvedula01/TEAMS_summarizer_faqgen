"""
Test script for FAQ extraction functionality
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Override the openai_client module with our mock version
sys.modules['openai_client'] = type('MockModule', (), {
    'call_openai_chat': lambda *args, **kwargs: {
        "choices": [{
            "message": {
                "content": '[{"question": "What is the purpose of this meeting?", "answer": "The purpose of this meeting is to discuss the quarterly business review and plan for the next quarter."}, {"question": "Who will be attending the meeting?", "answer": "The meeting will be attended by the CEO, CFO, and department heads."}]'
            }
        }]
    }
})()

from src.faq import extract_faqs
import json

def test_faq_extraction():
    # Test document with multiple question-answer pairs
    test_text = """
    16:42pm What is the purpose of this meeting?
    16:43pm The purpose of this meeting is to discuss the quarterly business review and plan for the next quarter.

    16:44pm You are off for the weekend. So, lets get this done today itself.

    16:45pm Are there any specific action items from the previous meeting?
    16:46pm Yes, we need to finalize the budget for the marketing campaign and update the product roadmap.
    """

    print("\nTesting FAQ Extraction...")
    faqs = extract_faqs(test_text)
    
    print("\nExtracted FAQs:")
    for idx, faq_item in enumerate(faqs, 1):
        print(f"\nFAQ {idx}")
        print(f"Question: {faq_item['question']}")
        print(f"Answer: {faq_item['answer']}")
    
    # Save results to file for review
    with open('faq_extraction_results.json', 'w') as f:
        json.dump(faqs, f, indent=2)

if __name__ == "__main__":
    test_faq_extraction()
