"""
Mock version of openai_client for testing purposes
"""

def call_openai_chat(messages, model="gpt-4o", temperature=0, max_tokens=2048):
    """
    Mock implementation that returns a simple response
    """
    return {
        "choices": [{
            "message": {
                "content": "[{'question': 'What is the purpose of this meeting?', 'answer': 'The purpose of this meeting is to discuss the quarterly business review and plan for the next quarter.'}, {'question': 'Who will be attending the meeting?', 'answer': 'The meeting will be attended by the CEO, CFO, and department heads.'}]"
            }
        }]
    }
