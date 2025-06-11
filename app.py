"""
@file app.py
@description Flask application acting as the backend for Google Docs AI Add-on.
             Handles requests from Apps Script, interacts with Gemini LLM,
             and applies changes to Google Docs via its API.
"""
import os
from dotenv import load_dotenv

load_dotenv() # This line loads variables from .env into os.environ


from flask import Flask, request, jsonify
import google.auth
from google.auth.transport.requests import Request as GoogleAuthRequest
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import json
import requests # For calling Gemini API

app = Flask(__name__)

# --- Configuration ---
# IMPORTANT: Never hardcode sensitive information like API keys or client secrets
# in production code. Use environment variables.
# You will need to set up OAuth 2.0 for a web application in Google Cloud Console
# and download the client_secret.json.
# For production, securely store and retrieve credentials, possibly using Google Secret Manager.

# Replace with your actual Gemini API Key from environment variables.
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
if not GEMINI_API_KEY:
    print("WARNING: GEMINI_API_KEY environment variable not set. Gemini API calls will fail.")

# Google Docs API Scopes
SCOPES = ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive.readonly']
# --- End Configuration ---

def get_google_docs_service(user_token):
    """
    Authenticates and returns a Google Docs API service object using a user's token.
    In a real scenario, the Apps Script would pass the user's OAuth2 access token
    to this backend, and you'd refresh it if necessary.
    For simplicity, this example assumes `user_token` is a valid access token.
    A more robust solution would involve Google Cloud Identity-Aware Proxy (IAP)
    or a dedicated OAuth2 flow on the backend.
    """
    # For a real deployed backend, you would typically receive an access token
    # from your Apps Script frontend or handle the OAuth flow server-side.
    # This is a placeholder for demonstration.
    creds = Credentials(token=user_token) # Assuming user_token is a valid access token

    try:
        service = build('docs', 'v1', credentials=creds)
        return service
    except HttpError as err:
        print(f"Error building Docs API service: {err}")
        return None

def get_document_content(docs_service, document_id):
    """Retrieves the full content of a Google Doc along with its structural elements."""
    try:
        document = docs_service.documents().get(documentId=document_id).execute()
        # Return the raw document body content for analysis, including structural elements
        return document.get('body', {})
    except HttpError as err:
        print(f"Error getting document content: {err}")
        return None

def call_gemini_for_suggestions(context, prompt_type, user_prompt):
    """
    Calls the Gemini API to get intelligent suggestions based on the context and prompt.
    The response is structured JSON for easy parsing.
    """
    if not GEMINI_API_KEY:
        return {"suggestions": []}

    # Crafting a robust prompt for the LLM is crucial for good results.
    # It explains the task, provides context, and defines the desired JSON output format.
    base_prompt = f"""
    You are an AI assistant for Google Docs. Based on the provided context,
    generate suggestions in a structured JSON format.

    **Context:**
    {json.dumps(context, indent=2)}

    **User Request:** "{user_prompt}"

    ---

    **Instructions for JSON Output:**
    The output should be a JSON object with a single key: "requests".
    The value of "requests" should be an array of objects, each representing a Google Docs API batchUpdate request.
    Each object should correspond to a single Google Docs API request type (e.g., "updateTextStyle", "updateParagraphStyle", "insertText").
    Ensure all `startIndex` and `endIndex` values are correct character offsets relative to the *beginning of the document*.
    When inserting text, provide `cursorIndex` or `insertionIndex` as the target index.

    **Supported Request Types & Examples:**

    1.  **To set a Heading (e.g., H1, H2, H3):**
        {{
            "updateParagraphStyle": {{
                "range": {{"startIndex": <start_index>, "endIndex": <end_index>}},
                "paragraphStyle": {{"namedStyleType": "HEADING_<level>"}},
                "fields": "namedStyleType"
            }}
        }}
        (Levels: 1-6)

    2.  **To italicize text:**
        {{
            "updateTextStyle": {{
                "range": {{"startIndex": <start_index>, "endIndex": <end_index>}},
                "textStyle": {{"italic": true}},
                "fields": "italic"
            }}
        }}

    3.  **To bold text:**
        {{
            "updateTextStyle": {{
                "range": {{"startIndex": <start_index>, "endIndex": <end_index>}},
                "textStyle": {{"bold": true}},
                "fields": "bold"
            }}
        }}

    4.  **To insert new text:**
        {{
            "insertText": {{
                "location": {{"index": <insertion_index>}},
                "text": "<generated_text_here>"
            }}
        }}

    5.  **To create a bulleted list:**
        {{
            "createParagraphBullets": {{
                "range": {{"startIndex": <paragraph_start_index>, "endIndex": <paragraph_end_index>}},
                "bulletPreset": "BULLET_DISC_CIRCLE_SQUARE"
            }}
        }}
        (Applies to paragraph(s) in the range)

    6.  **To delete content:**
        {{
            "deleteContentRange": {{
                "range": {{"startIndex": <start_index>, "endIndex": <end_index>}}
            }}
        }}

    Your task is to generate the JSON array of Google Docs API requests that fulfill the user's request, based on the provided context.
    Do not include any other text or explanation outside the JSON.
    """

    api_url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={GEMINI_API_KEY}"
    payload = {
        "contents": [{"role": "user", "parts": [{"text": base_prompt}]}],
        "generationConfig": {
            "responseMimeType": "application/json",
            "responseSchema": {
                "type": "OBJECT",
                "properties": {
                    "requests": {
                        "type": "ARRAY",
                        "items": {
                            "type": "OBJECT",
                            "properties": {
                                "insertText": {"type": "OBJECT"},
                                "updateTextStyle": {"type": "OBJECT"},
                                "updateParagraphStyle": {"type": "OBJECT"},
                                "createParagraphBullets": {"type": "OBJECT"},
                                "deleteContentRange": {"type": "OBJECT"},
                                # Add more as needed
                            }
                        }
                    }
                }
            }
        }
    }

    try:
        response = requests.post(api_url, headers={'Content-Type': 'application/json'}, json=payload)
        response.raise_for_status()
        result = response.json()
        if result.get('candidates') and result['candidates'][0].get('content'):
            llm_response_text = result['candidates'][0]['content']['parts'][0]['text']
            # LLM output is a stringified JSON, so parse it.
            return json.loads(llm_response_text)
        return {"requests": []}
    except requests.exceptions.RequestException as e:
        print(f"Error calling Gemini API: {e}")
        return {"requests": []}
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON from Gemini API: {e}. Raw response: {response.text}")
        return {"requests": []}

@app.route('/format', methods=['POST'])
def handle_format_request():
    """
    Handles requests to format selected text.
    Receives document ID, selected text, start/end indices, and user prompt.
    """
    data = request.json
    document_id = data.get('documentId')
    selected_text = data.get('selectedText')
    start_index = data.get('startIndex')
    end_index = data.get('endIndex')
    user_prompt = data.get('prompt')
    # In a real app, you'd get the user's auth token here.
    # For now, let's assume a placeholder token.
    user_auth_token = os.environ.get("GOOGLE_DOCS_AUTH_TOKEN") # Placeholder

    if not all([document_id, selected_text, start_index is not None, end_index is not None, user_prompt, user_auth_token]):
        return jsonify({"error": "Missing data or authentication token"}), 400

    docs_service = get_google_docs_service(user_auth_token)
    if not docs_service:
        return jsonify({"error": "Google Docs API authentication failed"}), 500

    # Get more context if needed (e.g., surrounding paragraphs)
    full_document_content = get_document_content(docs_service, document_id)
    context = {
        "documentId": document_id,
        "selectedText": selected_text,
        "selectionRange": {"startIndex": start_index, "endIndex": end_index},
        "fullDocumentBody": full_document_content # Send the full document structure
    }

    # Call Gemini for suggestions. The LLM's job is to decide the API requests.
    gemini_response = call_gemini_for_suggestions(context, "format", user_prompt)
    requests_to_apply = gemini_response.get('requests', [])

    if not requests_to_apply:
        return jsonify({"message": "AI did not generate any formatting requests."}), 200

    try:
        # Apply the requests generated by the LLM to the document.
        # Note: Index management can be complex if multiple operations shift indices.
        # The LLM must be smart about generating correct indices.
        # For simplicity, we assume the LLM provides correct absolute indices.
        docs_service.documents().batchUpdate(
            documentId=document_id, body={'requests': requests_to_apply}).execute()
        return jsonify({"message": "Formatting applied successfully."}), 200
    except HttpError as err:
        print(f"Error applying batch updates: {err}")
        return jsonify({"error": f"Failed to apply formatting: {err.content.decode()}"}), 500

@app.route('/generate', methods=['POST'])
def handle_generate_request():
    """
    Handles requests to generate new content at the cursor/selection.
    Receives document ID, cursor index, surrounding text, and user prompt.
    """
    data = request.json
    document_id = data.get('documentId')
    cursor_index = data.get('cursorIndex') # Or selection range
    surrounding_text = data.get('surroundingText')
    user_prompt = data.get('prompt')
    user_auth_token = os.environ.get("GOOGLE_DOCS_AUTH_TOKEN") # Placeholder

    if not all([document_id, cursor_index is not None, user_prompt, user_auth_token]):
        return jsonify({"error": "Missing data or authentication token"}), 400

    docs_service = get_google_docs_service(user_auth_token)
    if not docs_service:
        return jsonify({"error": "Google Docs API authentication failed"}), 500

    full_document_content = get_document_content(docs_service, document_id)
    context = {
        "documentId": document_id,
        "cursorIndex": cursor_index,
        "surroundingText": surrounding_text,
        "fullDocumentBody": full_document_content
    }

    gemini_response = call_gemini_for_suggestions(context, "generate", user_prompt)
    requests_to_apply = gemini_response.get('requests', [])

    if not requests_to_apply:
        return jsonify({"message": "AI did not generate any content."}), 200

    try:
        docs_service.documents().batchUpdate(
            documentId=document_id, body={'requests': requests_to_apply}).execute()
        return jsonify({"message": "Content generated and inserted successfully."}), 200
    except HttpError as err:
        print(f"Error applying batch updates: {err}")
        return jsonify({"error": f"Failed to generate content: {err.content.decode()}"}), 500

if __name__ == '__main__':
    # For local testing, you might use Flask's built-in server.
    # For deployment, use a production-ready WSGI server like Gunicorn or uWSGI.
    # Set the FLASK_APP environment variable: export FLASK_APP=app.py
    # Set the GOOGLE_DOCS_AUTH_TOKEN and GEMINI_API_KEY environment variables.
    # Example: export GOOGLE_DOCS_AUTH_TOKEN="your_actual_docs_access_token"
    # This token needs to be obtained through an OAuth2 flow in a real setup.
    app.run(host='0.0.0.0', port=5000, debug=True)