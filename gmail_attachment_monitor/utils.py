# -*- coding: utf-8 -*-
"""
Utility functions for the Gmail Attachment Monitor application.

Includes functions for parsing email bodies, formatting sender names,
escaping HTML, and potentially other helper tasks.
"""

import base64
import io
import re
import zipfile
import PyPDF2
import docx
import pandas as pd
from bs4 import BeautifulSoup
import html  # <-- Import the standard html module for escaping

# --- Helper Functions for Attachment Processing & Email Parsing ---

# NOTE: Attachment processing functions (download_attachment, extract_text_from_attachment)
# remain in main_window.py because they directly need access to 'self.service' and 'self.log'.
# If this project grew, dependency injection or passing service/logger instances
# could be used to move them here.


def get_attachments_info(parts):
    """
    Iterates through email message parts to identify and gather information
    about attachments (filename, mimeType, attachmentId, size).

    Handles nested parts (multipart messages).
    """
    attachments = []
    if not parts:
        return attachments
    for part in parts:
        # Recursively check nested parts
        if 'parts' in part:
            attachments.extend(get_attachments_info(part['parts']))

        # Check if the part represents an attachment
        filename = part.get('filename')
        body = part.get('body')
        if filename and body and body.get('attachmentId'):
            attachments.append({
                'filename': filename,
                'mimeType': part.get('mimeType', 'application/octet-stream'), # Default MIME type
                'attachmentId': body['attachmentId'],
                'size': body.get('size', 0) # Size in bytes
            })
        # TODO: Potentially add handling for inline content treated as attachments
        # based on Content-Disposition header if needed later.
    return attachments


def get_email_body(message):
    """
    Extracts the textual body content from a Gmail message object.

    Prefers 'text/plain' content, falls back to converting 'text/html' to plain text.
    Handles various multipart structures. Returns an empty string if no text body is found.
    """
    body_content = ""
    if 'payload' not in message:
        return "" # No payload to process

    payload = message['payload']
    plain_part_data = None
    html_part_data = None
    plain_part_encoding = 'utf-8' # Default assumption
    html_part_encoding = 'utf-8' # Default assumption

    def find_content_parts(parts_list):
        """Helper to recursively find plain and html parts."""
        nonlocal plain_part_data, html_part_data, plain_part_encoding, html_part_encoding
        if not parts_list:
            return

        for part in parts_list:
            mime_type = part.get('mimeType', '').lower()
            body = part.get('body', {})
            data = body.get('data')
            charset = 'utf-8' # Default per part
            # Try to find charset in headers
            for header in part.get('headers', []):
                 if header.get('name', '').lower() == 'content-type':
                      ctype_match = re.search(r'charset="?([^"]+)"?', header.get('value', ''), re.IGNORECASE)
                      if ctype_match:
                           charset = ctype_match.group(1).lower()
                           # Handle common aliases/misspellings if needed
                           if charset == 'windows-1252': charset = 'cp1252'
                           elif charset == 'iso-8859-1': charset = 'latin-1'
                           break # Found charset for this part

            if not data:
                 # If no direct data, check nested parts (like multipart/alternative)
                 if mime_type.startswith('multipart/') and 'parts' in part:
                     find_content_parts(part['parts'])
                 continue # Skip parts without data

            # Store the first found plain and html parts' data and encoding
            if mime_type == 'text/plain' and plain_part_data is None:
                plain_part_data = data
                plain_part_encoding = charset
            elif mime_type == 'text/html' and html_part_data is None:
                html_part_data = data
                html_part_encoding = charset

            # Optimization: If we found both, maybe stop early? Depends on email structure.
            # if plain_part_data and html_part_data:
            #     return

    # Start searching from the main payload parts
    find_content_parts(payload.get('parts'))

    # Handle case where body data is directly in the main payload (simple messages)
    if plain_part_data is None and html_part_data is None and 'data' in payload.get('body', {}):
        mime_type = payload.get('mimeType', '').lower()
        charset = 'utf-8'
        for header in payload.get('headers', []):
             if header.get('name', '').lower() == 'content-type':
                  ctype_match = re.search(r'charset="?([^"]+)"?', header.get('value', ''), re.IGNORECASE)
                  if ctype_match:
                       charset = ctype_match.group(1).lower()
                       if charset == 'windows-1252': charset = 'cp1252'
                       elif charset == 'iso-8859-1': charset = 'latin-1'
                       break
        data = payload['body']['data']
        if mime_type == 'text/plain':
            plain_part_data = data
            plain_part_encoding = charset
        elif mime_type == 'text/html':
            html_part_data = data
            html_part_encoding = charset


    # --- Decode the preferred part ---
    decoded_successfully = False
    if plain_part_data:
        try:
            # Decode base64
            decoded_bytes = base64.urlsafe_b64decode(plain_part_data)
            # Decode using detected or default encoding
            body_content = decoded_bytes.decode(plain_part_encoding, errors='replace')
            decoded_successfully = True
            # print(f"DEBUG: Decoded plain text with {plain_part_encoding}")
        except Exception as e:
            print(f"Warning: Failed to decode plain text part with {plain_part_encoding}: {e}")
            # Optionally try other encodings as a fallback here if needed

    # If plain text failed or wasn't found, try HTML
    if not decoded_successfully and html_part_data:
        try:
            # Decode base64
            decoded_bytes = base64.urlsafe_b64decode(html_part_data)
            # Decode using detected or default encoding
            html_content = decoded_bytes.decode(html_part_encoding, errors='replace')
            # Parse HTML and extract text
            soup = BeautifulSoup(html_content, 'html.parser')
            # Remove script and style elements
            for script_or_style in soup(["script", "style"]):
                script_or_style.decompose()
            # Get text content, separating blocks with newlines
            body_content = soup.get_text(separator='\n', strip=True)
            decoded_successfully = True
            # print(f"DEBUG: Decoded HTML part with {html_part_encoding} and extracted text")
        except Exception as e:
            print(f"Warning: Failed to decode/parse HTML part with {html_part_encoding}: {e}")


    if not decoded_successfully:
         print("Warning: Could not find or decode a text/plain or text/html body part.")
         return "[Body content not found or could not be decoded]"

    return body_content


def format_sender(sender: str) -> str:
    """
    Formats the sender string (e.g., "Display Name <email@example.com>")
    to extract just the Display Name if available, otherwise returns the original string.
    Handles optional quotes around the display name.
    """
    if not sender:
        return "Unknown Sender"
    # Regex to capture name inside optional quotes, followed by <email>
    match = re.match(r'^\s*"?([^"]*?)"?\s*<.+@.+>\s*$', sender.strip())
    if match and match.group(1):
        return match.group(1).strip() # Return captured name

    # Fallback: try to extract text before the first '<' if no quotes match
    match_simple = re.match(r'^(.*?)<.+@.+>\s*$', sender.strip())
    if match_simple and match_simple.group(1):
         return match_simple.group(1).strip() # Return text before <

    # If no angle brackets/email found, return the original string (might just be a name or email)
    return sender.strip()


# --- CORRECTED escape_html function ---
def escape_html(text: str) -> str:
     """
     Escapes special characters in a string for safe HTML display.

     Uses the standard library `html.escape` for robustness.
     """
     if not isinstance(text, str):
          text = str(text) # Ensure it's a string

     # Use the built-in html.escape which is more reliable and handles more cases
     # quote=True ensures " and ' are also escaped.
     return html.escape(text, quote=True)
# --- End corrected function ---


def format_body_text_no_highlight(body: str) -> str:
    """
    Formats the extracted plain text email body for safe HTML display in QTextEdit.
    Replaces newlines with <br> tags. Does NOT perform keyword highlighting.
    """
    if not body:
        return ""
    # Use the corrected escape_html function
    safe_body = escape_html(body)
    # Convert newlines to HTML line breaks
    formatted_body = safe_body.replace("\n", "<br>")
    return formatted_body

# --- Other potential utils could go here ---
# e.g., date formatting, file size formatting, etc.