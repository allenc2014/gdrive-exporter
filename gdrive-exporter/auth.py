"""
Authentication and API helpers module.
Handles Google API authentication and provides retry functionality.
"""

import os
import time
import logging
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = [
    "https://www.googleapis.com/auth/documents.readonly",
    "https://www.googleapis.com/auth/drive.readonly"
]

API_DELAY_SECONDS = 0.05  # small delay between API calls
RETRY_WAIT_SECONDS = 1.0
MAX_RETRIES = 5


def get_google_services():
    """
    Authenticate and get Google Docs and Drive services.
    
    Returns:
        Tuple of (docs_service, drive_service)
    """
    creds = None
    token_path = "token.json"

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)

    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
        creds = flow.run_local_server(port=0)
        with open(token_path, "w") as token:
            token.write(creds.to_json())

    docs_service = build("docs", "v1", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)
    return docs_service, drive_service


def call_with_retry(func, *args, **kwargs):
    """
    Call an API function with retry logic for rate limiting and server errors.
    
    Args:
        func: The API function to call
        *args: Arguments to pass to the function
        **kwargs: Keyword arguments to pass to the function
        
    Returns:
        The result of the function call
        
    Raises:
        HttpError: If the API call fails after retries
    """
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            result = func(*args, **kwargs)
            time.sleep(API_DELAY_SECONDS)
            return result
        except HttpError as e:
            status = e.resp.status if e.resp else None
            if status in (429, 503) and attempt < MAX_RETRIES:
                logging.warning(f"RETRY: HTTP {status}, waiting {RETRY_WAIT_SECONDS}s (attempt {attempt}/{MAX_RETRIES})")
                time.sleep(RETRY_WAIT_SECONDS)
                continue
            raise
