# =============================================================================
# Author: Mustafa Nassar 
# Version: 3.0
# This script demonstrates how to send an email with an attachment using the Microsoft Graph API.
#
# Description:
# - Uses MSAL (Microsoft Authentication Library) for authentication.
# - Utilizes the Microsoft Graph API for sending emails, including file attachments.
# - Requires an Azure AD application with the following permissions:
#   - Application Permission: "Mail.Send"
#   - Scope: "https://graph.microsoft.com/.default"
# - The script sends an email with an attachment to a specified recipient.
# - All configurable variables (like Azure credentials, recipient email, and file path) are declared at the beginning.
# - Suppresses debug logs for `urllib3` and `msal` for cleaner output.
#
# Required Python Libraries:
# - msal (`pip install msal`)
# - requests (`pip install requests`)
#
# Make sure to update the variables (CLIENT_ID, CLIENT_SECRET, TENANT_ID, etc.) with your credentials.
# =============================================================================

import os
import base64
import msal
import requests
import logging

# ====================
# CONFIGURATION BLOCK
# ====================
# Azure AD App Credentials
CLIENT_ID = "915a5ce2-9346-4120-asa7-7caa4378052e"  # Your Azure App Client ID
CLIENT_SECRET = "WWa;sdlfkajd;sfjad;fja;ldfY"  # Your Azure App Secret
TENANT_ID = "5d223e2382c-d230-4e2e-84a6-02390537fbdce"  # Your Azure Tenant ID
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

# Email Settings
SENDER_EMAIL = "sender@domain.local"  # The email address of the sender
RECIPIENT_EMAIL = "recipient@domain.de"  # The email address of the recipient
EMAIL_SUBJECT = "Your Requested Report from OnPrem"  # Email subject
EMAIL_BODY = "Please find the attached report."  # Email body content

# Optional File Attachment
ATTACHMENT_PATH = r"C:\Users\mnass\Downloads\pythonApp\Report.txt"  # Path to the attachment file (comment out if not needed)

# Microsoft Graph Endpoint
ENDPOINT = f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail"

# ====================
# LOGGING CONFIGURATION
# ====================
# Set the logging level for urllib3 and msal to WARNING
logging.getLogger('urllib3').setLevel(logging.WARNING)
logging.getLogger('msal').setLevel(logging.WARNING)

# If you want to suppress all debug messages, set the root logger level to WARNING
logging.basicConfig(level=logging.WARNING)

# ====================
# AUTHENTICATION
# ====================
print("Acquiring access token...")  # Log progress for user visibility
app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)

# Acquire an access token from Microsoft Identity Platform
result = app.acquire_token_for_client(scopes=SCOPE)

# Check if the token acquisition was successful
if "access_token" in result:
    access_token = result["access_token"]
    print("Access token acquired successfully.")
else:
    print("Error acquiring access token.")
    print(f"Error: {result.get('error')}")
    print(f"Description: {result.get('error_description')}")
    exit()

# ====================
# PREPARE EMAIL
# ====================
print("Preparing the email...")  # Log progress for user visibility
file_name = os.path.basename(ATTACHMENT_PATH)  # Extract the file name from the file path

# Read and encode the attachment file in Base64
try:
    with open(ATTACHMENT_PATH, 'rb') as file:
        file_content = file.read()
        encoded_content = base64.b64encode(file_content).decode('utf-8')
except FileNotFoundError:
    print(f"Error: File not found at {ATTACHMENT_PATH}. Please check the file path.")
    exit()

# Construct the email message payload
email_msg = {
    "message": {
        "subject": EMAIL_SUBJECT,
        "body": {
            "contentType": "Text",
            "content": EMAIL_BODY
        },
        "toRecipients": [
            {
                "emailAddress": {
                    "address": RECIPIENT_EMAIL
                }
            }
        ],
        "attachments": [
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": file_name,
                "contentType": "application/octet-stream",
                "contentBytes": encoded_content
            }
        ]
    }
}

# ====================
# SEND EMAIL
# ====================
print("Sending the email...")  # Log progress for user visibility
headers = {
    'Authorization': f'Bearer {access_token}',  # Include the access token for authentication
    'Content-Type': 'application/json'  # Specify the content type for JSON payload
}

# Send the email via Microsoft Graph API
response = requests.post(ENDPOINT, headers=headers, json=email_msg)

# Check the response from the Graph API
if response.status_code == 202:  # Status code 202 indicates success
    print("Email sent successfully.")
else:
    print(f"Error sending email: {response.status_code}")
    print("Response details:", response.json())
