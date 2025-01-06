

# Author: Mustafa Nassar
# Version: 2.2
# Synopsis: This script demonstrates how to send an email with an attachment using Microsoft Graph API.
# Description:
# - Utilizes the Microsoft Graph SDK for Python for cleaner and maintainable code.
# - All configurable variables (like sender, recipient, attachment path, email body, and subject) are declared at the beginning.
# - Requires Azure AD application with necessary permissions and credentials.
# - Required Python Modules:
#   - msgraph (`pip install msgraph`)
#   - msal (`pip install msal`)
# - Ensure these modules are installed before running the script.

from msgraph import GraphServiceClient
from msgraph.generated.models import (
    Message,
    ItemBody,
    BodyType,
    Recipient,
    EmailAddress,
    FileAttachment,
)
from msgraph.generated.users.item.send_mail.send_mail_post_request_body import (
    SendMailPostRequestBody,
)
from msal import ConfidentialClientApplication
import os

# ====================
# CONFIGURATION BLOCK
# ====================
# Azure AD App Credentials
CLIENT_ID = "915a5ce2-9986-4540-a5a7-7caa4378052e"  # Your Azure App Client ID
CLIENT_SECRET = "aaaa~35~asdfasdfasNNaCAdaygSMc9eLFaMY"  # Your Azure App Secret
TENANT_ID = "5d23882c-d9f0-4e2e-84a6-0f290d7fbdce"  # Your Azure Tenant ID

# Email Settings
SENDER_EMAIL = "Reports@1590.eu"  # The email address of the sender
RECIPIENT_EMAIL = "mnassar365@outlook.com"  # The email address of the recipient
EMAIL_SUBJECT = "Your Requested Report from OnPrem"  # Email subject
EMAIL_BODY = "Please find the attached report."  # Email body content

# File Attachment
ATTACHMENT_PATH = r"C:\Users\mnassar\Downloads\pythonApp\Report.txt"  # Path to the attachment file

# ====================
# FUNCTIONS
# ====================
# Authentication: Get an access token using MSAL
def get_access_token():
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception("Failed to acquire access token")

# Initialize Graph client
def get_graph_client(access_token):
    return GraphServiceClient(
        auth_provider=lambda request: request.headers.update(
            {"Authorization": f"Bearer {access_token}"}
        )
    )

# Build and send email
def send_email(graph_client):
    # Read and encode the attachment
    file_name = os.path.basename(ATTACHMENT_PATH)
    with open(ATTACHMENT_PATH, "rb") as file:
        file_content = file.read()

    attachment = FileAttachment(
        odata_type="#microsoft.graph.fileAttachment",
        name=file_name,
        content_type="application/octet-stream",
        content_bytes=file_content,
    )

    email_message = Message(
        subject=EMAIL_SUBJECT,
        body=ItemBody(
            content_type=BodyType.TEXT,
            content=EMAIL_BODY,
        ),
        to_recipients=[
            Recipient(
                email_address=EmailAddress(address=RECIPIENT_EMAIL),
            )
        ],
        attachments=[attachment],
    )

    request_body = SendMailPostRequestBody(message=email_message, save_to_sent_items=True)

    # Send the email
    graph_client.users_by_id(SENDER_EMAIL).send_mail(request_body).post()
    print("Email sent successfully.")

# ====================
# MAIN SCRIPT EXECUTION
# ====================
if __name__ == "__main__":
    try:
        print("Acquiring access token...")
        access_token = get_access_token()

        print("Initializing Graph client...")
        graph_client = get_graph_client(access_token)

        print("Sending email...")
        send_email(graph_client)

    except Exception as e:
        print(f"An error occurred: {e}")
