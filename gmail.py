import os
import base64
import pickle
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from typing import List
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from mcp.server.fastmcp import FastMCP
import mimetypes

mcp = FastMCP("enhanced_gmail_mcp_server")

# Extended OAuth2 scopes for full email functionality
SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.modify'
]

def get_gmail_service():
    """Authenticate and return authorized Gmail API service instance."""
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    
    return build('gmail', 'v1', credentials=creds)

@mcp.tool()
def get_inbox_emails(max_results: int = 10) -> List[dict]:
    """Retrieve emails from inbox with message details."""
    try:
        service = get_gmail_service()
        result = service.users().messages().list(
            userId='me',
            labelIds=['INBOX'],
            maxResults=max_results
        ).execute()
        
        messages = []
        for msg in result.get('messages', []):
            message = service.users().messages().get(
                userId='me',
                id=msg['id'],
                format='metadata'
            ).execute()

            
            headers = {}
            for header in message['payload']['headers']:
                try:
                    name = header['name']
                    value = header['value'].encode('utf-8').decode('utf-8', errors='replace')
                    headers[name] = value
                except Exception as e:
                    headers[header['name']] = f"[Error decoding header: {str(e)}]"

            # Extract body content
            body = ''
            try:
                if 'parts' in message['payload']:
                    # Multipart message
                    for part in message['payload']['parts']:
                        if part['mimeType'] == 'text/plain':
                            body_bytes = base64.urlsafe_b64decode(part['body']['data'])
                            body = body_bytes.decode('utf-8', errors='replace')
                            break
                elif 'body' in message['payload']:
                    # Single part message
                    body_bytes = base64.urlsafe_b64decode(message['payload']['body']['data'])
                    body = body_bytes.decode('utf-8', errors='replace')
            except Exception as e:
                body = f"[Error decoding body: {str(e)}]"
            
            messages.append({
                'id': message['id'],
                'threadId': message['threadId'],
                'snippet': message.get('snippet', ''),
                'from': headers.get('From', ''),
                'to': headers.get('To', ''),
                'subject': headers.get('Subject', ''),
                'date': headers.get('Date', ''),
                'body': body,
                'labels': message.get('labelIds', []),
                'has_attachments': bool(message['payload'].get('parts', [])),
                'importance': headers.get('Importance', 'normal'),
                'cc': headers.get('Cc', ''),
                'bcc': headers.get('Bcc', '')
            })
        
        return messages

    
    except Exception as e:
        return {'error': str(e)}

def create_message_with_attachments(
    to: str,
    subject: str,
    body: str,
    attachments: List[str] = None
) -> dict:
    """Create MIME message with optional attachments."""
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = 'me'
    message['subject'] = subject
    message.attach(MIMEText(body, 'plain'))
    
    if attachments:
        for filepath in attachments:
            content_type, encoding = mimetypes.guess_type(filepath)
            if content_type is None or encoding is not None:
                content_type = 'application/octet-stream'
            
            main_type, sub_type = content_type.split('/', 1)
            with open(filepath, 'rb') as f:
                part = MIMEBase(main_type, sub_type)
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename="{os.path.basename(filepath)}"'
                )
                message.attach(part)
    
    return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}

@mcp.tool()
def send_email(
    to: str,
    subject: str,
    body: str,
    attachments: List[str] = None
) -> dict:
    """Send email with optional attachments."""
    try:
        service = get_gmail_service()
        message = create_message_with_attachments(to, subject, body, attachments)
        sent_message = service.users().messages().send(
            userId='me',
            body=message
        ).execute()
        
        return {
            'status': 'success',
            'messageId': sent_message['id'],
            'threadId': sent_message['threadId']
        }
    
    except Exception as e:
        return {'error': str(e)}


@mcp.tool()
def search_emails(query: str, max_results: int = 20) -> List[dict]:
    """Search emails using Gmail search operators."""
    try:
        service = get_gmail_service()
        result = service.users().messages().list(
            userId='me',
            q=query,
            maxResults=max_results
        ).execute()

        messages = []

        for msg in result.get('messages', []):
            message = service.users().messages().get(
                userId='me',
                id=msg['id'],
                format='metadata'
            ).execute()

            
            headers = {}
            for header in message['payload']['headers']:
                try:
                    name = header['name']
                    value = header['value'].encode('utf-8').decode('utf-8', errors='replace')
                    headers[name] = value
                except Exception as e:
                    headers[header['name']] = f"[Error decoding header: {str(e)}]"

            # Extract body content
            body = ''
            try:
                if 'parts' in message['payload']:
                    # Multipart message
                    for part in message['payload']['parts']:
                        if part['mimeType'] == 'text/plain':
                            body_bytes = base64.urlsafe_b64decode(part['body']['data'])
                            body = body_bytes.decode('utf-8', errors='replace')
                            break
                elif 'body' in message['payload']:
                    # Single part message
                    body_bytes = base64.urlsafe_b64decode(message['payload']['body']['data'])
                    body = body_bytes.decode('utf-8', errors='replace')
            except Exception as e:
                body = f"[Error decoding body: {str(e)}]"
            
            messages.append({
                'id': message['id'],
                'threadId': message['threadId'],
                'snippet': message.get('snippet', ''),
                'from': headers.get('From', ''),
                'to': headers.get('To', ''),
                'subject': headers.get('Subject', ''),
                'date': headers.get('Date', ''),
                'body': body,
                'labels': message.get('labelIds', []),
                'has_attachments': bool(message['payload'].get('parts', [])),
                'importance': headers.get('Importance', 'normal'),
                'cc': headers.get('Cc', ''),
                'bcc': headers.get('Bcc', '')
            })
        
        return messages

        
    except Exception as e:
        return {'error': str(e)}

@mcp.tool()
def get_sent_emails(max_results: int = 10) -> List[dict]:
    """Retrieve emails from sent folder with message details."""
    try:
        service = get_gmail_service()
        result = service.users().messages().list(
            userId='me',
            labelIds=['SENT'],
            maxResults=max_results
        ).execute()
        
        messages = []
        for msg in result.get('messages', []):
            message = service.users().messages().get(
                userId='me',
                id=msg['id'],
                format='full'
            ).execute()
            
            # Handle headers with proper encoding
            headers = {}
            for header in message['payload']['headers']:
                try:
                    name = header['name']
                    value = header['value'].encode('utf-8').decode('utf-8', errors='replace')
                    headers[name] = value
                except Exception as e:
                    headers[header['name']] = f"[Error decoding header: {str(e)}]"

            # Extract body content
            body = ''
            try:
                if 'parts' in message['payload']:
                    for part in message['payload']['parts']:
                        if part['mimeType'] == 'text/plain':
                            body_bytes = base64.urlsafe_b64decode(part['body']['data'])
                            body = body_bytes.decode('utf-8', errors='replace')
                            break
                elif 'body' in message['payload']:
                    body_bytes = base64.urlsafe_b64decode(message['payload']['body']['data'])
                    body = body_bytes.decode('utf-8', errors='replace')
            except Exception as e:
                body = f"[Error decoding body: {str(e)}]"
            
            messages.append({
                'id': message['id'],
                'threadId': message['threadId'],
                'snippet': message.get('snippet', '').encode('utf-8').decode('utf-8', errors='replace'),
                'from': headers.get('From', ''),
                'to': headers.get('To', ''),
                'subject': headers.get('Subject', ''),
                'date': headers.get('Date', ''),
                'body': body,
                'labels': message.get('labelIds', []),
                'has_attachments': bool(message['payload'].get('parts', [])),
                'importance': headers.get('Importance', 'normal'),
                'cc': headers.get('Cc', ''),
                'bcc': headers.get('Bcc', '')
            })
        
        return messages
    
    except Exception as e:
        return {'error': str(e)}

if __name__ == "__main__":
    print("Starting Enhanced Gmail MCP Server...")
    mcp.run(transport='stdio')
