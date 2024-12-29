from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os.path
import pickle
import base64
import datetime

class GmailClient:
    def __init__(self):
        # If modifying these scopes, delete the token.pickle file
        self.SCOPES = ['https://www.googleapis.com/auth/gmail.modify']
        self.creds = None
        self.service = None

    def authenticate(self):
        """Handles the OAuth2 authentication flow."""
        # Check if token.pickle exists with saved credentials
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                self.creds = pickle.load(token)
        
        # If no valid credentials available, let the user log in
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', self.SCOPES)
                self.creds = flow.run_local_server(port=0)
            
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(self.creds, token)

        self.service = build('gmail', 'v1', credentials=self.creds)

    def list_messages(self, query='', max_results=None):
        """Lists messages matching the specified query.
            max_results (int, optional): Maximum number of messages to return. If None, returns all messages.
        """
        try:
            messages = []
            next_page_token = None
            page_count = 0
            
            while True:
                page_count += 1

                #Make the request
                request = self.service.users().messages().list(
                    userId='me', 
                    q=query,
                    pageToken=next_page_token,
                    maxResults=min(max_results - len(messages), 100) if max_results else 100
                ).execute()
                
                # print("\nAPI Response:")
                # print(f"Response keys: {request.keys()}")
                if 'messages' in request:
                    print(f"Messages in this page: {len(request['messages'])}")
                    print("First few messages:")
                    for msg in request['messages'][:3]:
                        print(f"  - ID: {msg['id']}, ThreadId: {msg['threadId']}")
                else:
                    print("No messages in this page")
                
                # Get messages from response
                if 'messages' in request:
                    messages.extend(request['messages'])
                
                # Check if we've got enough messages
                if max_results and len(messages) >= max_results:
                    # print(f"\nReached max_results limit ({max_results}). Breaking.")
                    messages = messages[:max_results]
                    break
                    
                # Check if there are more pages
                next_page_token = request.get('nextPageToken')
                # print(f"Next page token: {next_page_token}")
                if not next_page_token:
                    # print("No more pages available")
                    break
            
            # print(f"\nFinal message count: {len(messages)}")
            
            if not messages:
                print('No messages found.')
                return []
                
            return messages
            
        except Exception as e:
            print(f'An error occurred: {e}')
            return []
    def read_message(self, msg_id):
        """Reads a specific message by ID and returns its content."""
        try:
            message = self.service.users().messages().get(
                userId='me', id=msg_id, format='full').execute()
            
            payload = message['payload']
            headers = payload.get('headers', [])
            
            # Extract subject and sender
            subject = next(h['value'] for h in headers if h['name'] == 'Subject')
            sender = next(h['value'] for h in headers if h['name'] == 'From')
            
            # Get the message body
            if 'parts' in payload:
                parts = payload['parts']
                data = parts[0]['body']['data']
            else:
                data = payload['body']['data']
            
            text = base64.urlsafe_b64decode(data).decode('utf-8')
            
            return {
                'subject': subject,
                'sender': sender,
                'body': text
            }
        except Exception as e:
            print(f'An error occurred: {e}')
            return None

    def read_message(self, msg_id):
        """Reads a specific message by ID and returns its content and attachment info."""
        try:
            message = self.service.users().messages().get(
                userId='me', id=msg_id, format='full').execute()
            
            # Extract headers
            payload = message['payload']
            headers = payload.get('headers', [])
            
            # Get subject and sender
            subject = next((h['value'] for h in headers if h['name'] == 'Subject'), '')
            sender = next((h['value'] for h in headers if h['name'] == 'From'), '')
            
            # Initialize body text and attachments list
            body_text = ''
            attachments = []
            
            def process_parts(part, prefix=""):
                """Recursively process message parts to find body and attachments."""
                nonlocal body_text
                
                # Get part metadata
                mime_type = part.get('mimeType', '')
                filename = part.get('filename', '')
                part_id = part.get('body', {}).get('attachmentId', '')
                
                # Handle text parts
                if mime_type == 'text/plain' and not filename:
                    body_data = part.get('body', {}).get('data', '')
                    if body_data:
                        body_text = base64.urlsafe_b64decode(body_data).decode('utf-8')
                
                # Handle attachments
                elif filename or part_id:
                    attachment_info = {
                        'filename': filename or f'attachment_{part_id}',
                        'mime_type': mime_type,
                        'attachment_id': part_id,
                        'size': part.get('body', {}).get('size', 0)
                    }
                    attachments.append(attachment_info)
                
                # Process nested parts
                if 'parts' in part:
                    for subpart in part['parts']:
                        process_parts(subpart, prefix + "  ")

            # Process the message parts
            if 'parts' in payload:
                for part in payload['parts']:
                    process_parts(part)
            else:
                # Handle single-part message
                process_parts(payload)

            return {
                'subject': subject,
                'sender': sender,
                'body': body_text,
                'attachments': attachments
            }
            
        except Exception as e:
            print(f'An error occurred while reading message: {str(e)}')
            print(f'Error type: {type(e)}')
            import traceback
            traceback.print_exc()
            return None


    def download_attachment(self, message_id, attachment_id, filename):
        """Downloads an attachment from a message."""
        try:
            attachment = self.service.users().messages().attachments().get(
                userId='me',
                messageId=message_id,
                id=attachment_id
            ).execute()

            data = attachment['data']
            file_data = base64.urlsafe_b64decode(data)

            with open(filename, 'wb') as f:
                f.write(file_data)
                
            return True
        except Exception as e:
            print(f'An error occurred downloading attachment: {e}')
            return False


    def clean_email_body(self, body):
        """Removes quoted text from email body."""
        # Split the body into lines
        lines = body.split('\n')
        cleaned_lines = []
        
        for line in lines:
            # Stop when we hit a line starting with "On ... wrote:"
            if line.startswith('On ') and ' wrote:' in line:
                break
            # Skip quoted lines (starting with >)
            if not line.startswith('>'):
                cleaned_lines.append(line)
        
        # Remove trailing empty lines
        while cleaned_lines and not cleaned_lines[-1].strip():
            cleaned_lines.pop()
            
        return '\n'.join(cleaned_lines)

    def get_thread_messages(self, thread_id):
        """Gets all messages in a specific thread/conversation and offers to reply."""
        try:
            from email.utils import parsedate_to_datetime
            
            thread = self.service.users().threads().get(
                userId='me',
                id=thread_id
            ).execute()
            
            thread_messages = []
            for message in thread['messages']:
                headers = message['payload'].get('headers', [])
                
                header_dict = {}
                for header in headers:
                    header_dict[header['name'].lower()] = header['value']
                
                # Get message body
                body_text = ''
                if 'parts' in message['payload']:
                    for part in message['payload']['parts']:
                        if part['mimeType'] == 'text/plain':
                            if 'data' in part['body']:
                                body_text = base64.urlsafe_b64decode(
                                    part['body']['data']).decode('utf-8')
                else:
                    if 'body' in message['payload'] and 'data' in message['payload']['body']:
                        body_text = base64.urlsafe_b64decode(
                            message['payload']['body']['data']).decode('utf-8')
                
                # Clean the body text
                body_text = self.clean_email_body(body_text)
                
                # Parse the date
                date_str = header_dict.get('date', '')
                try:
                    parsed_date = parsedate_to_datetime(date_str)
                except:
                    parsed_date = datetime.datetime.fromtimestamp(int(message['internalDate'])/1000.0)
                
                thread_messages.append({
                    'message_id': header_dict.get('message-id', ''),
                    'subject': header_dict.get('subject', ''),
                    'sender': header_dict.get('from', ''),
                    'recipient': header_dict.get('to', ''),
                    'date_str': date_str,
                    'date': parsed_date,
                    'references': header_dict.get('references', ''),
                    'body': body_text,
                    'internal_date': int(message['internalDate'])
                })
            
            # Sort messages by parsed datetime
            thread_messages = sorted(thread_messages, key=lambda x: x['date'])
            
            # Display the conversation
            print("\n=== Email Thread ===")
            for i, msg in enumerate(thread_messages, 1):
                print(f"\nMessage {i} of {len(thread_messages)}")
                print(f"Date: {msg['date_str']}")
                print(f"From: {msg['sender']}")
                print(f"To: {msg['recipient']}")
                print(f"Subject: {msg['subject']}")
                print("\nBody:")
                print(msg['body'] if msg['body'].strip() else "(No content)")
                print("-" * 50)

            # Ask if user wants to reply
            while True:
                reply = input("\nWould you like to reply to this thread? (yes/no): ").lower()
                if reply in ['yes', 'no']:
                    break
                print("Please enter 'yes' or 'no'")
            
            if reply == 'yes':
                # Get the most recent message details for reply
                latest_msg = thread_messages[-1]
                reply_to = latest_msg['sender']
                subject = latest_msg['subject']
                if not subject.startswith('Re: '):
                    subject = 'Re: ' + subject
                
                # Build references string from all messages
                references = ' '.join(msg['message_id'] for msg in thread_messages if msg['message_id'])
                
                # Get reply message from user
                print("\nEnter your reply message (press Ctrl+D or Ctrl+Z on a new line when finished):")
                reply_lines = []
                try:
                    while True:
                        line = input()
                        reply_lines.append(line)
                except EOFError:
                    reply_text = '\n'.join(reply_lines)
                
                # Send the reply with all threading information
                self.send_message(
                    to=reply_to,
                    subject=subject,
                    message_text=reply_text,
                    thread_id=thread_id,
                    references=references,
                    in_reply_to=latest_msg['message_id']
                )
                print("\nReply sent successfully!")
                
            return thread_messages
            
        except Exception as e:
            print(f'An error occurred: {e}')
            import traceback
            traceback.print_exc()
            return []

    def send_message(self, to, subject, message_text, thread_id=None, references=None, in_reply_to=None):
        """Sends an email message with proper threading headers."""
        try:
            from email.mime.text import MIMEText
            from email.mime.multipart import MIMEMultipart
            import base64
            from email.utils import formatdate, make_msgid

            message = MIMEMultipart()
            message['to'] = to
            message['subject'] = subject
            message['date'] = formatdate(localtime=True)
            message['message-id'] = make_msgid()  # Generate a proper Message-ID
            
            # Add threading headers if available
            if references:
                message['References'] = references
            if in_reply_to:
                message['In-Reply-To'] = in_reply_to

            msg = MIMEText(message_text)
            message.attach(msg)
            
            # Debug print
            print("\nDEBUG - Message Details:")
            print(f"To: {message['to']}")
            print(f"Subject: {message['subject']}")
            print(f"Message-ID: {message['message-id']}")
            print(f"References: {message.get('References', 'None')}")
            print(f"In-Reply-To: {message.get('In-Reply-To', 'None')}")
            print(f"Thread ID: {thread_id}")
            
            raw_message = base64.urlsafe_b64encode(
                message.as_bytes()
            ).decode('utf-8')

            body = {'raw': raw_message}
            if thread_id:
                body['threadId'] = thread_id
            
            send_message = self.service.users().messages().send(
                userId='me', 
                body=body
            ).execute()
            
            print(f'\nMessage sent successfully. Message Id: {send_message["id"]}')
            print(f"Thread Id of sent message: {send_message.get('threadId', 'None')}")
            return send_message
        except Exception as e:
            print(f'An error occurred: {e}')
            import traceback
            traceback.print_exc()
            return None

    def display_recent_threads(self, query='', max_threads=3):
        """Displays recent email threads and allows user to select one to read and reply."""
        try:
            print("\nFetching recent threads...")
            messages = self.list_messages(query, max_results=max_threads)
            
            if not messages:
                print("No threads found.")
                return
            
            # Get unique thread IDs
            thread_ids = list(dict.fromkeys(msg['threadId'] for msg in messages))
            
            print("\nRecent Email Threads:")
            for i, thread_id in enumerate(thread_ids, 1):
                # Get the first message of each thread for preview
                thread = self.service.users().threads().get(userId='me', id=thread_id).execute()
                first_msg = thread['messages'][0]
                headers = first_msg['payload']['headers']
                subject = next((h['value'] for h in headers if h['name'] == 'Subject'), 'No Subject')
                sender = next((h['value'] for h in headers if h['name'] == 'From'), 'Unknown')
                
                print(f"\n{i}. Subject: {subject}")
                print(f"   From: {sender}")
            
            while True:
                try:
                    choice = input("\nEnter thread number to read (or 0 to exit): ")
                    if choice == '0':
                        return
                    choice = int(choice)
                    if 1 <= choice <= len(thread_ids):
                        self.get_thread_messages(thread_ids[choice-1])
                        break
                    else:
                        print("Invalid choice. Please try again.")
                except ValueError:
                    print("Please enter a valid number.")
                    
        except Exception as e:
            print(f'An error occurred: {e}')

def main():
    # Initialize and authenticate the Gmail client
    gmail_client = GmailClient()
    gmail_client.authenticate()

    # Display recent threads and handle interactions
    gmail_client.display_recent_threads(query='from:your_email@gmail.com', max_threads=7)
    # messages = gmail_client.list_messages('from:your_email@gmail.com', max_results=6)

    # #
    # if messages:
    #     # Get the first message's thread
    #     thread_messages = gmail_client.get_thread_messages(messages[0]['threadId'])
    #     print(messages[0]['threadId'])
        
    #     # Print all messages in the thread
    #     for msg in thread_messages:
    #         print(f"\nFrom: {msg['sender']}")
    #         # print(f"Date: {msg['date']}")
    #         print(f"Subject: {msg['subject']}")
    #         print(f"Body: {msg['body']}")
    #         print("-" * 50)


if __name__ == '__main__':
    main()

"""
virtualenv env_name
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
.\env_name\Scripts\activate
where python3
python main.py
"""

                                                            
