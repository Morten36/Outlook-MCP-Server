import win32com.client
from datetime import datetime, timedelta
import pytz

def access_shared_mailbox():
    """
    Access shared mailboxes and get today's emails
    """
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Method 1: Access shared mailbox by email address
        shared_email = "someone@example.com"  # Your actual shared mailbox
        
        try:
            # Get the shared mailbox folder
            shared_mailbox = namespace.CreateRecipient(shared_email)
            shared_mailbox.Resolve()
            
            if shared_mailbox.Resolved:
                print(f"[OK] Found shared mailbox: {shared_email}")
                
                # Get the shared mailbox inbox
                shared_inbox = namespace.GetSharedDefaultFolder(shared_mailbox, 6)  # 6 = Inbox
                print(f"[INFO] Shared inbox accessed: {shared_inbox.Name}")
                
                # Get today's emails
                today_emails = get_todays_emails(shared_inbox)
                return today_emails
                
        except Exception as e:
            print(f"Method 1 failed: {e}")
            
            # Method 2: Access through Stores (if shared mailbox is mounted)
            print("Trying Method 2: Accessing through Stores...")
            return access_through_stores(namespace, shared_email)
            
    except Exception as e:
        print(f"Error accessing shared mailbox: {e}")
        return []

def access_through_stores(namespace, shared_email):
    """
    Access shared mailbox through Stores
    """
    try:
        stores = namespace.Stores
        print(f"\n[STORES] Available Stores ({stores.Count}):")
        
        for i in range(1, stores.Count + 1):
            store = stores[i]
            print(f"   {i}. {store.DisplayName}")
            
            # Check if this is the shared mailbox we want
            if "accessfx" in store.DisplayName.lower() or "escalation" in store.DisplayName.lower():
                print(f"[OK] Found shared mailbox store: {store.DisplayName}")
                
                # Get inbox from this store
                root_folder = store.GetRootFolder()
                inbox = root_folder.Folders["Inbox"]
                
                today_emails = get_todays_emails(inbox)
                return today_emails
                
        return []
        
    except Exception as e:
        print(f"Error in stores method: {e}")
        return []

def get_todays_emails(inbox):
    """
    Get emails from today and return relevant information - FIXED datetime comparison
    """
    try:
        # Get today's date (timezone-naive)
        today = datetime.now().date()
        print(f"[SEARCH] Searching for emails from {today}...")
        
        # Get messages
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)  # Sort by newest first
        
        today_emails = []
        processed_count = 0
        
        # Filter today's emails - FIXED approach
        for i in range(1, min(messages.Count + 1, 100)):  # Limit to first 100 for performance
            try:
                message = messages[i]
                received_time = message.ReceivedTime
                
                # Convert Outlook datetime to Python date for comparison (removes timezone issue)
                if hasattr(received_time, 'date'):
                    received_date = received_time.date()
                else:
                    # Fallback: convert to datetime first
                    received_date = datetime.fromtimestamp(received_time.timestamp()).date()
                
                # Compare dates only (not times) - this avoids timezone issues
                if received_date == today:
                    email_data = {
                        'subject': getattr(message, 'Subject', 'No Subject'),
                        'sender': getattr(message, 'SenderName', 'Unknown Sender'),
                        'sender_email': getattr(message, 'SenderEmailAddress', ''),
                        'received_time': received_time,
                        'body': getattr(message, 'Body', '')[:1000],  # First 1000 chars
                        'importance': getattr(message, 'Importance', 1),
                        'categories': getattr(message, 'Categories', ''),
                        'size': getattr(message, 'Size', 0),
                        'attachments': message.Attachments.Count if hasattr(message, 'Attachments') else 0
                    }
                    today_emails.append(email_data)
                
                processed_count += 1
                
                # Stop if we've gone past today (since emails are sorted by newest first)
                if received_date < today:
                    break
                    
            except Exception as e:
                print(f"Error processing message {i}: {e}")
                continue
        
        print(f"[RESULT] Found {len(today_emails)} emails from today (processed {processed_count} messages)")
        return today_emails
        
    except Exception as e:
        print(f"Error getting today's emails: {e}")
        return []

def summarize_issues(emails):
    """
    Enhanced issue summarization
    """
    if not emails:
        print("No emails to summarize")
        return
    
    print(f"\n[SUMMARY] TODAY'S EMAIL SUMMARY ({len(emails)} emails)")
    print("=" * 60)
    
    # Enhanced keywords for better classification
    urgent_keywords = ['urgent', 'critical', 'down', 'outage', 'error', 'failure', 'issue', 'problem', 'alert', 'incident']
    warning_keywords = ['warning', 'caution', 'attention', 'investigate', 'check', 'review']
    
    urgent_emails = []
    warning_emails = []
    normal_emails = []
    
    for email in emails:
        subject_lower = email['subject'].lower()
        body_lower = email['body'].lower()
        
        is_urgent = any(keyword in subject_lower or keyword in body_lower 
                       for keyword in urgent_keywords)
        is_warning = any(keyword in subject_lower or keyword in body_lower 
                        for keyword in warning_keywords)
        
        if is_urgent or email['importance'] > 1:
            urgent_emails.append(email)
        elif is_warning:
            warning_emails.append(email)
        else:
            normal_emails.append(email)
    
    # Print urgent emails
    if urgent_emails:
        print(f"\n[URGENT] URGENT/CRITICAL EMAILS ({len(urgent_emails)}):")
        for i, email in enumerate(urgent_emails, 1):
            print(f"{i}. Subject: {email['subject']}")
            print(f"   From: {email['sender']}")
            print(f"   Time: {email['received_time']}")
            print(f"   Attachments: {email['attachments']}")
            print(f"   Preview: {email['body'][:150]}...")
            print("-" * 40)
    
    # Print warning emails (top 3)
    if warning_emails:
        print(f"\n[WARNING] WARNING EMAILS ({len(warning_emails)}) - Showing top 3:")
        for i, email in enumerate(warning_emails[:3], 1):
            print(f"{i}. {email['subject'][:60]}...")
            print(f"   From: {email['sender']}")
            print(f"   Time: {email['received_time']}")
    
    # Print summary stats
    print(f"\n[STATS] SUMMARY STATS:")
    print(f"   Total Emails: {len(emails)}")
    print(f"   Urgent/Critical: {len(urgent_emails)}")
    print(f"   Warnings: {len(warning_emails)}")
    print(f"   Normal: {len(normal_emails)}")
    
    # Top senders
    senders = {}
    for email in emails:
        sender = email['sender']
        senders[sender] = senders.get(sender, 0) + 1
    
    print(f"\n[SENDERS] TOP SENDERS TODAY:")
    for sender, count in sorted(senders.items(), key=lambda x: x[1], reverse=True)[:5]:
        print(f"   {sender}: {count} emails")
    
    # Attachment summary
    total_attachments = sum(email['attachments'] for email in emails)
    emails_with_attachments = sum(1 for email in emails if email['attachments'] > 0)
    
    if total_attachments > 0:
        print(f"\n[ATTACHMENTS] ATTACHMENTS:")
        print(f"   Total attachments: {total_attachments}")
        print(f"   Emails with attachments: {emails_with_attachments}")

# Alternative method using Outlook's filtering (might be more efficient)
def get_todays_emails_with_filter(inbox):
    """
    Alternative method using Outlook's built-in filtering
    """
    try:
        today_str = datetime.now().strftime('%m/%d/%Y')
        
        # Use Outlook's Restrict method to filter emails
        filter_criteria = f"[ReceivedTime] >= '{today_str} 12:00 AM' AND [ReceivedTime] < '{datetime.now() + timedelta(days=1):%m/%d/%Y} 12:00 AM'"
        
        print(f"[FILTER] Using filter: {filter_criteria}")
        
        messages = inbox.Items
        filtered_messages = messages.Restrict(filter_criteria)
        
        print(f"[RESULT] Found {filtered_messages.Count} emails from today using filter")
        
        today_emails = []
        for i in range(1, filtered_messages.Count + 1):
            try:
                message = filtered_messages[i]
                email_data = {
                    'subject': getattr(message, 'Subject', 'No Subject'),
                    'sender': getattr(message, 'SenderName', 'Unknown Sender'),
                    'received_time': message.ReceivedTime,
                    'body': getattr(message, 'Body', '')[:1000],
                    'importance': getattr(message, 'Importance', 1),
                    'attachments': message.Attachments.Count if hasattr(message, 'Attachments') else 0
                }
                today_emails.append(email_data)
            except Exception as e:
                print(f"Error processing filtered message {i}: {e}")
                continue
        
        return today_emails
        
    except Exception as e:
        print(f"Filter method failed: {e}")
        return []

# Main execution
if __name__ == "__main__":
    print("[START] Starting Shared Mailbox Access")
    print("=" * 50)
    
    emails = access_shared_mailbox()
    
    if emails:
        summarize_issues(emails)
    else:
        print("[ERROR] No emails found or unable to access shared mailbox")
