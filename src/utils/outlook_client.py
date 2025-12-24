"""High-performance Outlook client for mailbox access and email search."""

import win32com.client
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional, Tuple
import logging
import pythoncom
import re
import time
import threading
from functools import lru_cache
from concurrent.futures import ThreadPoolExecutor, as_completed
import queue
import uuid

from ..config.config_reader import config

logging.basicConfig(
    level=logging.WARNING,
    format="%(asctime)s - %(levelname)s - %(message)s",
    force=True  # <- overrides any prior basicConfig
)
logger = logging.getLogger(__name__)

class OutlookClient:
    """High-performance client for accessing Outlook mailboxes with optimized search."""
    
    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.connected = False
        self._search_cache = {}  # Cache for search results
        self._folder_cache = {}  # Cache for folder references
        self._shared_recipient_cache = None  # Cache for resolved shared recipient
        self._connection_retry_count = 0
        self._max_retries = config.get_int('max_connection_retries', 3)
    
    def connect(self, retry_attempt: int = 0) -> bool:
        """Connect to Outlook application with retry logic."""
        try:
            logger.info("Connecting to Outlook...")
            start_time = time.time()
            
            # Initialize COM for thread
            pythoncom.CoInitialize()
            
            # Try to connect to existing Outlook instance first (much faster)
            try:
                self.outlook = win32com.client.GetActiveObject("Outlook.Application")
                logger.info("Connected to existing Outlook instance")
            except:
                # Fall back to creating new instance
                logger.info("No existing Outlook instance, launching new one...")
                self.outlook = win32com.client.Dispatch("Outlook.Application")
            
            self.namespace = self.outlook.GetNamespace("MAPI")
            
            # Try Extended MAPI approach to potentially reduce security prompts (if enabled)
            if config.get_bool('use_extended_mapi_login', True):
                try:
                    logger.info("Attempting Extended MAPI login to reduce security prompts...")
                    self.namespace.Logon(None, None, False, True)
                    logger.info("Extended MAPI login successful")
                except Exception as logon_error:
                    logger.warning(f"Extended MAPI login failed: {logon_error}")
            
            self.connected = True
            connection_time = time.time() - start_time
            logger.info(f"Successfully connected to Outlook in {connection_time:.2f} seconds")
            return True
        except Exception as e:
            logger.error(f"Failed to connect to Outlook (attempt {retry_attempt + 1}): {e}")
            self.connected = False
            
            # Implement exponential backoff retry
            if retry_attempt < self._max_retries - 1:
                wait_time = (2 ** retry_attempt) * 1  # 1s, 2s, 4s
                logger.info(f"Retrying connection in {wait_time} seconds...")
                time.sleep(wait_time)
                return self.connect(retry_attempt + 1)
            
            return False
    
    def check_access(self) -> Dict[str, Any]:
        """Check access to personal and shared mailboxes."""
        if not self.connected:
            if not self.connect():
                return {"error": "Could not connect to Outlook"}
        
        result = {
            "outlook_connected": True,
            "personal_accessible": False,
            "shared_accessible": False,
            "shared_configured": bool(config.get('shared_mailbox_email')),
            "retention_personal_months": config.get_int('personal_retention_months', 6),
            "retention_shared_months": config.get_int('shared_retention_months', 12),
            "errors": []
        }
        
        # Test personal mailbox
        try:
            personal_inbox = self.namespace.GetDefaultFolder(6)  # 6 = Inbox
            if personal_inbox:
                result["personal_accessible"] = True
                result["personal_name"] = self._get_store_display_name(personal_inbox)
        except Exception as e:
            result["errors"].append(f"Personal mailbox error: {str(e)}")
        
        # Test shared mailbox
        shared_email = config.get('shared_mailbox_email')
        if shared_email:
            try:
                # Use cached recipient if available
                if not self._shared_recipient_cache:
                    self._shared_recipient_cache = self.namespace.CreateRecipient(shared_email)
                    self._shared_recipient_cache.Resolve()
                
                if self._shared_recipient_cache.Resolved:
                    shared_inbox = self.namespace.GetSharedDefaultFolder(self._shared_recipient_cache, 6)
                    if shared_inbox:
                        result["shared_accessible"] = True
                        result["shared_name"] = self._get_store_display_name(shared_inbox)
            except Exception as e:
                result["errors"].append(f"Shared mailbox error: {str(e)}")
                self._shared_recipient_cache = None  # Clear cache on error
        
        return result
    
    def search_emails(self, search_text: str, 
                     include_personal: bool = True, 
                     include_shared: bool = True) -> List[Dict[str, Any]]:
        """Search emails in both subject and body using exact phrase matching with parallel execution."""
        if not self.connected:
            if not self.connect():
                return []
        
        # Enhanced cache key including max_results
        max_results = config.get_int('max_search_results', 500)
        cache_key = f"{search_text}_{include_personal}_{include_shared}_{max_results}"
        
        if cache_key in self._search_cache:
            # Check cache age (simple time-based invalidation)
            cache_entry = self._search_cache[cache_key]
            if time.time() - cache_entry['timestamp'] < 3600:  # 1 hour cache
                logger.info(f"Returning cached results for '{search_text}'")
                return cache_entry['data']
        
        all_emails = []
        
        # Use parallel search for multiple mailboxes
        if include_personal and include_shared and config.get('shared_mailbox_email'):
            # Parallel search using threading
            with ThreadPoolExecutor(max_workers=2) as executor:
                futures = []
                
                # Submit personal mailbox search
                if include_personal:
                    futures.append(
                        executor.submit(
                            self._search_mailbox_wrapper,
                            'personal',
                            search_text,
                            max_results
                        )
                    )
                
                # Submit shared mailbox search
                if include_shared:
                    futures.append(
                        executor.submit(
                            self._search_mailbox_wrapper,
                            'shared',
                            search_text,
                            max_results
                        )
                    )
                
                # Collect results
                for future in as_completed(futures):
                    try:
                        emails = future.result()
                        all_emails.extend(emails)
                    except Exception as e:
                        logger.error(f"Error in parallel search: {e}")
        else:
            # Sequential search for single mailbox
            if include_personal:
                personal_emails = self._search_mailbox_comprehensive(
                    self.namespace.GetDefaultFolder(6), 
                    search_text, 
                    'personal',
                    max_results
                )
                all_emails.extend(personal_emails)
                logger.info(f"Found {len(personal_emails)} emails in personal mailbox")
            
            if include_shared and config.get('shared_mailbox_email'):
                try:
                    # Use cached recipient if available
                    if not self._shared_recipient_cache:
                        shared_email = config.get('shared_mailbox_email')
                        self._shared_recipient_cache = self.namespace.CreateRecipient(shared_email)
                        self._shared_recipient_cache.Resolve()
                    
                    if self._shared_recipient_cache.Resolved:
                        shared_inbox = self.namespace.GetSharedDefaultFolder(self._shared_recipient_cache, 6)
                        shared_emails = self._search_mailbox_comprehensive(
                            shared_inbox,
                            search_text,
                            'shared',
                            max_results - len(all_emails)
                        )
                        all_emails.extend(shared_emails)
                        logger.info(f"Found {len(shared_emails)} emails in shared mailbox")
                except Exception as e:
                    logger.error(f"Error searching shared mailbox: {e}")
                    self._shared_recipient_cache = None
        
        # Sort by received time (newest first)
        all_emails.sort(key=lambda x: x.get('received_time', datetime.min), reverse=True)
        
        # Cache results with timestamp
        limited_results = all_emails[:max_results]
        self._search_cache[cache_key] = {
            'data': limited_results,
            'timestamp': time.time()
        }
        
        # Limit cache size
        if len(self._search_cache) > 100:
            # Remove oldest entries
            oldest_key = min(self._search_cache.keys(), 
                           key=lambda k: self._search_cache[k].get('timestamp', 0))
            del self._search_cache[oldest_key]
        
        return limited_results
    
    def search_emails_by_subject(self, subject: str, 
                                include_personal: bool = True, 
                                include_shared: bool = True) -> List[Dict[str, Any]]:
        """Legacy method - redirects to search_emails for backward compatibility."""
        return self.search_emails(subject, include_personal, include_shared)

    def _search_mailbox_wrapper(self, mailbox_type: str, search_text: str, 
                                max_results: int) -> List[Dict[str, Any]]:
        """Wrapper for parallel mailbox search with proper per-thread COM usage."""
        # Explicit STA init for this thread
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        try:
            # Create Outlook COM objects *in this thread*
            outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
            session = outlook.Session  # same as outlook.GetNamespace("MAPI")

            if mailbox_type == 'personal':
                inbox = session.GetDefaultFolder(6)  # olFolderInbox
                return self._search_mailbox_comprehensive(
                    inbox, search_text, 'personal', max_results
                )

            elif mailbox_type == 'shared':
                # Don't reuse a cached Recipient across threads.
                shared_email = config.get('shared_mailbox_email')
                recip = session.CreateRecipient(shared_email)
                recip.Resolve()
                if recip.Resolved:
                    shared_inbox = session.GetSharedDefaultFolder(recip, 6)  # olFolderInbox
                    return self._search_mailbox_comprehensive(
                        shared_inbox, search_text, 'shared', max_results
                    )
                else:
                    logger.error(f"Could not resolve shared recipient: {shared_email}")
                    return []

            return []

        except Exception as e:
            logger.error(f"Error in mailbox wrapper for {mailbox_type}: {e}")
            return []
        finally:
            pythoncom.CoUninitialize()

    def _search_mailbox_comprehensive(self, inbox_folder, search_text: str,
                                      mailbox_type: str, max_results: int):
        emails = []
        found_ids = set()

        app = inbox_folder.Application  # keep COM objects on this thread

        # ---- Build Scope safely (single-quoted, apostrophes doubled) ----
        folder_path = inbox_folder.FolderPath or ""
        scope = self._get_english_folder_path(inbox_folder)

        # ---- Build Filter safely (DASL with @SQL=) ----
        esc = (search_text or "").replace("'", "''")  # single quotes doubled for literals
        query = (
            "@SQL="
            "\"urn:schemas:httpmail:subject\" ci_phrasematch '{t}' "
            "OR "
            "\"urn:schemas:httpmail:textdescription\" ci_phrasematch '{t}'"
        ).format(t=esc)

        # (Optional) keep tags reasonably short; some environments are picky
        tag = "EmailBodySearch-" + str(uuid.uuid4())[:8]

        logger.info("AdvancedSearch Scope=%s Filter=%s", scope, query)

        try:
            # ---- Call positionally to avoid named-arg binding issues ----
            # Signature: AdvancedSearch(Scope, Filter, SearchSubFolders, Tag)
            search = app.AdvancedSearch(scope, query, False, tag)

            # Poll until results stabilize or timeout
            start = time.time()
            last = -1
            stable = 0.0
            while True:
                results = search.Results  # same-thread COM access
                count = results.Count
                if count == last:
                    stable += 0.1
                else:
                    stable = 0.0
                    last = count
                if stable >= 0.5 or (time.time() - start) > 30:
                    break
                time.sleep(0.1)

            take = min(results.Count, max_results)
            logger.info("AdvancedSearch returned %d (taking %d)", results.Count, take)

            for i in range(1, take + 1):
                try:
                    item = results.Item(i)
                    entry_id = getattr(item, 'EntryID', '')
                    if entry_id and entry_id not in found_ids:
                        email_data = self._extract_email_data(item, inbox_folder.Name, mailbox_type)
                        if email_data:
                            emails.append(email_data)
                            found_ids.add(entry_id)
                except Exception as e:
                    logger.debug("Error processing result %d: %s", i, e)
                    continue

            # Clean up the running search (optional but tidy)
            try:
                app.AdvancedSearchStop(search.SearchID)
            except Exception:
                pass

        except Exception as e:

            logger.info("AdvancedSearch failed: %s", e)
            logger.info("Falling back to Restrict() search")

            # -- Fallback: subject LIKE (fast) --
            try:
                items = inbox_folder.Items
                items.Sort("[ReceivedTime]", True)
                esc_like = (search_text or "").replace("'", "''")
                subject_filter = (
                    "@SQL=\"urn:schemas:httpmail:subject\" LIKE '%" + esc_like + "%'"
                )
                for item in items.Restrict(subject_filter):
                    if len(emails) >= max_results:
                        break
                    entry_id = getattr(item, 'EntryID', '')
                    if entry_id and entry_id not in found_ids:
                        email_data = self._extract_email_data(item, inbox_folder.Name, mailbox_type)
                        if email_data:
                            emails.append(email_data)
                            found_ids.add(entry_id)
            except Exception as fallback_error:
                logger.error("Fallback subject filter failed: %s", fallback_error)

        # Optional: search sibling folders
        if len(emails) < max_results and config.get_bool('search_all_folders', True):
            try:
                more = self._search_other_folders(
                    inbox_folder.Parent, search_text, mailbox_type,
                    max_results - len(emails), found_ids
                )
                emails.extend(more)
            except Exception as e:
                logger.error("Error searching other folders: %s", e)

        return emails
    
    def _search_other_folders(self, store, search_text: str, mailbox_type: str, 
                             max_results: int, found_ids: set) -> List[Dict[str, Any]]:
        """Search other folders using AdvancedSearch for consistency."""
        emails = []
        key_folders = ['Sent Items', 'Drafts']  # Extend if needed
        
        for folder_name in key_folders:
            if len(emails) >= max_results:
                break
            
            try:
                folder = self._get_folder_by_name(store, folder_name)
                if folder:
                    # Use AdvancedSearch for this folder as well
                    scope = f"'{folder.FolderPath}'"
                    search_text_escaped = search_text.replace('"', '""')
                    query = (
                        f'urn:schemas:httpmail:subject ci_phrasematch "{search_text_escaped}" OR '
                        f'urn:schemas:httpmail:textdescription ci_phrasematch "{search_text_escaped}"'
                    )
                    
                    logger.info(f"AdvancedSearch in {folder_name} for '{search_text}'")
                    
                    search = self.outlook.AdvancedSearch(
                        Scope=scope, 
                        Filter=query, 
                        SearchSubFolders=False, 
                        Tag=f"OtherFolderSearch_{folder_name}"
                    )
                    
                    # Poll with shorter timeout for secondary folders
                    start_time = time.time()
                    while not search.SearchComplete:
                        time.sleep(0.1)
                        if time.time() - start_time > 10:  # Shorter timeout for secondary folders
                            break
                    
                    if search.SearchComplete:
                        results = search.Results
                        result_count = min(results.Count, max_results - len(emails))
                        
                        for i in range(1, result_count + 1):
                            item = results.Item(i)
                            entry_id = getattr(item, 'EntryID', '')
                            if entry_id and entry_id not in found_ids:
                                email_data = self._extract_email_data(item, folder_name, mailbox_type)
                                if email_data:
                                    emails.append(email_data)
                                    found_ids.add(entry_id)
            except Exception as e:
                logger.debug(f"Error searching {folder_name}: {e}")
        
        return emails
    
    def _extract_email_data(self, item, folder_name: str, 
                           mailbox_type: str) -> Dict[str, Any]:
        """Extract email data with optimized body and recipient handling."""
        try:
            # Get the full email body
            body = getattr(item, 'Body', '')
            
            # Apply max_body_chars if configured (0 means no limit)
            max_body_chars = config.get_int('max_body_chars', 0)
            if max_body_chars > 0 and len(body) > max_body_chars:
                body = body[:max_body_chars] + " [truncated]"
            
            # Clean HTML if configured
            if config.get_bool('clean_html_content', True) and body:
                body = self._clean_html(body)
            
            # Get recipients list with limit for performance
            recipients = []
            max_recipients = config.get_int('max_recipients_display', 10)
            try:
                recipient_count = 0
                for recipient in item.Recipients:
                    if recipient_count >= max_recipients:
                        recipients.append(f"... and {item.Recipients.Count - recipient_count} more")
                        break
                    recipients.append(getattr(recipient, 'Name', getattr(recipient, 'Address', '')))
                    recipient_count += 1
            except:
                pass
            
            email_data = {
                'subject': getattr(item, 'Subject', 'No Subject'),
                'sender_name': getattr(item, 'SenderName', 'Unknown'),
                'sender_email': getattr(item, 'SenderEmailAddress', ''),
                'recipients': recipients,
                'received_time': getattr(item, 'ReceivedTime', datetime.now()),
                'folder_name': folder_name,
                'mailbox_type': mailbox_type,
                'importance': getattr(item, 'Importance', 1),
                'body': body,  # Full body for summarization
                'size': getattr(item, 'Size', 0),
                'attachments_count': getattr(item.Attachments, 'Count', 0) if hasattr(item, 'Attachments') else 0,
                'unread': getattr(item, 'Unread', False),
                'entry_id': getattr(item, 'EntryID', '')
            }
            
            # Release COM reference to free memory
            item = None
            
            return email_data
        except Exception as e:
            logger.error(f"Error extracting email data: {e}")
            return None
    
    def _get_store_display_name(self, folder) -> str:
        """Safely get store display name from a folder."""
        try:
            if hasattr(folder, 'Parent'):
                parent = folder.Parent
                if hasattr(parent, 'DisplayName'):
                    return parent.DisplayName
                elif hasattr(parent, 'Name'):
                    return parent.Name
            return "Mailbox"
        except:
            return "Mailbox"
    
    def _get_folder_by_name(self, store, name: str):
        """Get folder by name from cache or store."""
        cache_key = f"{id(store)}_{name}"
        
        if cache_key in self._folder_cache:
            return self._folder_cache[cache_key]
        
        try:
            for folder in store.GetRootFolder().Folders:
                if folder.Name.lower() == name.lower():
                    self._folder_cache[cache_key] = folder
                    return folder
        except:
            pass
        
        return None
    
    def _clean_html(self, text: str) -> str:
        """Clean HTML from email body."""
        import re
        
        # Remove HTML tags
        text = re.sub(r'<[^>]+>', '', text)
        
        # Decode common HTML entities
        html_entities = {
            '&amp;': '&',
            '&lt;': '<',
            '&gt;': '>',
            '&quot;': '"',
            '&#39;': "'",
            '&nbsp;': ' '
        }
        
        for entity, char in html_entities.items():
            text = text.replace(entity, char)
        
        # Clean up whitespace
        text = re.sub(r'\s+', ' ', text).strip()
        
        return text

    def _get_english_folder_path(self, folder):
        # Convert localized names (like "Posteingang") to internal English names
        # Works because MAPIFolder.FolderPath always uses localized names
        path = folder.FolderPath
        # Normalize localized root paths for common locales
        replacements = {
            "\\Posteingang": "\\Inbox",
            "\\Gesendete Elemente": "\\Sent Items",
            "\\Entwürfe": "\\Drafts",
            "\\Gelöschte Elemente": "\\Deleted Items",
            "\\Junk-E-Mail": "\\Junk Email",
        }
        for de, en in replacements.items():
            if path.endswith(de):
                path = path.replace(de, en)
        return f"'{path.replace("'", "''")}'"

# Global client instance
outlook_client = OutlookClient()
