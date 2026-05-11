"""Email operations tools for Outlook MCP Server."""

from typing import Dict, Any, Union, List, Optional
from ..backend.email_composition import (
    reply_to_email_by_number,
    compose_email,
    save_draft_email,
    save_reply_draft_by_number,
    send_draft_by_entry_id,
    delete_draft_by_entry_id,
)
from ..backend.outlook_session import OutlookSessionManager
from ..backend.validation import ValidationError


def reply_to_email_by_number_tool(
    email_number: int, 
    reply_text: str, 
    to_recipients: Union[str, List[str], None] = None, 
    cc_recipients: Union[str, List[str], None] = None
) -> Dict[str, Any]:
    """Reply to an email with custom recipients if provided

    Args:
        email_number: Email's position in the last listing
        reply_text: Text to prepend to the reply
        to_recipients: Either a single email string OR a list of email strings (None preserves original recipients)
                      Examples: "user@company.com" OR ["user@company.com", "boss@company.com"]
        cc_recipients: Either a single email string OR a list of email strings (None preserves original recipients)
                      Examples: "user@company.com" OR ["user@company.com", "boss@company.com"]

    Behavior:
        - When both to_recipients and cc_recipients are None:
          * Uses ReplyAll() to maintain original recipients
        - When either parameter is provided:
          * Uses Reply() with specified recipients
          * Any None parameters will result in empty recipient fields
        - Single email strings and lists of email strings are both accepted

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Confirmation message here"
        }
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValidationError("Email number must be a positive integer")
    if not reply_text or not isinstance(reply_text, str):
        raise ValidationError("Reply text must be a non-empty string")
    
    try:
        result = reply_to_email_by_number(email_number, reply_text, to_recipients, cc_recipients)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error replying to email: {str(e)}"}


def compose_email_tool(recipient_email: str, subject: str, body: str, cc_email: Optional[str] = None) -> Dict[str, Any]:
    """Compose and send a new email

    Args:
        recipient_email: Email address(es) of the recipient(s) - can be single email or semicolon-separated list
        subject: Subject line of the email
        body: Main content of the email
        cc_email: Optional CC email address(es) - can be single email or semicolon-separated list

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Confirmation message here"
        }
    """
    if not recipient_email or not isinstance(recipient_email, str):
        raise ValidationError("Recipient email must be a non-empty string")
    if not subject or not isinstance(subject, str):
        raise ValidationError("Subject must be a non-empty string")
    if not body or not isinstance(body, str):
        raise ValidationError("Body must be a non-empty string")
    
    try:
        # Parse semicolon-separated email addresses into lists
        to_recipients = [email.strip() for email in recipient_email.split(';') if email.strip()]
        cc_recipients = None
        if cc_email:
            cc_recipients = [email.strip() for email in cc_email.split(';') if email.strip()]
        
        result = compose_email(to_recipients, subject, body, cc_recipients)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error composing email: {str(e)}"}


def save_draft_email_tool(
    recipient_email: str,
    subject: str,
    body: str,
    cc_email: Optional[str] = None,
    html: bool = False,
) -> Dict[str, Any]:
    """Compose a new email and save it as a DRAFT in the Drafts folder (does NOT send).

    Args:
        recipient_email: Recipient email address(es), single or semicolon-separated.
        subject: Email subject.
        body: Email body content (plain text or HTML).
        cc_email: Optional CC email address(es), single or semicolon-separated.
        html: If True, body is treated as HTML.

    Returns:
        dict: {"type": "text", "text": "Draft saved to Drafts folder. EntryID: ..."}
    """
    if not recipient_email or not isinstance(recipient_email, str):
        raise ValidationError("Recipient email must be a non-empty string")
    if not subject or not isinstance(subject, str):
        raise ValidationError("Subject must be a non-empty string")
    if not body or not isinstance(body, str):
        raise ValidationError("Body must be a non-empty string")

    try:
        to_recipients = [e.strip() for e in recipient_email.split(';') if e.strip()]
        cc_recipients = None
        if cc_email:
            cc_recipients = [e.strip() for e in cc_email.split(';') if e.strip()]
        result = save_draft_email(to_recipients, subject, body, cc_recipients, html=html)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error saving draft: {str(e)}"}


def save_reply_draft_by_number_tool(
    email_number: int,
    reply_text: str,
    to_recipients: Union[str, List[str], None] = None,
    cc_recipients: Union[str, List[str], None] = None,
    html: bool = False,
) -> Dict[str, Any]:
    """Create a reply draft for an email in the cache and save it to the Drafts folder (does NOT send).

    Uses Outlook's native Reply()/ReplyAll() so the original quoted thread is preserved
    exactly as Outlook would render it. The reply_text is prepended at the top.

    Args:
        email_number: Email's position in the cache (1-based)
        reply_text: Text to prepend (HTML if html=True)
        to_recipients: Override To (None = ReplyAll defaults). String or list of strings.
        cc_recipients: Override CC. String or list of strings.
        html: If True, reply_text is treated as HTML.

    Returns:
        dict: {"type": "text", "text": "Reply draft saved to Drafts folder. EntryID: ..."}
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValidationError("Email number must be a positive integer")
    if not reply_text or not isinstance(reply_text, str):
        raise ValidationError("Reply text must be a non-empty string")

    try:
        result = save_reply_draft_by_number(email_number, reply_text, to_recipients, cc_recipients, html=html)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error saving reply draft: {str(e)}"}


def send_draft_by_entry_id_tool(entry_id: str) -> Dict[str, Any]:
    """Open an existing draft (by EntryID) and send it.

    Use this AFTER save_draft_email_tool / save_reply_draft_by_number_tool
    once the user has explicitly confirmed (e.g. "send it"). This sends the
    EXACT draft the user reviewed (HTML, signature, footer, quoted thread).

    Args:
        entry_id: Outlook EntryID returned by the save_*_draft tools.
    """
    if not entry_id or not isinstance(entry_id, str):
        raise ValidationError("entry_id must be a non-empty string")
    try:
        result = send_draft_by_entry_id(entry_id)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error sending draft: {str(e)}"}


def delete_draft_by_entry_id_tool(entry_id: str) -> Dict[str, Any]:
    """Delete a draft (by EntryID). Moves it to Deleted Items.

    Use when the user cancels/discards a previously saved draft.

    Args:
        entry_id: Outlook EntryID of the draft to delete.
    """
    if not entry_id or not isinstance(entry_id, str):
        raise ValidationError("entry_id must be a non-empty string")
    try:
        result = delete_draft_by_entry_id(entry_id)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error deleting draft: {str(e)}"}


def move_email_tool(email_number: int, target_folder_name: str) -> Dict[str, Any]:
    """Move an email to the specified folder.

    Args:
        email_number: The number of the email in the cache to move (1-based)
        target_folder_name: Name or path of the target folder (supports nested paths like "user@company.com/Inbox/SubFolder1/SubFolder2")

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Email moved successfully to target_folder"
        }

    Note:
        Requires emails to be loaded first via list_recent_emails or search_emails.
        After moving, the cache will be cleared to reflect the new email positions.
        
        IMPORTANT: Target folder paths must include the email address as the root folder.
        Use format: "user@company.com/Inbox/SubFolder" not just "Inbox/SubFolder"
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValidationError("Email number must be a positive integer")
    if not target_folder_name or not isinstance(target_folder_name, str):
        raise ValidationError("Target folder name must be a non-empty string")

    try:
        # Use direct email operations instead of session manager wrapper
        from ..backend.outlook_session.email_operations import move_email_to_folder
        result = move_email_to_folder(email_number, target_folder_name)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error moving email: {str(e)}"}


def delete_email_by_number_tool(email_number: int) -> Dict[str, Any]:
    """Move an email to the Deleted Items folder.

    Args:
        email_number: The number of the email in the cache to delete (1-based)

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Email moved to Deleted Items successfully"
        }

    Note:
        Requires emails to be loaded first via list_recent_emails or search_emails.
        This tool moves the email to the Deleted Items folder instead of permanently deleting it.
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValidationError("Email number must be a positive integer")

    try:
        # Use direct email operations instead of session manager wrapper
        from ..backend.outlook_session.email_operations import delete_email_by_number
        result = delete_email_by_number(email_number)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error deleting email: {str(e)}"}