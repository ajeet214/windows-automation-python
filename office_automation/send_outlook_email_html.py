"""
Send an HTML email via Outlook with attachments and optional inline images.
Requires: pywin32
"""

from __future__ import annotations

import logging
import time
from pathlib import Path
from typing import Iterable, Mapping, Optional

import win32com.client as win32  # pywin32

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

# Outlook MAPI property for content-id on attachments (used for cid: inline images)
PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"


def _as_path_list(items: Optional[Iterable[str | Path]]) -> list[Path]:
    """Normalize an iterable of path-like values to a list of `Path`."""
    if not items:
        return []
    return [Path(p) for p in items if p]


def _map_importance(value: int | str | None) -> int:
    """
    Map importance input to Outlook numeric values:
      0=Low, 1=Normal, 2=High
    """
    if value is None:
        return 1
    if isinstance(value, int):
        return 2 if value > 1 else (0 if value < 1 else 1)
    v = str(value).strip().lower()
    if v in {"low", "0"}:
        return 0
    if v in {"high", "2"}:
        return 2
    return 1


def send_html_email(
    to: str,
    subject: str,
    html_body: str,
    attachments: Optional[Iterable[str | Path]] = None,
    inline_image: Optional[str | Path] = None,
    *,
    # New optional features:
    inline_cid: str = "banner",
    inline_placeholder: str = "{{INLINE_CID}}",
    inline_images: Optional[Mapping[str, str | Path]] = None,  # {cid: file_path}
    cc: Optional[str] = None,
    bcc: Optional[str] = None,
    reply_to: Optional[str] = None,
    importance: int | str | None = None,
    display_before_send: bool = True,
    send_immediately: bool = True,
    paste_delay: float = 0.3,
    skip_missing_attachments: bool = False,
) -> None:
    """
    Create and optionally display/send an Outlook email.

    Parameters
    ----------
    to : str
        Primary recipients (semicolon-separated).
    subject : str
        Email subject.
    html_body : str
        HTML body. If `inline_image` is provided, use <img src="cid:{{INLINE_CID}}"> in your HTML
        (or set `inline_placeholder` to match your marker).
    attachments : Iterable[str | Path] | None
        Collection of file paths to attach.
    inline_image : str | Path | None
        Single inline image file. Will use `inline_cid` in HTML via placeholder replacement.
    inline_cid : str
        Content-ID for `inline_image` (default: "banner").
    inline_placeholder : str
        Placeholder token inside `html_body` to be replaced with `inline_cid`.
    inline_images : Mapping[str, str | Path] | None
        Multiple inline images as a dict {cid: file_path}. Expect your HTML to reference each as <img src="cid:cid">.
        You can use this together with (or instead of) `inline_image`.
    cc, bcc, reply_to : str | None
        Optional recipients.
    importance : int | str | None
        "Low"/0, "Normal"/1, "High"/2. Defaults to Normal.
    display_before_send : bool
        If True, displays the email (helps ensure WordEditor/inspector readiness).
    send_immediately : bool
        If True, calls Send(); otherwise leaves the window open for manual inspection.
    paste_delay : float
        Time to wait (seconds) after Display() to ensure Outlook is ready for property operations.
    skip_missing_attachments : bool
        If True, logs and skips missing attachment files; otherwise raises FileNotFoundError.

    Notes
    -----
    - We explicitly set `BodyFormat = 2` (HTML) to avoid COM constants dependency issues.
    - Inline images require the `PR_ATTACH_CONTENT_ID` property to match `cid:` used in HTML.
    """
    outlook = win32.gencache.EnsureDispatch("Outlook.Application")
    # ns = outlook.GetNamespace("MAPI")  # available if profile auth needed

    msg = outlook.CreateItem(0)  # 0 = MailItem
    msg.To = to
    if cc:
        msg.CC = cc
    if bcc:
        msg.BCC = bcc
    msg.Subject = subject

    # Backward-compatible single-inline-image support via placeholder
    if inline_image:
        html_body = html_body.replace(inline_placeholder, inline_cid)

    msg.HTMLBody = html_body

    if display_before_send:
        msg.Display()
        d = max(0.0, float(paste_delay))
        if d:
            logging.info("Waiting %.3fs to allow Outlook to initialize...", d)
            time.sleep(d)

    # Add regular attachments (validate existence)
    for path in _as_path_list(attachments):
        if not path.exists():
            if skip_missing_attachments:
                logging.warning("Attachment not found, skipping: %s", path)
                continue
            raise FileNotFoundError(f"Attachment not found: {path}")
        logging.info("Adding attachment: %s", path)
        msg.Attachments.Add(str(path))

    # Add single inline image (if provided)
    if inline_image:
        path = Path(inline_image)
        if not path.exists():
            raise FileNotFoundError(f"Inline image not found: {path}")
        logging.info("Adding inline image (cid=%s): %s", inline_cid, path)
        attach = msg.Attachments.Add(str(path))
        # Set the content-id to match your HTML src="cid:inline_cid"
        attach.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, inline_cid)

    # Add multiple inline images (mapping of cid -> path)
    if inline_images:
        for cid, file_path in inline_images.items():
            p = Path(file_path)
            if not p.exists():
                raise FileNotFoundError(f"Inline image not found for cid '{cid}': {p}")
            logging.info("Adding inline image (cid=%s): %s", cid, p)
            a = msg.Attachments.Add(str(p))
            a.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, cid)

    # Set importance (0=Low, 1=Normal, 2=High)
    msg.Importance = _map_importance(importance)

    # Explicit HTML format to avoid constants import issues
    msg.BodyFormat = 2  # HTML

    # Add Reply-To (optional)
    if reply_to:
        try:
            msg.ReplyRecipients.Add(reply_to)
        except Exception as e:
            # Some profiles may not allow programmatic reply-to modifications
            logging.warning("Could not set Reply-To (%s): %s", reply_to, e)

    if send_immediately:
        logging.info("Sending email...")
        msg.Send()
        logging.info("Email sent.")
    else:
        logging.info("Email displayed but not sent (send_immediately=False).")


if __name__ == "__main__":
    body = """
    <html><body style="font-family:Segoe UI,Arial">
      <h2>Weekly Update</h2>
      <p>Hi team,</p>
      <p>Please find the latest report attached.</p>
      <img src="cid:{{INLINE_CID}}" alt="Banner" />
      <p>Regards,<br/>Automation Bot</p>
    </body></html>
    """

    send_html_email(
        to="your_email@outlook.com",
        subject="Weekly Report",
        html_body=body,
        attachments=[
            r"D:\LifeLongLearning\windows-automation-python\data\reports\weekly.xlsx",
            r"D:\LifeLongLearning\windows-automation-python\data\reports\summary.pdf",
        ],
        inline_image=r"D:\LifeLongLearning\windows-automation-python\data\assets\banner.png",
        inline_cid="banner",
        inline_placeholder="{{INLINE_CID}}",
        # You can also embed multiple inline images:
        # inline_images={"logo": r"D:\path\logo.png", "chart": r"D:\path\chart.png"},
        cc=None,
        bcc=None,
        reply_to=None,
        importance="High",
        display_before_send=True,
        send_immediately=True,
        paste_delay=0.5,
        skip_missing_attachments=False,
    )
