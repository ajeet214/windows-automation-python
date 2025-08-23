import msoffcrypto
import tempfile
import win32com.client as win32
from pathlib import Path
import logging

# -----------------------
# Configuration & Logging
# -----------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)


def open_protected_ppt(file_path: str, password: str):
    """
    Decrypts a password-protected PowerPoint file and opens it in PowerPoint.

    Args:
        file_path (str): Path to the password-protected PPTX file.
        password (str): Password for the PPTX file.

    Returns:
        tuple: (PowerPoint Application COM object, Presentation COM object)
    """
    file_path = Path(file_path)
    try:
        # Decrypt to a temporary file
        with open(file_path, "rb") as f:
            office = msoffcrypto.OfficeFile(f)
            office.load_key(password=password)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as out:
                office.decrypt(out)
            decrypted_path = out.name

        # Open decrypted copy
        ppt = win32.Dispatch("PowerPoint.Application")
        ppt.Visible = True
        pres = ppt.Presentations.Open(decrypted_path, ReadOnly=True, WithWindow=True)

        return ppt, pres

    except msoffcrypto.exceptions.InvalidKeyError as e:
        logging.error(e)


# ---------------- Example usage ---------------- #
if __name__ == "__main__":
    ppt_file = r"<file path>"
    pwd = "<password>"
    open_protected_ppt(ppt_file, pwd)
