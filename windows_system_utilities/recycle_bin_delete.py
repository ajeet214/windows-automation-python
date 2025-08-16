"""
Utility script to send files or folders to the Windows Recycle Bin
using the pywin32 shell API.

This script demonstrates how to perform a soft delete (Recycle Bin)
instead of permanently deleting files. It uses the `SHFileOperation`
function from the Windows Shell.
"""

from __future__ import annotations
import logging
from pathlib import Path
from win32com.shell import shell, shellcon  # type: ignore

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")


def send_to_recycle_bin(path: Path) -> None:
    """
    Move a file or folder to the Windows Recycle Bin.

    Parameters
    ----------
    path : Path
        The file or folder path to delete.

    Raises
    ------
    RuntimeError
        If the shell operation fails (non-zero return code).
    """
    # FO_DELETE with FOF_ALLOWUNDO => move to Recycle Bin (not permanent delete)
    res = shell.SHFileOperation((
        0,                            # hwnd (owner window handle)
        shellcon.FO_DELETE,           # wFunc (delete operation)
        str(path),                    # pFrom (source path)
        None,                         # pTo (unused here)
        shellcon.FOF_NOCONFIRMATION   # no confirmation dialogs
        | shellcon.FOF_SILENT         # no UI
        | shellcon.FOF_ALLOWUNDO,     # allow undo (Recycle Bin)
        None, None
    ))
    if res[0] != 0:
        raise RuntimeError(f"Shell delete failed, code={res[0]}")
    logging.info("Moved to Recycle Bin -> %s", path)


if __name__ == "__main__":
    send_to_recycle_bin(Path(r"D:\Book1.xlsx"))
