"""
File utility functions for Word Document Server.
"""

import asyncio
import os
import sys

# Per-file locks to prevent concurrent read-modify-write on the same document
_file_locks: dict[str, asyncio.Lock] = {}


def get_file_lock(filepath: str) -> asyncio.Lock:
    """Get or create an asyncio.Lock for a given file path.

    Serializes concurrent async operations on the same document.
    Different files can proceed in parallel.
    """
    path = os.path.realpath(filepath)
    if sys.platform == "win32":
        path = path.casefold()
    return _file_locks.setdefault(path, asyncio.Lock())


def check_file_writeable(filepath: str) -> tuple[bool, str]:
    """
    Check if a file can be written to.

    Args:
        filepath: Path to the file

    Returns:
        Tuple of (is_writeable, error_message)
    """
    # If file doesn't exist, check if directory is writeable
    if not os.path.exists(filepath):
        directory = os.path.dirname(filepath)
        # If no directory is specified (empty string), use current directory
        if directory == "":
            directory = "."
        if not os.path.exists(directory):
            return False, f"Directory {directory} does not exist"
        if not os.access(directory, os.W_OK):
            return False, f"Directory {directory} is not writeable"
        return True, ""

    # If file exists, check if it's writeable
    if not os.access(filepath, os.W_OK):
        return False, f"File {filepath} is not writeable (permission denied)"

    # Try to open the file for writing to see if it's locked
    try:
        with open(filepath, "a"):
            pass
        return True, ""
    except OSError as e:
        return False, f"File {filepath} is not writeable: {str(e)}"
    except Exception as e:
        return False, f"Unknown error checking file permissions: {str(e)}"


def ensure_docx_extension(filename: str) -> str:
    """
    Ensure filename has .docx extension.

    Args:
        filename: The filename to check

    Returns:
        Filename with .docx extension
    """
    if not filename.casefold().endswith(".docx"):
        return filename + ".docx"
    return filename
