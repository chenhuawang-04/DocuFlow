"""Path validation utilities to prevent directory traversal."""
import os
from pathlib import Path


class PathValidationError(Exception):
    """Raised when a path fails validation."""
    pass


def validate_path(path_str: str, must_exist: bool = False) -> str:
    """Normalize and validate a user-supplied file path.

    - Resolves to absolute path
    - Blocks null bytes

    Args:
        path_str: Raw path string from user
        must_exist: If True, raise if file does not exist

    Returns:
        Resolved absolute path as string

    Raises:
        PathValidationError: On invalid path
        FileNotFoundError: If must_exist=True and file missing
    """
    if not path_str or not isinstance(path_str, str):
        raise PathValidationError("Path must be a non-empty string")

    if "\x00" in path_str:
        raise PathValidationError("Path contains null bytes")

    resolved = str(Path(path_str).resolve())

    if must_exist and not os.path.exists(resolved):
        raise FileNotFoundError(f"File not found: {resolved}")

    return resolved
