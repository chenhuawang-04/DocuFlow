"""Shared dependency availability checks with caching."""
import importlib
import subprocess
from typing import List, Optional

_cache: dict = {}


def check_import(*module_names: str) -> bool:
    """Check whether Python package(s) are importable. Results are cached.

    Args:
        module_names: One or more module names. All must succeed.
                      First name is used as cache key.

    Examples:
        check_import("openpyxl")
        check_import("pptx")
        check_import("reportlab")
    """
    cache_key = module_names[0]
    if cache_key in _cache:
        return _cache[cache_key]
    try:
        for name in module_names:
            importlib.import_module(name)
        _cache[cache_key] = True
    except ImportError:
        _cache[cache_key] = False
    return _cache[cache_key]


def check_command(command: str, args: Optional[List[str]] = None,
                  timeout: int = 10) -> bool:
    """Check whether a CLI tool is available. Results are cached.

    Args:
        command: Executable name (e.g. "tesseract", "pandoc")
        args: Arguments to pass (default: ["--version"])
        timeout: Timeout in seconds
    """
    if command in _cache:
        return _cache[command]
    if args is None:
        args = ["--version"]
    try:
        result = subprocess.run(
            [command] + args,
            capture_output=True, text=True, timeout=timeout
        )
        _cache[command] = result.returncode == 0
    except (subprocess.SubprocessError, FileNotFoundError, OSError):
        _cache[command] = False
    return _cache[command]


def clear_cache():
    """Clear all cached results (useful for testing)."""
    _cache.clear()
