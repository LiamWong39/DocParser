"""Normalize messy Word documents so they can be chunked reliably."""

from .pipeline import normalize, normalize_file

__all__ = ["normalize", "normalize_file"]
