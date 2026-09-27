"""Normalization boundary for activity, category, role/level and score fields."""

import re
import unicodedata


_WHITESPACE_RE = re.compile(r"\s+")
_LEADING_INDEX_RE = re.compile(r"^\s*\d+\s*[-—–、.．]\s*")


def normalize_text(value: object) -> str:
    """Apply conservative Unicode and whitespace normalization to a cell."""
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value)).replace("\u3000", " ")
    return _WHITESPACE_RE.sub(" ", text).strip()


def normalize_activity_name(value: object) -> str:
    """Normalize an activity name and remove a simple leading item number."""
    return _LEADING_INDEX_RE.sub("", normalize_text(value), count=1)


def normalize_reference_score(value: object) -> float | None:
    """Convert an Excel score cell to a numeric reference score."""
    text = normalize_text(value)
    if not text:
        return None
    try:
        return float(text)
    except ValueError as error:
        raise ValueError(f"无法识别参考分数: {text!r}") from error
