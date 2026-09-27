"""Dependency-free lexical baseline for activity retrieval."""

from dataclasses import dataclass
from difflib import SequenceMatcher
import re
from typing import Iterable

from .catalog import ActivityNode
from .normalize import normalize_activity_name


_YEAR_RE = re.compile(r"(?<!\d)(?:19|20)\d{2}年?")
_EDITION_RE = re.compile(r"第\s*\d+\s*届")
_NON_WORD_RE = re.compile(r"[^0-9A-Za-z\u4e00-\u9fff]+")


@dataclass(frozen=True, slots=True)
class MatchQuery:
    activity_name: str
    major_category: str = ""
    sub_category: str = ""


@dataclass(frozen=True, slots=True)
class MatchResult:
    activity: ActivityNode
    score: float
    activity_score: float
    category_score: float


def retrieve(
    query: MatchQuery,
    activities: Iterable[ActivityNode],
    *,
    limit: int = 5,
) -> list[MatchResult]:
    """Rank activity nodes; role selection intentionally happens after retrieval."""
    if limit < 1:
        raise ValueError("limit must be at least 1")
    query_key = _search_key(query.activity_name)
    if not query_key:
        return []

    results: list[MatchResult] = []
    for activity in activities:
        candidate_key = _search_key(activity.activity_name)
        activity_score = _text_similarity(query_key, candidate_key)
        category_parts: list[float] = []
        if query.major_category:
            category_parts.append(float(query.major_category == activity.major_category))
        if query.sub_category:
            category_parts.append(float(query.sub_category == activity.sub_category))
        category_score = (
            sum(category_parts) / len(category_parts) if category_parts else 0.0
        )
        total_score = (
            0.85 * activity_score + 0.15 * category_score
            if category_parts
            else activity_score
        )
        results.append(
            MatchResult(
                activity=activity,
                score=round(total_score, 6),
                activity_score=round(activity_score, 6),
                category_score=round(category_score, 6),
            )
        )

    return sorted(
        results,
        key=lambda result: (
            -result.score,
            result.activity.needs_review,
            result.activity.activity_name,
            result.activity.variant_label,
        ),
    )[:limit]


def _search_key(value: str) -> str:
    text = normalize_activity_name(value)
    text = _YEAR_RE.sub("", text)
    text = _EDITION_RE.sub("", text)
    return _NON_WORD_RE.sub("", text).lower()


def _text_similarity(left: str, right: str) -> float:
    if left == right:
        return 1.0
    sequence_score = SequenceMatcher(None, left, right).ratio()
    left_bigrams = _character_ngrams(left, 2)
    right_bigrams = _character_ngrams(right, 2)
    union = left_bigrams | right_bigrams
    jaccard_score = len(left_bigrams & right_bigrams) / len(union) if union else 0.0
    containment_bonus = 1.0 if left in right or right in left else 0.0
    return 0.5 * sequence_score + 0.4 * jaccard_score + 0.1 * containment_bonus


def _character_ngrams(text: str, size: int) -> set[str]:
    if len(text) <= size:
        return {text} if text else set()
    return {text[index : index + size] for index in range(len(text) - size + 1)}
