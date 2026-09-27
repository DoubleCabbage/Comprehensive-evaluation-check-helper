import pytest

from activity_matcher.normalize import (
    normalize_activity_name,
    normalize_reference_score,
    normalize_text,
)


def test_normalize_text_handles_full_width_and_whitespace() -> None:
    assert normalize_text("  Ａ　 B  ") == "A B"


def test_normalize_activity_name_removes_simple_leading_index() -> None:
    assert normalize_activity_name("1-2025大学生创新创业训练计划结题") == "2025大学生创新创业训练计划结题"


def test_normalize_reference_score() -> None:
    assert normalize_reference_score(" 10 ") == 10.0
    assert normalize_reference_score(2.5) == 2.5
    assert normalize_reference_score("") is None


def test_invalid_reference_score_is_rejected() -> None:
    with pytest.raises(ValueError):
        normalize_reference_score("待定")
