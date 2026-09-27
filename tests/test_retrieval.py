from activity_matcher.catalog import ActivityNode
from activity_matcher.retrieval import MatchQuery, retrieve


def _activity(name: str) -> ActivityNode:
    return ActivityNode(
        activity_id=name,
        activity_name=name,
        major_category="学术创新",
        sub_category="学术科技竞赛",
        variant_label="",
        needs_review=False,
        roles=(),
    )


def test_year_and_edition_changes_do_not_block_match() -> None:
    activities = [
        _activity("第42届全国部分地区大学生物理竞赛"),
        _activity("全国青少年无线电通信锦标赛"),
    ]
    result = retrieve(
        MatchQuery(activity_name="第43届全国部分地区大学生物理比赛"),
        activities,
        limit=1,
    )
    assert result[0].activity.activity_name == "第42届全国部分地区大学生物理竞赛"
