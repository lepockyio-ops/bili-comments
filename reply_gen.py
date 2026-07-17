"""
BiliComments v2 · 回复推荐模块
==============================
读取 reply_templates.yaml，对每条评论做意图分类 + 推荐 2-3 条日语回复模板。

依赖：pyyaml（stdlib random/re/pathlib 足矣）
"""
from __future__ import annotations

import random
import re
from pathlib import Path
from typing import Optional

import yaml


SCRIPT_DIR = Path(__file__).parent.resolve()
TEMPLATES_PATH = SCRIPT_DIR / "reply_templates.yaml"


class ReplyGenerator:
    """基于关键词的评论意图分类 + 日语回复模板匹配。"""

    def __init__(self):
        self.config = self._load()
        self.intents = self.config.get("intents", {}) or {}
        self.rules = self.config.get("rules", {}) or {}
        self.long_min = self.rules.get("long_comment_min_chars", 50)
        self.skip_up = self.rules.get("skip_up_comments", True)
        self.skip_sub = self.rules.get("skip_sub_replies", True)

        # 按意图顺序检测（前面的先匹配）
        self.intent_order = list(self.intents.keys())

    def _load(self) -> dict:
        if not TEMPLATES_PATH.exists():
            return {}
        with TEMPLATES_PATH.open(encoding="utf-8") as f:
            return yaml.safe_load(f) or {}

    def classify(self, message: str, is_up: bool = False, is_sub: bool = False) -> str:
        """
        判断评论的意图。返回 intent 名字。
        规则：
        1. 若 is_up=True → "UP主本人"（跳过）
        2. 若 is_sub=True 且 skip_sub_replies → 也跳过
        3. 按 intents 顺序找第一个关键词匹配的意图
        4. 长评论（>50 字）无匹配时提升到"情感深度"
        5. 否则默认"其他"
        """
        if is_up and self.skip_up:
            return "UP主本人"
        if is_sub and self.skip_sub:
            return "楼中楼跳过"

        # 关键词匹配
        for intent_name in self.intent_order:
            keywords = self.intents[intent_name].get("keywords", []) or []
            for kw in keywords:
                if kw.lower() in message.lower():
                    return intent_name

        # 长评论 → 情感深度
        if len(message) >= self.long_min:
            return "情感深度"

        return "其他"

    def suggest_replies(self, intent: str, k: int = 3) -> list[str]:
        """从对应 intent 的模板池里随机挑 k 条。"""
        if intent in ("UP主本人", "楼中楼跳过"):
            return []
        pool = self.intents.get(intent, {}).get("templates", []) or []
        if not pool:
            pool = self.intents.get("其他", {}).get("templates", []) or []
        if len(pool) <= k:
            return list(pool)
        return random.sample(pool, k)

    def generate_for(
        self, message: str, is_up: bool = False, is_sub: bool = False, k: int = 3
    ) -> tuple[str, list[str]]:
        """一步到位：返回 (intent, [reply1, reply2, reply3])"""
        intent = self.classify(message, is_up=is_up, is_sub=is_sub)
        replies = self.suggest_replies(intent, k=k)
        return intent, replies


# 便捷函数
def generate_replies_for_comments(comments: list[dict], k: int = 3) -> list[tuple[str, list[str]]]:
    """
    输入：collect_comments.extract_comment 输出的 comment 列表
    输出：[(intent, [replies]), ...] 与输入等长
    """
    gen = ReplyGenerator()
    results = []
    for c in comments:
        is_up = c.get("is_up") == "是"
        is_sub = c.get("is_sub", False)
        msg = c.get("message", "")
        intent, replies = gen.generate_for(msg, is_up=is_up, is_sub=is_sub, k=k)
        results.append((intent, replies))
    return results


if __name__ == "__main__":
    # 自测
    g = ReplyGenerator()
    tests = [
        ("这次的Teto也太可爱了吧！？", False, False),
        ("求授权想翻唱到日站", False, False),
        ("求MP3！求kara版！", False, False),
        ("YYDS 单曲循环一整天了", False, False),
        ("好像有版权问题？这段旋律和XX很像", False, False),
        ("这段歌词太戳我了，让我想起去年那段感情，眼泪根本止不住，反复听了三十遍还是要哭", False, False),
        ("回复UP：谢谢你的曲子", True, False),
    ]
    for msg, is_up, is_sub in tests:
        intent, replies = g.generate_for(msg, is_up=is_up, is_sub=is_sub)
        print(f"评论: {msg[:40]}")
        print(f"意图: {intent}")
        for i, r in enumerate(replies, 1):
            print(f"  推荐{i}: {r}")
        print()
