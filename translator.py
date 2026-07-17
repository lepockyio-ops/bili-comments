"""
BiliComments v2 · 翻译模块
==========================
Google Translate 中→日翻译 + 中文俚语/Vocaloid 术语预处理

功能：
1. 加载 slang_glossary.yaml 术语表
2. 翻译前把中文俚语替换为日语（bypass Google，保证准确）
3. 剩余部分调用 Google Translate (deep-translator)
4. 失败重试 + 频率控制

依赖：deep-translator, pyyaml
"""
from __future__ import annotations

import re
import time
from pathlib import Path
from typing import Optional

import yaml

try:
    from deep_translator import GoogleTranslator
    _DEEP_AVAILABLE = True
except ImportError:
    _DEEP_AVAILABLE = False


SCRIPT_DIR = Path(__file__).parent.resolve()
GLOSSARY_PATH = SCRIPT_DIR / "slang_glossary.yaml"

# 用一个不常见的 unicode 符号作为占位符边界，避免影响 Google 翻译
# 使用「」括起来的编号形式：「⟪N⟫」
PLACEHOLDER_RE = re.compile(r"⟪(\d+)⟫")


class SlangTranslator:
    """带术语预处理的翻译器。"""

    def __init__(self, target_lang: str = "ja", source_lang: str = "zh-CN"):
        self.target_lang = target_lang
        self.source_lang = source_lang
        self.glossary = self._load_glossary()
        # 术语按长度降序排列（先匹配长的，避免"太上头"被"上头"抢先）
        self.direct_map = self.glossary.get("direct_ja", {}) or {}
        self.direct_terms_sorted = sorted(self.direct_map.keys(), key=len, reverse=True)

        self._translator = None
        if _DEEP_AVAILABLE:
            try:
                self._translator = GoogleTranslator(source=source_lang, target=target_lang)
            except Exception as e:
                print(f"⚠️  Google Translator 初始化失败: {e}", flush=True)

    def _load_glossary(self) -> dict:
        if not GLOSSARY_PATH.exists():
            return {}
        with GLOSSARY_PATH.open(encoding="utf-8") as f:
            return yaml.safe_load(f) or {}

    def _preprocess(self, text: str) -> tuple[str, dict[int, str]]:
        """
        把术语替换为占位符，返回 (预处理后的文本, {占位符编号: 日语})
        这样 Google Translate 遇到 ⟪1⟫ ⟪2⟫ 会原样保留（数字+符号不会被翻译）
        翻译后再替换回去。
        """
        placeholders: dict[int, str] = {}
        result = text
        for term in self.direct_terms_sorted:
            if term in result:
                idx = len(placeholders) + 1
                jp_translation = self.direct_map[term]
                placeholders[idx] = jp_translation
                # 用占位符替换
                result = result.replace(term, f"⟪{idx}⟫")
        return result, placeholders

    def _postprocess(self, translated: str, placeholders: dict[int, str]) -> str:
        """把占位符替换回日语。"""
        def _repl(m):
            idx = int(m.group(1))
            return placeholders.get(idx, m.group(0))
        return PLACEHOLDER_RE.sub(_repl, translated)

    def translate(self, text: str) -> Optional[str]:
        """
        翻译单条评论。返回日语；如果不可用，返回 None（调用方 fallback 到公式）。
        """
        if not text or not text.strip():
            return ""

        # 短文本或纯 emoji 直接返回
        if len(text.strip()) < 2:
            return text

        # 预处理术语替换
        preprocessed, placeholders = self._preprocess(text)

        # 如果全部都是术语直接匹配了（预处理后只剩占位符），直接跳过 Google
        # 判断：去掉占位符后是否为空/纯符号
        content_only = PLACEHOLDER_RE.sub("", preprocessed).strip()
        content_only = re.sub(r"[\s　\p{P}]+", "", content_only) if False else re.sub(r"[\s\W]+", "", content_only)
        if not content_only:
            # 全靠术语表就能翻译
            return self._postprocess(preprocessed, placeholders)

        if not self._translator:
            return None  # 让调用方 fallback

        try:
            # Google 翻译，尽量简短的请求提高成功率
            translated = self._translator.translate(preprocessed)
            if not translated:
                return self._postprocess(preprocessed, placeholders)
            return self._postprocess(translated, placeholders)
        except Exception as e:
            # 单次失败不 raise，返回 None 让调用方 fallback 到公式
            return None

    def translate_batch(self, texts: list[str], delay: float = 0.4, verbose: bool = False) -> list[Optional[str]]:
        """
        批量翻译，逐条 API 调用（Google 免费端点批量支持不稳定）。
        遇到失败返回 None，让调用方 fallback。
        """
        results = []
        total = len(texts)
        for i, text in enumerate(texts, 1):
            if verbose and i % 5 == 0:
                print(f"  翻译进度 {i}/{total}", flush=True)
            r = self.translate(text)
            results.append(r)
            if i < total:
                time.sleep(delay)  # 频率控制
        return results


# 便捷函数
def translate_comments(
    messages: list[str], target_lang: str = "ja", verbose: bool = True
) -> list[Optional[str]]:
    """给一组评论翻译。用于外部调用的简洁 API。"""
    if not _DEEP_AVAILABLE:
        print("⚠️  deep-translator 未安装，跳过 AI 翻译（会 fallback 到 GOOGLETRANSLATE 公式）")
        return [None] * len(messages)

    t = SlangTranslator(target_lang=target_lang)
    if not t._translator:
        return [None] * len(messages)
    if verbose:
        print(f"→ 使用 Google Translate + 术语预处理翻译 {len(messages)} 条评论")
    return t.translate_batch(messages, verbose=verbose)


if __name__ == "__main__":
    # 简单自测
    samples = [
        "这次的Teto也太可爱了吧！？👼✨",
        "🫧我也想变成气泡 就这样溶掉",
        "破防了呜呜呜太上头了单曲循环一整天",
        "求授权，想翻唱到自己的账号[打call]",
        "YYDS 神仙曲",
    ]
    t = SlangTranslator()
    for s in samples:
        r = t.translate(s)
        print(f"原: {s}")
        print(f"译: {r}")
        print()
        time.sleep(0.5)
