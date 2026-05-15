"""Phase 1 (optional): Use an LLM to refine chapter text into slide bullets.

Supports both Claude API (Anthropic) and OpenAI-compatible endpoints.
"""

import json
import os
from typing import Any


_SUMMARIZE_SYSTEM = """\
你是一个答辩PPT制作助手。用户会提供论文某一章节的正文内容。
请将其提炼为3-6条bullet points（每条不超过一行），适合放入答辩PPT。

要求：
- 每条bullet必须简洁，控制在25字以内
- 保留核心数据指标（百分比、系数、金额等）
- 去掉推导过程和方法论细节
- 使用中文输出

严格按以下JSON格式输出，不要输出任何其他内容：
[{"bullet": "...", "ref_page": N}, ...]"""


_SUMMARIZE_USER_TEMPLATE = """\
论文章节：{chapter_title}

正文内容：
{chapter_text}"""


# ---------------------------------------------------------------------------
# Claude (Anthropic)
# ---------------------------------------------------------------------------

def summarize_with_claude(
    chapter_title: str,
    chapter_text: str,
    *,
    api_key: str | None = None,
    model: str = "claude-sonnet-4-6",
    max_tokens: int = 1024,
) -> list[dict[str, Any]]:
    """Summarize a chapter into bullets using the Anthropic Claude API."""
    from anthropic import Anthropic

    client = Anthropic(api_key=api_key or os.environ.get("ANTHROPIC_API_KEY"))

    response = client.messages.create(
        model=model,
        max_tokens=max_tokens,
        system=_SUMMARIZE_SYSTEM,
        messages=[
            {
                "role": "user",
                "content": _SUMMARIZE_USER_TEMPLATE.format(
                    chapter_title=chapter_title,
                    chapter_text=chapter_text[:15000],  # safety cap
                ),
            }
        ],
    )

    raw = response.content[0].text
    # Strip markdown code fences if present
    raw = raw.strip()
    if raw.startswith("```"):
        raw = raw[raw.index("\n") + 1 :]
        if raw.endswith("```"):
            raw = raw[: raw.rindex("```")].strip()

    return json.loads(raw)


# ---------------------------------------------------------------------------
# OpenAI-compatible
# ---------------------------------------------------------------------------

def summarize_with_openai(
    chapter_title: str,
    chapter_text: str,
    *,
    api_key: str | None = None,
    base_url: str | None = None,
    model: str = "gpt-4o",
    max_tokens: int = 1024,
) -> list[dict[str, Any]]:
    """Summarize a chapter into bullets using an OpenAI-compatible API."""
    from openai import OpenAI

    client = OpenAI(
        api_key=api_key or os.environ.get("OPENAI_API_KEY"),
        base_url=base_url or os.environ.get("OPENAI_BASE_URL"),
    )

    response = client.chat.completions.create(
        model=model,
        max_tokens=max_tokens,
        messages=[
            {"role": "system", "content": _SUMMARIZE_SYSTEM},
            {
                "role": "user",
                "content": _SUMMARIZE_USER_TEMPLATE.format(
                    chapter_title=chapter_title,
                    chapter_text=chapter_text[:15000],
                ),
            },
        ],
    )

    raw = response.choices[0].message.content.strip()
    if raw.startswith("```"):
        raw = raw[raw.index("\n") + 1 :]
        if raw.endswith("```"):
            raw = raw[: raw.rindex("```")].strip()

    return json.loads(raw)


# ---------------------------------------------------------------------------
# Dispatcher
# ---------------------------------------------------------------------------

def summarize_chapter(
    chapter_title: str,
    chapter_text: str,
    *,
    provider: str = "claude",
    **kwargs: Any,
) -> list[dict[str, Any]]:
    """Refine chapter text into bullet points via LLM.

    Args:
        chapter_title: Section heading (used in the prompt context).
        chapter_text: The raw text content of this chapter/section.
        provider: "claude" or "openai".
        **kwargs: Forwarded to the provider-specific function
                  (model, api_key, max_tokens, etc.).

    Returns:
        A list of {"bullet": str, "ref_page": int} dicts.
    """
    if provider == "claude":
        return summarize_with_claude(chapter_title, chapter_text, **kwargs)
    elif provider == "openai":
        return summarize_with_openai(chapter_title, chapter_text, **kwargs)
    else:
        raise ValueError(f"Unknown LLM provider: {provider}")
