"""Phase 1 (optional): Use an LLM to refine chapter text into slide bullets.

Supports both Claude API (Anthropic) and OpenAI-compatible endpoints.

For long chapters that exceed the context window, this module applies a
recursive map-reduce strategy:
  1. Split the text into paragraph-aligned chunks.
  2. Summarize each chunk independently.
  3. Merge / deduplicate the resulting bullet sets into one final list.
"""

import json
import os
import time
from typing import Any


# ---------------------------------------------------------------------------
# Prompts
# ---------------------------------------------------------------------------

_SYSTEM_PROMPT = """\
你是一个答辩PPT制作助手。用户会提供论文某一章节的正文内容。
请将其提炼为3-6条bullet points（每条不超过一行），适合放入答辩PPT。

要求：
- 每条bullet必须简洁，控制在25字以内
- 保留核心数据指标（百分比、系数、金额等）
- 去掉推导过程和方法论细节
- 使用中文输出

严格按以下JSON格式输出，不要输出任何其他内容：
[{"bullet": "...", "ref_page": N}, ...]"""

_USER_TEMPLATE = """\
论文章节：{chapter_title}

正文内容：
{chapter_text}"""

_MERGE_SYSTEM = """\
你是一个答辩PPT制作助手。以下是同一章节拆分为多段后分别提炼的bullet points。
请将这些bullet points合并为3-6条（每条不超过一行），去掉重复和冗余的内容。

严格按以下JSON格式输出，不要输出任何其他内容：
[{"bullet": "...", "ref_page": N}, ...]"""

_MERGE_USER = """\
章节主题：{chapter_title}

各段提炼结果：
{bullet_sets}"""


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _clean_json(raw: str) -> str:
    """Strip markdown code fences from an LLM response."""
    raw = raw.strip()
    if raw.startswith("```"):
        idx = raw.index("\n")
        raw = raw[idx + 1:]
        if raw.endswith("```"):
            raw = raw[: raw.rindex("```")].strip()
    return raw


def _parse_bullets(raw: str) -> list[dict[str, Any]]:
    """Parse the LLM JSON response, returning a bullet list or empty list."""
    try:
        return json.loads(_clean_json(raw))
    except json.JSONDecodeError:
        return []


def _split_text(text: str, max_chars: int) -> list[str]:
    """Split text into chunks at paragraph boundaries, each ≤ max_chars."""
    paragraphs = text.split("\n\n")
    chunks: list[str] = []
    current: list[str] = []
    current_len = 0

    for para in paragraphs:
        para = para.strip()
        if not para:
            continue
        if current_len + len(para) > max_chars and current:
            chunks.append("\n\n".join(current))
            current = [para]
            current_len = len(para)
        else:
            current.append(para)
            current_len += len(para)

    if current:
        chunks.append("\n\n".join(current))

    return chunks


# ---------------------------------------------------------------------------
# LLM call wrappers (with retry)
# ---------------------------------------------------------------------------

def _call_claude(
    system: str,
    user: str,
    *,
    api_key: str | None = None,
    model: str = "claude-sonnet-4-6",
    max_tokens: int = 1024,
) -> str:
    from anthropic import Anthropic

    client = Anthropic(api_key=api_key or os.environ.get("ANTHROPIC_API_KEY"))

    response = client.messages.create(
        model=model,
        max_tokens=max_tokens,
        system=system,
        messages=[{"role": "user", "content": user}],
    )
    return response.content[0].text


def _call_openai(
    system: str,
    user: str,
    *,
    api_key: str | None = None,
    base_url: str | None = None,
    model: str = "gpt-4o",
    max_tokens: int = 1024,
) -> str:
    from openai import OpenAI

    client = OpenAI(
        api_key=api_key or os.environ.get("OPENAI_API_KEY"),
        base_url=base_url or os.environ.get("OPENAI_BASE_URL"),
    )
    response = client.chat.completions.create(
        model=model,
        max_tokens=max_tokens,
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
    )
    return response.choices[0].message.content


def _call_with_retry(
    caller,
    system: str,
    user: str,
    max_retries: int = 3,
    **kwargs: Any,
) -> str:
    """Call an LLM with exponential-backoff retry."""
    last_error: Exception | None = None
    for attempt in range(max_retries):
        try:
            return caller(system, user, **kwargs)
        except Exception as e:
            last_error = e
            if attempt < max_retries - 1:
                time.sleep(2 ** attempt)
    raise last_error  # type: ignore


# ---------------------------------------------------------------------------
# Map-Reduce summarizer
# ---------------------------------------------------------------------------

def _summarize_raw(
    chapter_title: str,
    chapter_text: str,
    *,
    provider: str = "claude",
    **kwargs: Any,
) -> list[dict[str, Any]]:
    """Single-shot: send text to LLM, return parsed bullets."""
    caller = _call_claude if provider == "claude" else _call_openai
    raw = _call_with_retry(
        caller,
        _SYSTEM_PROMPT,
        _USER_TEMPLATE.format(
            chapter_title=chapter_title,
            chapter_text=chapter_text,
        ),
        **kwargs,
    )
    return _parse_bullets(raw)


def _merge_bullets(
    chapter_title: str,
    bullet_sets: list[list[dict[str, Any]]],
    *,
    provider: str = "claude",
    **kwargs: Any,
) -> list[dict[str, Any]]:
    """Merge multiple bullet lists into one, removing duplicates."""
    # Flatten for the prompt
    parts: list[str] = []
    for i, bs in enumerate(bullet_sets, 1):
        parts.append(f"段{i}: {json.dumps(bs, ensure_ascii=False)}")
    merged_text = "\n".join(parts)

    caller = _call_claude if provider == "claude" else _call_openai
    raw = _call_with_retry(
        caller,
        _MERGE_SYSTEM,
        _MERGE_USER.format(
            chapter_title=chapter_title,
            bullet_sets=merged_text,
        ),
        **kwargs,
    )
    return _parse_bullets(raw)


def summarize_chapter(
    chapter_title: str,
    chapter_text: str,
    *,
    provider: str = "claude",
    max_chars: int = 10000,  # leave headroom for prompt boilerplate
    **kwargs: Any,
) -> list[dict[str, Any]]:
    """Refine chapter text into bullet points via LLM.

    If the text exceeds *max_chars*, it is split into paragraph-aligned
    chunks.  Each chunk is summarized independently, then the results are
    merged with a second LLM pass to deduplicate.

    Args:
        chapter_title: Section heading (used in the prompt context).
        chapter_text: The raw text content of this chapter/section.
        provider: ``"claude"`` or ``"openai"``.
        max_chars: Maximum characters to send in a single LLM request.
        **kwargs: Forwarded (model, api_key, max_tokens, etc.).

    Returns:
        A list of ``{"bullet": str, "ref_page": int}`` dicts.
    """
    if not chapter_text.strip():
        return []

    if len(chapter_text) <= max_chars:
        return _summarize_raw(chapter_title, chapter_text, provider=provider, **kwargs)

    # ---- Map: summarise each chunk independently ----
    chunks = _split_text(chapter_text, max_chars)
    all_bullets: list[list[dict[str, Any]]] = []
    for i, chunk in enumerate(chunks):
        chunk_title = f"{chapter_title}（第{i + 1}/{len(chunks)}部分）"
        bullets = _summarize_raw(chunk_title, chunk, provider=provider, **kwargs)
        if bullets:
            all_bullets.append(bullets)

    if not all_bullets:
        return []
    if len(all_bullets) == 1:
        return all_bullets[0]

    # ---- Reduce: merge bullet sets ----
    return _merge_bullets(chapter_title, all_bullets, provider=provider, **kwargs)
