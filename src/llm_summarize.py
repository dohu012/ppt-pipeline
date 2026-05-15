"""Phase 1 (optional): Use an LLM to refine chapter text into slide bullets,
and — when full-page screenshots are provided — to decide which figures and
tables are worth including in the PPT and extract table data from the images.

Supports both Claude API (Anthropic) and OpenAI-compatible endpoints.

For long chapters that exceed the context window, a recursive map-reduce
strategy is applied (see ``summarize_chapter``).
"""

import base64
import json
import os
import time
from typing import Any


# ---------------------------------------------------------------------------
# Prompts (text-only)
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
# Prompts (vision — with figure / table screenshots)
# ---------------------------------------------------------------------------

_VISION_SYSTEM = """\
你是一个答辩PPT制作助手。你会收到论文某一章节的正文内容，以及该章节内图表所在页面的整页截图。

你需要完成三项任务：
1. 将正文提炼为3-6条bullet points（每条不超过25字，保留核心数据指标）
2. 逐一审视每张图表截图，判断它是否值得放进答辩PPT
   - 标准：核心结论图、模型框架图、关键数据对比表 → 保留
   - 太细节的推导图、次要附表、装饰性图表 → 跳过
3. 对于保留的表格，直接从截图中读取并输出完整的数据（表头+所有行）

严格按以下JSON格式输出，不要输出任何其他内容：
{
  "bullets": [{"bullet": "...", "ref_page": N}, ...],
  "figures": [
    {"number": "图X-Y", "keep": true, "caption": "原标题", "reason": "一句话理由"},
    {"number": "图X-Y", "keep": false, "reason": "一句话理由"}
  ],
  "tables": [
    {"number": "表X-Y", "keep": true, "caption": "原标题", "header": ["列1", "列2"], "rows": [["值1", "值2"]]},
    {"number": "表X-Y", "keep": false, "reason": "一句话理由"}
  ]
}

注意：
- figures 数组里必须包含输入中的每一张图（keep true 或 false）
- tables 数组里必须包含输入中的每一个表（keep true 或 false）
- 表格数据请逐行逐列仔细抄录，保留所有数字精度"""

_VISION_USER_HEADER = """\
论文章节：{chapter_title}

正文内容：
{chapter_text}

---

以下为该章节内的图表整页截图，请逐一审视："""

_VISION_ITEM = """\
[{idx}] {number}：{caption}"""


# ---------------------------------------------------------------------------
# Prompts (chapter-level multi-slide)
# ---------------------------------------------------------------------------

_CHAPTER_SYSTEM = """\
你是一个答辩PPT制作助手。用户会提供论文某一章的完整内容及该章内图表截图。

请完成两项任务：
1. 为这一章设计2-4页PPT内容，将整章内容提炼为核心要点
   - 不要照搬原章节的小节标题，而是按主题重新组织
   - 每页PPT 3-6条bullet points，每条不超过25字
   - 保留核心数据指标（百分比、系数、金额等）
   - 去掉推导过程和方法论细节
   - 各页之间内容不重复
   - 使用中文输出
2. 逐一审视每张图表截图，判断它是否值得放进答辩PPT
   - 标准：核心结论图、模型框架图、关键数据对比表 → 保留
   - 太细节的推导图、次要附表、装饰性图表 → 跳过

严格按以下JSON格式输出，不要输出任何其他内容：
{
  "slides": [
    {
      "title": "本页标题（简短，4-8字）",
      "bullets": [{"bullet": "...", "ref_page": N}, ...]
    },
    ...
  ],
  "figures": [
    {"number": "图X-Y", "keep": true, "caption": "原标题", "reason": "一句话理由"},
    {"number": "图X-Y", "keep": false, "reason": "一句话理由"}
  ],
  "tables": [
    {"number": "表X-Y", "keep": true, "caption": "原标题", "header": ["列1", "列2"], "rows": [["值1", "值2"]]},
    {"number": "表X-Y", "keep": false, "reason": "一句话理由"}
  ]
}

注意：
- slides 数组长度2-4
- figures 数组里必须包含输入中的每一张图（keep true 或 false）
- tables 数组里必须包含输入中的每一个表（keep true 或 false）
- 表格数据请逐行逐列仔细抄录，保留所有数字精度"""

_CHAPTER_USER_TEMPLATE = """\
论文章节：{chapter_title}

完整正文内容：
{chapter_text}"""


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _clean_json(raw: str) -> str:
    raw = raw.strip()
    if raw.startswith("```"):
        idx = raw.index("\n")
        raw = raw[idx + 1:]
        if raw.endswith("```"):
            raw = raw[: raw.rindex("```")].strip()
    return raw


def _parse_bullets(raw: str) -> list[dict[str, Any]]:
    try:
        return json.loads(_clean_json(raw))
    except json.JSONDecodeError:
        return []


def _parse_vision_result(raw: str) -> dict[str, Any]:
    try:
        return json.loads(_clean_json(raw))
    except json.JSONDecodeError:
        return {"bullets": [], "figures": [], "tables": []}


def _parse_chapter_result(raw: str) -> dict[str, Any]:
    """Parse multi-slide chapter-level LLM output."""
    try:
        return json.loads(_clean_json(raw))
    except json.JSONDecodeError:
        return {"slides": [], "figures": [], "tables": []}


def _split_text(text: str, max_chars: int) -> list[str]:
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


def _encode_image(path: str) -> dict[str, Any]:
    """Read an image file and return a base64 data block for the API."""
    with open(path, "rb") as f:
        data = base64.standard_b64encode(f.read()).decode("ascii")

    suffix = path.rsplit(".", 1)[-1].lower()
    media = "image/png" if suffix == "png" else "image/jpeg"
    return {"type": "image", "source": {"type": "base64", "media_type": media, "data": data}}


# ---------------------------------------------------------------------------
# LLM call wrappers (text-only, with retry)
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


def _call_with_retry(caller, system: str, user: str, max_retries: int = 3, **kwargs: Any) -> str:
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
# Vision LLM call wrappers
# ---------------------------------------------------------------------------

def _call_claude_vision(
    system: str,
    user_text: str,
    images: list[dict[str, Any]],
    *,
    api_key: str | None = None,
    model: str = "claude-sonnet-4-6",
    max_tokens: int = 4096,
) -> str:
    from anthropic import Anthropic

    client = Anthropic(api_key=api_key or os.environ.get("ANTHROPIC_API_KEY"))

    content: list[dict[str, Any]] = [{"type": "text", "text": user_text}]
    content.extend(images)

    response = client.messages.create(
        model=model,
        max_tokens=max_tokens,
        system=system,
        messages=[{"role": "user", "content": content}],
    )
    return response.content[0].text


def _call_openai_vision(
    system: str,
    user_text: str,
    images: list[dict[str, Any]],
    *,
    api_key: str | None = None,
    base_url: str | None = None,
    model: str = "gpt-4o",
    max_tokens: int = 4096,
) -> str:
    from openai import OpenAI

    client = OpenAI(
        api_key=api_key or os.environ.get("OPENAI_API_KEY"),
        base_url=base_url or os.environ.get("OPENAI_BASE_URL"),
    )

    # OpenAI uses a different image format
    openai_images: list[dict[str, Any]] = []
    for img in images:
        src = img["source"]
        openai_images.append({
            "type": "image_url",
            "image_url": {"url": f"data:{src['media_type']};base64,{src['data']}"},
        })

    content: list[dict[str, Any]] = [{"type": "text", "text": user_text}]
    content.extend(openai_images)

    response = client.chat.completions.create(
        model=model,
        max_tokens=max_tokens,
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": content},
        ],
    )
    return response.choices[0].message.content


def _call_vision_with_retry(
    caller,
    system: str,
    user_text: str,
    images: list[dict[str, Any]],
    max_retries: int = 3,
    **kwargs: Any,
) -> str:
    last_error: Exception | None = None
    for attempt in range(max_retries):
        try:
            return caller(system, user_text, images, **kwargs)
        except Exception as e:
            last_error = e
            if attempt < max_retries - 1:
                time.sleep(2 ** attempt)
    raise last_error  # type: ignore


# ---------------------------------------------------------------------------
# Text-only map-reduce
# ---------------------------------------------------------------------------

def _summarize_raw(
    chapter_title: str,
    chapter_text: str,
    *,
    provider: str = "claude",
    **kwargs: Any,
) -> list[dict[str, Any]]:
    caller = _call_claude if provider == "claude" else _call_openai
    raw = _call_with_retry(
        caller,
        _SYSTEM_PROMPT,
        _USER_TEMPLATE.format(chapter_title=chapter_title, chapter_text=chapter_text),
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
    parts: list[str] = []
    for i, bs in enumerate(bullet_sets, 1):
        parts.append(f"段{i}: {json.dumps(bs, ensure_ascii=False)}")
    merged_text = "\n".join(parts)

    caller = _call_claude if provider == "claude" else _call_openai
    raw = _call_with_retry(
        caller, _MERGE_SYSTEM,
        _MERGE_USER.format(chapter_title=chapter_title, bullet_sets=merged_text),
        **kwargs,
    )
    return _parse_bullets(raw)


# ---------------------------------------------------------------------------
# Vision summarizer
# ---------------------------------------------------------------------------

def _summarize_with_visuals_raw(
    chapter_title: str,
    chapter_text: str,
    *,
    figures: list[dict[str, Any]] | None = None,
    tables: list[dict[str, Any]] | None = None,
    provider: str = "claude",
    **kwargs: Any,
) -> dict[str, Any]:
    """Single-shot vision call: text + screenshots → {bullets, figures, tables}."""
    figures = figures or []
    tables = tables or []

    if not figures and not tables:
        # No visuals — fall back to text-only
        bullets = _summarize_raw(chapter_title, chapter_text, provider=provider, **kwargs)
        return {"bullets": bullets, "figures": [], "tables": []}

    # Build the user message
    user_lines = [
        _VISION_USER_HEADER.format(
            chapter_title=chapter_title,
            chapter_text=chapter_text[:10000],
        ),
    ]

    idx = 0
    item_to_source: list[dict[str, Any]] = []  # {type: "figure"|"table", number, screenshot}

    for fig in figures:
        idx += 1
        user_lines.append(_VISION_ITEM.format(idx=idx, number=fig["number"], caption=fig["caption"]))
        item_to_source.append({"type": "figure", "number": fig["number"], "screenshot": fig.get("screenshot", "")})

    for tab in tables:
        idx += 1
        user_lines.append(_VISION_ITEM.format(idx=idx, number=tab["number"], caption=tab["caption"]))
        item_to_source.append({"type": "table", "number": tab["number"], "screenshot": tab.get("screenshot", "")})

    user_text = "\n".join(user_lines)

    # Encode images
    image_blocks: list[dict[str, Any]] = []
    for item in item_to_source:
        if item["screenshot"] and os.path.isfile(item["screenshot"]):
            try:
                image_blocks.append(_encode_image(item["screenshot"]))
            except Exception:
                image_blocks.append({"type": "text", "text": f"\n[图片 {item['number']} 无法加载]"})
        else:
            image_blocks.append({"type": "text", "text": f"\n[图片 {item['number']} 缺失]"})

    # Call vision API
    if provider == "claude":
        vision_caller = _call_claude_vision
    else:
        vision_caller = _call_openai_vision

    raw = _call_vision_with_retry(
        vision_caller, _VISION_SYSTEM, user_text, image_blocks, **kwargs,
    )
    return _parse_vision_result(raw)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def summarize_chapter(
    chapter_title: str,
    chapter_text: str,
    *,
    figures: list[dict[str, Any]] | None = None,
    tables: list[dict[str, Any]] | None = None,
    provider: str = "claude",
    max_chars: int = 10000,
    **kwargs: Any,
) -> dict[str, Any]:
    """Refine a chapter into slide content, optionally with figure/table screenshots.

    Args:
        chapter_title: Section heading.
        chapter_text: Raw text content (may be long — map-reduce kicks in).
        figures: List of ``{"number": "图3-1", "caption": "...", "screenshot": "path"}``.
        tables: List of ``{"number": "表3-1", "caption": "...", "screenshot": "path"}``.
        provider: ``"claude"`` or ``"openai"``.
        max_chars: Max chars per LLM request (for text; images are extra).
        **kwargs: Forwarded (model, api_key, max_tokens, etc.).

    Returns:
        ``{"bullets": [...], "figures": [...], "tables": [...]}``
    """
    if not chapter_text.strip() and not figures and not tables:
        return {"bullets": [], "figures": [], "tables": []}

    has_visuals = bool(figures or tables)

    # Short text + no visuals → simple text call (backwards compatible)
    if len(chapter_text) <= max_chars and not has_visuals:
        bullets = _summarize_raw(chapter_title, chapter_text, provider=provider, **kwargs)
        return {"bullets": bullets, "figures": [], "tables": []}

    # Short text + visuals → single vision call
    if len(chapter_text) <= max_chars and has_visuals:
        return _summarize_with_visuals_raw(
            chapter_title, chapter_text,
            figures=figures, tables=tables,
            provider=provider, **kwargs,
        )

    # Long text + no visuals → map-reduce text only
    if not has_visuals:
        chunks = _split_text(chapter_text, max_chars)
        all_bullets: list[list[dict[str, Any]]] = []
        for i, chunk in enumerate(chunks):
            chunk_title = f"{chapter_title}（第{i + 1}/{len(chunks)}部分）"
            bullets = _summarize_raw(chunk_title, chunk, provider=provider, **kwargs)
            if bullets:
                all_bullets.append(bullets)
        if not all_bullets:
            return {"bullets": [], "figures": [], "tables": []}
        if len(all_bullets) == 1:
            return {"bullets": all_bullets[0], "figures": [], "tables": []}
        merged = _merge_bullets(chapter_title, all_bullets, provider=provider, **kwargs)
        return {"bullets": merged, "figures": [], "tables": []}

    # Long text + visuals → map-reduce text, then single vision call for visuals
    chunks = _split_text(chapter_text, max_chars)
    all_bullets: list[list[dict[str, Any]]] = []
    for i, chunk in enumerate(chunks):
        chunk_title = f"{chapter_title}（第{i + 1}/{len(chunks)}部分）"
        bullets = _summarize_raw(chunk_title, chunk, provider=provider, **kwargs)
        if bullets:
            all_bullets.append(bullets)

    if len(all_bullets) == 1:
        bullets = all_bullets[0]
    elif all_bullets:
        bullets = _merge_bullets(chapter_title, all_bullets, provider=provider, **kwargs)
    else:
        bullets = []

    # Now handle visuals with a short summary + screenshots
    bullet_summary = json.dumps(bullets, ensure_ascii=False)
    vision_text = f"该章节bullet points：\n{bullet_summary}\n\n以下为该章节内的图表截图，请逐一审视："

    result = _summarize_with_visuals_raw(
        chapter_title, vision_text,
        figures=figures, tables=tables,
        provider=provider, **kwargs,
    )
    # Keep the text-only bullets (already high quality from map-reduce)
    result["bullets"] = bullets
    return result


# ---------------------------------------------------------------------------
# Chapter-level multi-slide summarization
# ---------------------------------------------------------------------------

def summarize_chapter_multi(
    chapter_title: str,
    chapter_text: str,
    *,
    figures: list[dict[str, Any]] | None = None,
    tables: list[dict[str, Any]] | None = None,
    provider: str = "claude",
    **kwargs: Any,
) -> dict[str, Any]:
    """Summarize a whole chapter into 2-4 slides + figure/table decisions.

    Returns:
        ``{"slides": [{"title": str, "bullets": [...]}, ...],
           "figures": [...], "tables": [...]}``
    """
    if not chapter_text.strip() and not figures and not tables:
        return {"slides": [], "figures": [], "tables": []}

    has_visuals = bool(figures or tables)

    if has_visuals:
        # Build message with figure/table list + screenshots
        user_lines = [
            _CHAPTER_USER_TEMPLATE.format(
                chapter_title=chapter_title,
                chapter_text=chapter_text[:10000],
            ),
        ]

        idx = 0
        item_to_source: list[str] = []
        for fig in (figures or []):
            idx += 1
            user_lines.append(_VISION_ITEM.format(idx=idx, number=fig["number"], caption=fig["caption"]))
            item_to_source.append(fig.get("screenshot", ""))
        for tab in (tables or []):
            idx += 1
            user_lines.append(_VISION_ITEM.format(idx=idx, number=tab["number"], caption=tab["caption"]))
            item_to_source.append(tab.get("screenshot", ""))

        user_text = "\n".join(user_lines)

        image_blocks: list[dict[str, Any]] = []
        for path in item_to_source:
            if path and os.path.isfile(path):
                try:
                    image_blocks.append(_encode_image(path))
                except Exception:
                    image_blocks.append({"type": "text", "text": "\n[图片无法加载]"})

        vision_caller = _call_claude_vision if provider == "claude" else _call_openai_vision
        raw = _call_vision_with_retry(
            vision_caller, _CHAPTER_SYSTEM, user_text, image_blocks, **kwargs,
        )
        return _parse_chapter_result(raw)
    else:
        caller = _call_claude if provider == "claude" else _call_openai
        raw = _call_with_retry(
            caller, _CHAPTER_SYSTEM,
            _CHAPTER_USER_TEMPLATE.format(chapter_title=chapter_title, chapter_text=chapter_text),
            **kwargs,
        )
        return _parse_chapter_result(raw)
