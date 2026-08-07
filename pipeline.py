"""
ZT Scribe pipeline.

Loads team knowledge from team_context.json (tracked in git, never hardcoded here).
Three entry points used by bot.py:
  - load_context()              → dict
  - extract(img, mime, ctx)     → board dict
  - render(board, ctx)          → markdown string
  - build_clarification_card()  → Adaptive Card dict | None
"""

import base64
import json
import os
import sys
from collections import defaultdict
from pathlib import Path

import anthropic

# ---------------------------------------------------------------------------
# Context loading
# ---------------------------------------------------------------------------

_CONTEXT_PATH = Path(__file__).parent / "team_context.json"

_FALLBACK_CONTEXT = {
    "members": [],
    "column_order": ["progress", "plans", "pitfalls"],
    "column_labels": {"progress": "Progress", "plans": "Plans", "pitfalls": "Pitfalls"},
    "glossary": {},
    "people": {},
    "confirmed_expansions": {},
}


def load_context(path: str | Path = _CONTEXT_PATH) -> dict:
    """Load team_context.json.  Falls back silently if the file is absent."""
    try:
        with open(path, encoding="utf-8") as fh:
            data = json.load(fh)
        # Strip comment-only keys (prefixed with _)
        return {k: v for k, v in data.items() if not k.startswith("_")}
    except FileNotFoundError:
        print(f"[pipeline] team_context.json not found at {path} — using defaults", file=sys.stderr)
        return dict(_FALLBACK_CONTEXT)
    except json.JSONDecodeError as exc:
        print(f"[pipeline] team_context.json is invalid JSON: {exc}", file=sys.stderr)
        return dict(_FALLBACK_CONTEXT)


# ---------------------------------------------------------------------------
# Extraction
# ---------------------------------------------------------------------------

def _build_prompt(ctx: dict) -> str:
    members   = ctx.get("members", [])
    col_order = ctx.get("column_order", ["progress", "plans", "pitfalls"])
    glossary  = ctx.get("glossary", {})
    people    = ctx.get("people", {})
    confirmed = ctx.get("confirmed_expansions", {})

    members_str = ", ".join(members) if members else "(none listed)"
    col_keys    = " | ".join(f'"{c}"' for c in col_order)
    glossary_str = "\n".join(f"  {k}: {v}" for k, v in glossary.items()) if glossary else "  (none)"
    people_str   = "\n".join(f"  {k}: {v}" for k, v in people.items() if v) if people else "  (none)"
    confirmed_str = "\n".join(f"  {k} → {v}" for k, v in confirmed.items()) if confirmed else "  (none)"

    return f"""You are reading a whiteboard from a Friday Fifty meeting.
The meeting has {len(members)} participants with these initials: {members_str}.

The board has three columns in this order: {", ".join(col_order)}.

OWNERSHIP RULES — read carefully:
  1. Within each column, person initials appear as section headers (e.g. "DS" or "KS" written
     prominently). Every item below that header, until the next person's header, belongs to
     that person (owner_source: "section").
  2. If initials are written explicitly on or immediately beside a specific item, that overrides
     the section header for that item (owner_source: "explicit").
  3. Only set owner to null for items in a clearly shared/general area not under any person's
     section, or items with no discernible section context.
  4. Sub-bullets (indented items under a parent bullet) inherit the owner of their parent.

KNOWN ABBREVIATIONS — use these to resolve ambiguous handwriting:
{glossary_str}

KNOWN PEOPLE MENTIONED (non-team names that may appear in items):
{people_str}

CONFIRMED EXPANSIONS FROM PREVIOUS SESSIONS:
{confirmed_str}

Return a JSON object with this exact structure:
{{
  "items": [
    {{
      "text":         "<the item as written — keep abbreviations as-is, do not expand them>",
      "column":       {col_keys},
      "owner":        "<initials from {members_str}> | null",
      "owner_source": "explicit" | "section" | null,
      "parent_index": <0-based index of parent item, or null if top-level>,
      "markers":      ["circled" | "boxed" | "asterisk" | "double-asterisk"],
      "due":          "<any date or deadline mentioned, or null>",
      "confidence":   <0.0–1.0 confidence you read the handwriting correctly>
    }}
  ],
  "unreadable": ["<description of any section you could not decipher>"]
}}

Items must appear in top-to-bottom, left-to-right order as they appear on the board.
Output raw JSON only — no markdown fences, no commentary.
"""


def extract(image_bytes: bytes, media_type: str = "image/jpeg", ctx: dict | None = None) -> dict:
    """Send the whiteboard image to Claude and return the parsed board dict."""
    if ctx is None:
        ctx = load_context()

    client   = anthropic.Anthropic()
    b64_data = base64.standard_b64encode(image_bytes).decode("utf-8")
    prompt   = _build_prompt(ctx)

    response = client.messages.create(
        model="claude-sonnet-5",
        max_tokens=4096,
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": b64_data}},
                    {"type": "text",  "text": prompt},
                ],
            }
        ],
    )

    raw = response.content[0].text.strip()
    # Strip accidental markdown fences
    if raw.startswith("```"):
        raw = raw.split("\n", 1)[1] if "\n" in raw else raw[3:]
    if raw.endswith("```"):
        raw = raw.rsplit("```", 1)[0]

    board = json.loads(raw)
    print(f"[pipeline] extracted {len(board.get('items', []))} items", file=sys.stderr)
    return board


# ---------------------------------------------------------------------------
# Rendering
# ---------------------------------------------------------------------------

def render(board: dict, ctx: dict | None = None) -> str:
    """Render the board dict to a Markdown string."""
    if ctx is None:
        ctx = load_context()

    items     = board.get("items", [])
    members   = ctx.get("members", [])
    col_order = ctx.get("column_order", ["progress", "plans", "pitfalls"])
    col_labels = ctx.get("column_labels", {})

    # Index by position for parent lookup
    item_map = {i: item for i, item in enumerate(items)}

    def _effective_owner(idx: int, visited: set | None = None) -> str | None:
        """Walk up the parent chain to find the nearest explicit or section owner."""
        if visited is None:
            visited = set()
        if idx in visited:
            return None
        visited.add(idx)
        item  = item_map.get(idx)
        if item is None:
            return None
        owner = item.get("owner")
        if owner:
            return owner
        parent = item.get("parent_index")
        if parent is not None:
            return _effective_owner(parent, visited)
        return None

    # Bucket: by_owner[owner][col] = list of (idx, item)
    # unattr[col]              = list of (idx, item)
    by_owner: dict[str, dict[str, list]] = {m: defaultdict(list) for m in members}
    unattr: dict[str, list]              = defaultdict(list)

    for idx, item in enumerate(items):
        owner = _effective_owner(idx)
        col   = item.get("column", "?")
        if owner in members:
            by_owner[owner][col].append((idx, item))
        else:
            unattr[col].append((idx, item))

    def _fmt_bullet(item: dict, depth: int = 0) -> str:
        indent = "    " * depth
        text   = item["text"]
        if item.get("due"):
            text += f" _{item['due']}_"
        if item.get("confidence", 1.0) < 0.7:
            text += " `[?]`"
        tags = item.get("markers", [])
        if tags:
            text += "  _%s_" % ", ".join(tags)
        return f"{indent}* {text}"

    def _render_item_tree(idx_items: list) -> list[str]:
        """Render a list of (idx, item) pairs with sub-bullets nested under parents."""
        idx_set = {i for i, _ in idx_items}
        # Top-level: parent not in this column's set
        top_level  = [(i, it) for i, it in idx_items if it.get("parent_index") not in idx_set]
        children   = defaultdict(list)
        for i, it in idx_items:
            p = it.get("parent_index")
            if p in idx_set:
                children[p].append((i, it))

        lines: list[str] = []
        def _walk(i, it, depth):
            lines.append(_fmt_bullet(it, depth))
            for ci, child in children.get(i, []):
                _walk(ci, child, depth + 1)

        for i, it in top_level:
            _walk(i, it, 0)
        return lines

    # ---------------------------------------------------------------------------
    # Build output
    # ---------------------------------------------------------------------------
    out: list[str] = ["# Minutes\n"]

    for owner in members:
        cols = by_owner.get(owner, {})
        has_any = any(cols.get(c) for c in col_order)
        if not has_any:
            continue
        out.append(f"\n## {owner}\n")
        for col in col_order:
            group = cols.get(col, [])
            if not group:
                continue
            label = col_labels.get(col, col.title())
            out.append(f"**{label}**\n")
            out.extend(_render_item_tree(group))
            out.append("")

    if any(unattr.get(c) for c in col_order):
        out.append("\n---\n\n## Unattributed — assign before circulating\n")
        for col in col_order:
            group = unattr.get(col, [])
            if not group:
                continue
            label = col_labels.get(col, col.title())
            out.append(f"\n**{label}**\n")
            out.extend(_render_item_tree(group))

    if board.get("unreadable"):
        out.append("\n---\n\n## Could not read\n")
        out += [f"* {u}" for u in board["unreadable"]]

    return "\n".join(out)


# ---------------------------------------------------------------------------
# Adaptive card for clarification
# ---------------------------------------------------------------------------

_LOW_CONF = 0.7
_MAX_UNATTR = 10   # cap cards at this many unattributed items
_MAX_LOW    = 5    # cap cards at this many low-confidence items


def build_clarification_card(board: dict, ctx: dict | None = None) -> dict | None:
    """Return a Teams Adaptive Card dict for items needing review, or None."""
    if ctx is None:
        ctx = load_context()

    items   = board.get("items", [])
    members = ctx.get("members", [])

    # Only surface top-level unattributed items (sub-bullets inherit from parent)
    item_map = {i: item for i, item in enumerate(items)}

    def _effective_owner(idx, visited=None):
        if visited is None:
            visited = set()
        if idx in visited:
            return None
        visited.add(idx)
        item  = item_map.get(idx)
        if item is None:
            return None
        owner = item.get("owner")
        if owner:
            return owner
        parent = item.get("parent_index")
        if parent is not None:
            return _effective_owner(parent, visited)
        return None

    unattr   = [(i, it) for i, it in enumerate(items)
                if _effective_owner(i) not in members and it.get("parent_index") is None]
    low_conf = [(i, it) for i, it in enumerate(items)
                if it.get("confidence", 1.0) < _LOW_CONF]

    if not unattr and not low_conf:
        return None

    body: list[dict] = [
        {
            "type": "TextBlock",
            "text": "ZT Scribe — Please review before circulating",
            "weight": "Bolder",
            "size": "Medium",
        }
    ]

    member_choices = [{"title": m, "value": m} for m in members]
    member_choices.append({"title": "(skip)", "value": ""})

    if unattr:
        body.append({
            "type": "TextBlock",
            "text": f"{len(unattr)} unattributed item(s) — assign owners:",
            "weight": "Bolder",
            "spacing": "Medium",
            "wrap": True,
        })
        for i, it in unattr[:_MAX_UNATTR]:
            col_label = it.get("column", "?").title()
            body.append({
                "type": "Container",
                "spacing": "Small",
                "items": [
                    {
                        "type": "TextBlock",
                        "text": f"{it['text']}  _[{col_label}]_",
                        "wrap": True,
                        "isSubtle": True,
                    },
                    {
                        "type": "Input.ChoiceSet",
                        "id": f"owner_{i}",
                        "placeholder": "Assign to…",
                        "choices": member_choices,
                    },
                ],
            })

    if low_conf:
        body.append({
            "type": "TextBlock",
            "text": f"{len(low_conf)} low-confidence reading(s) — verify or correct:",
            "weight": "Bolder",
            "spacing": "Medium",
            "wrap": True,
        })
        for i, it in low_conf[:_MAX_LOW]:
            col_label = it.get("column", "?").title()
            owner_lbl = it.get("owner") or "unattributed"
            body.append({
                "type": "Container",
                "spacing": "Small",
                "items": [
                    {
                        "type": "TextBlock",
                        "text": f"{col_label} / {owner_lbl}",
                        "isSubtle": True,
                        "wrap": True,
                    },
                    {
                        "type": "Input.Text",
                        "id": f"text_{i}",
                        "value": it["text"],
                        "placeholder": it["text"],
                    },
                ],
            })

    # Embed the full board JSON as base64 to avoid JSON-in-JSON escaping issues
    payload_b64 = base64.b64encode(json.dumps(board).encode("utf-8")).decode("ascii")

    card = {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.2",
        "body": body,
        "actions": [
            {
                "type": "Action.Submit",
                "title": "✅ Submit corrections",
                "data": {
                    "msteams": {"type": "messageBack", "text": "zt_finalize"},
                    "zt_action": "finalize_minutes",
                    "board_b64": payload_b64,
                },
            },
            {
                "type": "Action.Submit",
                "title": "Skip — draft looks good",
                "data": {
                    "msteams": {"type": "messageBack", "text": "zt_skip"},
                    "zt_action": "skip_review",
                },
            },
        ],
    }

    return card
