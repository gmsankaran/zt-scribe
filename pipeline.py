"""
ZT Scribe pipeline.

Loads team knowledge from team_context.json (tracked in git, never hardcoded here).
Public API used by bot.py:
  load_context()                        → dict
  resolve_template(ctx, key)            → dict   (merged ctx with template settings)
  extract(img, mime, ctx)               → board dict
  render(board, ctx)                    → markdown string
  build_template_card(ctx)              → Adaptive Card dict | None
  build_clarification_card(board, ctx)  → Adaptive Card dict | None
"""

import base64
import json
import sys
from collections import defaultdict
from pathlib import Path

import anthropic

# ---------------------------------------------------------------------------
# Context loading
# ---------------------------------------------------------------------------

_CONTEXT_PATH = Path(__file__).parent / "team_context.json"

_FALLBACK_CTX: dict = {
    "members": [],
    "default_template": "friday_fifty",
    "templates": {
        "friday_fifty": {
            "name": "Friday Fifty",
            "description": "Weekly standup with Progress, Plans, and Pitfalls columns per person.",
            "column_order": ["progress", "plans", "pitfalls"],
            "column_labels": {"progress": "Progress", "plans": "Plans", "pitfalls": "Pitfalls"},
        }
    },
    "glossary": {},
    "people": {},
    "confirmed_expansions": {},
}


def load_context(path: str | Path = _CONTEXT_PATH) -> dict:
    """Load team_context.json, stripping comment keys (prefixed _)."""
    try:
        with open(path, encoding="utf-8") as fh:
            data = json.load(fh)
        return {k: v for k, v in data.items() if not k.startswith("_")}
    except FileNotFoundError:
        print(f"[pipeline] team_context.json not found at {path} — using defaults", file=sys.stderr)
        return dict(_FALLBACK_CTX)
    except json.JSONDecodeError as exc:
        print(f"[pipeline] team_context.json is invalid JSON: {exc}", file=sys.stderr)
        return dict(_FALLBACK_CTX)


def resolve_template(ctx: dict, template_key: str | None = None) -> dict:
    """Return a copy of ctx with the chosen template's column settings merged in."""
    templates = ctx.get("templates", {})
    key  = template_key or ctx.get("default_template", "friday_fifty")
    tmpl = templates.get(key) or templates.get("friday_fifty") or {}
    merged = dict(ctx)
    merged["_template_key"]         = key
    merged["_template_name"]        = tmpl.get("name", key)
    merged["_template_description"] = tmpl.get("description", "")
    merged["column_order"]          = tmpl.get("column_order",  ctx.get("column_order",  ["progress", "plans", "pitfalls"]))
    merged["column_labels"]         = tmpl.get("column_labels", ctx.get("column_labels", {}))
    return merged


# ---------------------------------------------------------------------------
# Extraction
# ---------------------------------------------------------------------------

def _build_prompt(ctx: dict) -> str:
    members      = ctx.get("members", [])
    col_order    = ctx.get("column_order", ["progress", "plans", "pitfalls"])
    glossary     = ctx.get("glossary", {})
    people       = ctx.get("people", {})
    confirmed    = ctx.get("confirmed_expansions", {})
    tmpl_name    = ctx.get("_template_name", "this meeting")
    tmpl_desc    = ctx.get("_template_description", "")

    members_str   = ", ".join(members) if members else "(none listed)"
    col_keys      = " | ".join(f'"{c}"' for c in col_order)
    glossary_str  = "\n".join(f"  {k}: {v}" for k, v in glossary.items()) if glossary else "  (none)"
    people_str    = "\n".join(f"  {k}: {v}" for k, v in people.items() if v) if people else "  (none)"
    confirmed_str = "\n".join(f"  {k} → {v}" for k, v in confirmed.items()) if confirmed else "  (none)"

    return f"""You are reading a whiteboard from a {tmpl_name} meeting.

MEETING FORMAT:
{tmpl_desc if tmpl_desc else f"The board has {len(col_order)} columns: {', '.join(col_order)}."}

PARTICIPANTS (initials that appear as section headers): {members_str}

OWNERSHIP RULES — read carefully:
  1. Within each column, person initials appear as bold section headers (e.g. "DS" written
     prominently). Every item below that header, until the next person's header, belongs to
     that person (owner_source: "section").
  2. If initials are written explicitly on or beside a specific item, those override the
     section header for that item (owner_source: "explicit").
  3. Items can have MULTIPLE owners — capture all of them:
       • Comma notation  (DS, KS)  → owners: ["DS","KS"],  ownership_notation: ","
       • Arrow notation  (GS → SS) → owners: ["GS","SS"],  ownership_notation: "→"
       • Single owner               → owners: ["DS"],       ownership_notation: null
       • Genuinely unattributed     → owners: [],           ownership_notation: null
  4. Sub-bullets (indented items) inherit owners from their parent item.

KNOWN ABBREVIATIONS — for OCR disambiguation only; keep terms as-is in output, do NOT expand:
{glossary_str}

KNOWN PEOPLE (non-team names that may appear inside item text):
{people_str}

CONFIRMED EXPANSIONS FROM PREVIOUS SESSIONS:
{confirmed_str}

Return a JSON object with this exact structure:
{{
  "items": [
    {{
      "text":               "<item as written — abbreviations unchanged>",
      "column":             {col_keys},
      "owners":             ["<initials>", ...],
      "ownership_notation": "," | "→" | null,
      "owner_source":       "explicit" | "section" | null,
      "parent_index":       <0-based index of parent item, or null if top-level>,
      "markers":            ["circled" | "boxed" | "asterisk" | "double-asterisk"],
      "due":                "<date or deadline, or null>",
      "confidence":         <0.0–1.0>
    }}
  ],
  "unreadable": ["<description of any section you could not decipher>"]
}}

Items must appear top-to-bottom, left-to-right as on the board.
Output raw JSON only — no markdown fences, no commentary.
"""


def extract(image_bytes: bytes, media_type: str = "image/jpeg", ctx: dict | None = None) -> dict:
    """Send the whiteboard image to Claude and return the parsed board dict.
    The caller should pass a ctx already resolved via resolve_template()."""
    if ctx is None:
        ctx = resolve_template(load_context())

    client   = anthropic.Anthropic()
    b64_data = base64.standard_b64encode(image_bytes).decode("utf-8")
    prompt   = _build_prompt(ctx)

    model      = ctx.get("model", "claude-sonnet-5")
    max_tokens = ctx.get("max_tokens", 16000)

    create_kwargs: dict = dict(
        model=model,
        max_tokens=max_tokens,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": b64_data}},
                {"type": "text",  "text": prompt},
            ],
        }],
    )

    # Sonnet 5 / Opus 5 use adaptive thinking. "low" effort is right for OCR —
    # perception task, not deep reasoning.
    _THINKING_MODELS = {"claude-sonnet-5", "claude-opus-5", "claude-fable-5"}
    if model in _THINKING_MODELS:
        effort = ctx.get("thinking_effort", "low")
        create_kwargs["thinking"]      = {"type": "adaptive"}
        create_kwargs["output_config"] = {"effort": effort}

    response = client.messages.create(**create_kwargs)

    # Sonnet 5 may prepend ThinkingBlock(s) — find the first TextBlock.
    text_block = next((b for b in response.content if b.type == "text"), None)
    if text_block is None:
        types = [b.type for b in response.content]
        raise ValueError(f"Model returned no text block (got: {types})")

    raw = text_block.text.strip()
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

_LOW_CONF_RENDER = 0.5   # only flag clearly uncertain items in the draft


def render(board: dict, ctx: dict | None = None) -> str:
    """Render the board dict to Markdown."""
    if ctx is None:
        ctx = resolve_template(load_context())

    items      = board.get("items", [])
    members    = ctx.get("members", [])
    col_order  = ctx.get("column_order", ["progress", "plans", "pitfalls"])
    col_labels = ctx.get("column_labels", {})
    item_map   = {i: item for i, item in enumerate(items)}

    def _primary_owner(idx: int, visited: set | None = None) -> str | None:
        """Walk up the parent chain; return the first element of the owners list."""
        if visited is None:
            visited = set()
        if idx in visited:
            return None
        visited.add(idx)
        item = item_map.get(idx)
        if item is None:
            return None
        owners = item.get("owners") or ([item["owner"]] if item.get("owner") else [])
        if owners:
            return owners[0]
        parent = item.get("parent_index")
        if parent is not None:
            return _primary_owner(parent, visited)
        return None

    # Bucket items: by_owner[owner][col] = [(idx, item)]
    by_owner: dict[str, dict[str, list]] = {m: defaultdict(list) for m in members}
    unattr: dict[str, list]              = defaultdict(list)

    for idx, item in enumerate(items):
        owner = _primary_owner(idx)
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
        if item.get("confidence", 1.0) < _LOW_CONF_RENDER:
            text += " `[?]`"
        tags = item.get("markers", [])
        if tags:
            text += "  _%s_" % ", ".join(tags)
        return f"{indent}* {text}"

    def _render_tree(idx_items: list) -> list[str]:
        idx_set   = {i for i, _ in idx_items}
        top_level = [(i, it) for i, it in idx_items if it.get("parent_index") not in idx_set]
        children  = defaultdict(list)
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

    out: list[str] = ["# Minutes\n"]

    # Column-first: Progress → Plans → Pitfalls, members nested within each column.
    for col in col_order:
        label   = col_labels.get(col, col.title())
        has_any = any(by_owner.get(owner, {}).get(col) for owner in members)
        if not has_any:
            continue
        out.append(f"\n## {label}\n")
        for owner in members:
            group = by_owner.get(owner, {}).get(col, [])
            if not group:
                continue
            out.append(f"**{owner}**\n")
            out.extend(_render_tree(group))
            out.append("")

    if any(unattr.get(c) for c in col_order):
        out.append("\n---\n\n## Unattributed — assign before circulating\n")
        for col in col_order:
            group = unattr.get(col, [])
            if not group:
                continue
            label = col_labels.get(col, col.title())
            out.append(f"\n**{label}**\n")
            out.extend(_render_tree(group))

    if board.get("unreadable"):
        out.append("\n---\n\n## Could not read\n")
        out += [f"* {u}" for u in board["unreadable"]]

    return "\n".join(out)


# ---------------------------------------------------------------------------
# Adaptive Cards
# ---------------------------------------------------------------------------

_LOW_CONF_CARD = 0.7
_MAX_UNATTR    = 10
_MAX_LOW       = 5


def build_template_card(ctx: dict) -> dict | None:
    """Return a template-selection card when more than one template is available."""
    templates = ctx.get("templates", {})
    if len(templates) <= 1:
        return None

    default_key = ctx.get("default_template", next(iter(templates)))
    choices     = [{"title": t.get("name", k), "value": k} for k, t in templates.items()]
    desc_lines  = [
        f"**{t.get('name', k)}**: {t['description']}"
        for k, t in templates.items() if t.get("description")
    ]

    body: list[dict] = [
        {"type": "TextBlock", "text": "ZT Scribe — Select meeting format",
         "weight": "Bolder", "size": "Medium"},
        {"type": "TextBlock", "text": "Which template should I use for these minutes?", "wrap": True},
        {"type": "Input.ChoiceSet", "id": "template_key", "style": "compact",
         "value": default_key, "choices": choices},
    ]
    if desc_lines:
        body.append({"type": "TextBlock", "text": "\n\n".join(desc_lines),
                     "wrap": True, "isSubtle": True, "spacing": "Small"})

    return {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.2",
        "body": body,
        "actions": [{
            "type": "Action.Submit",
            "title": "Generate minutes →",
            "data": {
                "msteams": {"type": "messageBack", "text": "zt_template"},
                "zt_action": "select_template",
            },
        }],
    }


def build_clarification_card(board: dict, ctx: dict | None = None) -> dict | None:
    """Return a review card for unattributed or low-confidence items, or None."""
    if ctx is None:
        ctx = load_context()

    items    = board.get("items", [])
    members  = ctx.get("members", [])
    item_map = {i: item for i, item in enumerate(items)}

    def _primary_owner(idx, visited=None):
        if visited is None:
            visited = set()
        if idx in visited:
            return None
        visited.add(idx)
        item = item_map.get(idx)
        if item is None:
            return None
        owners = item.get("owners") or ([item["owner"]] if item.get("owner") else [])
        if owners:
            return owners[0]
        parent = item.get("parent_index")
        if parent is not None:
            return _primary_owner(parent, visited)
        return None

    # Only surface top-level unattributed items (sub-bullets inherit from parent).
    unattr = [(i, it) for i, it in enumerate(items)
              if _primary_owner(i) not in members and it.get("parent_index") is None]
    unattr_indices = {i for i, _ in unattr}

    # Exclude unattributed items from the low-confidence section — they already
    # appear above with owner dropdowns; showing them again as text edits is confusing.
    low_conf = [(i, it) for i, it in enumerate(items)
                if it.get("confidence", 1.0) < _LOW_CONF_CARD
                and i not in unattr_indices]

    if not unattr and not low_conf:
        return None

    member_choices = [{"title": m, "value": m} for m in members]

    body: list[dict] = [
        {"type": "TextBlock", "text": "ZT Scribe — Please review before circulating",
         "weight": "Bolder", "size": "Medium"},
    ]

    if unattr:
        body.append({
            "type": "TextBlock",
            "text": f"{min(len(unattr), _MAX_UNATTR)} unattributed item(s) — assign owners:",
            "weight": "Bolder", "spacing": "Medium", "wrap": True,
        })
        for i, it in unattr[:_MAX_UNATTR]:
            col_label = it.get("column", "?").title()
            body.append({
                "type": "Container", "spacing": "Small",
                "items": [
                    {"type": "TextBlock",
                     "text": f"{it['text']}  _[{col_label}]_",
                     "wrap": True, "isSubtle": True},
                    # Multi-select dropdown — picks "DS,MA" for joint ownership
                    {"type": "Input.ChoiceSet", "id": f"owner_{i}",
                     "isMultiSelect": True,
                     "placeholder": "Select owner(s)…", "choices": member_choices},
                    # Free-text override — use for names not in the list, or "GS → SS"
                    {"type": "Input.Text", "id": f"owner_text_{i}",
                     "placeholder": "Or type freely (overrides dropdown)"},
                ],
            })

    if low_conf:
        body.append({
            "type": "TextBlock",
            "text": f"{min(len(low_conf), _MAX_LOW)} uncertain reading(s) — verify or correct:",
            "weight": "Bolder", "spacing": "Medium", "wrap": True,
        })
        for i, it in low_conf[:_MAX_LOW]:
            col_label  = it.get("column", "?").title()
            owners_str = ", ".join(
                it.get("owners") or ([it["owner"]] if it.get("owner") else ["unattributed"])
            )
            body.append({
                "type": "Container", "spacing": "Small",
                "items": [
                    {"type": "TextBlock", "text": f"{col_label} / {owners_str}",
                     "isSubtle": True, "wrap": True},
                    {"type": "Input.Text", "id": f"text_{i}",
                     "value": it["text"], "placeholder": it["text"]},
                ],
            })

    payload_b64 = base64.b64encode(json.dumps(board).encode("utf-8")).decode("ascii")

    return {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.2",
        "body": body,
        "actions": [
            {
                "type": "Action.Submit", "title": "✅ Submit corrections",
                "data": {
                    "msteams": {"type": "messageBack", "text": "zt_finalize"},
                    "zt_action": "finalize_minutes",
                    "board_b64": payload_b64,
                },
            },
            {
                "type": "Action.Submit", "title": "Skip — draft looks good",
                "data": {
                    "msteams": {"type": "messageBack", "text": "zt_skip"},
                    "zt_action": "skip_review",
                },
            },
        ],
    }
