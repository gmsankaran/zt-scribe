"""
ZT Scribe pipeline — importable core.

Two stages:
  extract(image_bytes, media_type) -> dict   — one vision call, flat item list
  render(board)                    -> str    — pure Python, person-organised markdown
"""

import base64
import json
from collections import defaultdict

from anthropic import Anthropic
from dotenv import load_dotenv

load_dotenv()

MODEL = "claude-sonnet-5"

KNOWN_OWNERS = ["DS", "SG", "KS", "MA", "GS"]

GLOSSARY = """
Owners (initials): DS, SG, KS, MA, GS
Other people mentioned: PR (Projjol), AH, AS, MK, CK, Ankita, Karthik, Dimple, Sara, Vandhana
Domain terms: TA, DCM, CoMa, PBAP, HFP, A2DP, AVRCP, MIB4, ASPICE, ePN, Tabnine,
  CL56, CL8MIN, DE visit, KT, TC, PoC, SW
"""

EXTRACT_PROMPT = """You are reading a photograph of a whiteboard from a weekly team review.

The board has three columns: Progress, Plans, Pitfalls.
The Progress column is annotated "Plans from last week" — it is last week's
plans being reviewed for status, not a free-form list of accomplishments.

Return a flat JSON list of every distinct item on the board. Do not reorganize,
do not group by person, do not summarize. One object per bullet.

Each object:
  "text":     the item as written, expanded to readable English
  "column":   "progress" | "plans" | "pitfalls"
  "owner":    initials ONLY if written on or immediately beside the item.
              null otherwise. Do NOT infer the owner from context, from
              neighbouring items, or from who is likely to do the work.
              An empty owner is correct and useful; a guessed owner is a bug.
  "markers":  any of "circled", "boxed", "asterisk", "double-asterisk". [] if none.
  "due":      any date or day mentioned, else null
  "confidence": 0.0-1.0, your confidence you read the handwriting correctly

Then a top-level "unreadable" list: short descriptions of anything you could
not make out at all.

Ignore arrows between items — they are unreliable in a photograph.

Glossary. These are the correct spellings. If a word is close to one of these,
it IS that word:
{glossary}

Respond with JSON only. No preamble, no markdown fences.
Shape: {{"items": [...], "unreadable": [...]}}
"""


def extract(image_bytes: bytes, media_type: str = "image/jpeg") -> dict:
    data = base64.standard_b64encode(image_bytes).decode()

    client = Anthropic()
    resp = client.messages.create(
        model=MODEL,
        max_tokens=8000,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image",
                 "source": {"type": "base64", "media_type": media_type, "data": data}},
                {"type": "text",
                 "text": EXTRACT_PROMPT.format(glossary=GLOSSARY)},
            ],
        }],
    )

    raw = "".join(b.text for b in resp.content if b.type == "text").strip()
    raw = raw.removeprefix("```json").removeprefix("```").removesuffix("```").strip()
    try:
        return json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValueError(f"Model did not return valid JSON: {raw[:600]}") from exc


def render(board: dict) -> str:
    items = board.get("items", [])

    by_owner: dict = defaultdict(lambda: defaultdict(list))
    unattributed: dict = defaultdict(list)

    for item in items:
        owner = item.get("owner")
        if owner not in KNOWN_OWNERS:
            unattributed[item.get("column", "?")].append(item)
        else:
            by_owner[owner][item.get("column", "?")].append(item)

    def bullet(item) -> str:
        line = item["text"]
        if item.get("due"):
            line += f" ({item['due']})"
        if item.get("confidence", 1.0) < 0.7:
            line += "  `[?]`"
        if item.get("markers"):
            line += f"  _{', '.join(item['markers'])}_"
        return f"* {line}"

    out = ["# Minutes\n"]

    for owner in KNOWN_OWNERS:
        if owner not in by_owner:
            continue
        out.append(f"\n## {owner}\n")
        for label, col in (("Progress", "progress"), ("Next week's focus", "plans")):
            if by_owner[owner][col]:
                out.append(f"**{label}**\n")
                out += [bullet(i) for i in by_owner[owner][col]]
                out.append("")

    if any(unattributed.values()):
        out.append("\n---\n\n## Unattributed — assign before circulating\n")
        for col, group in unattributed.items():
            out.append(f"\n**{col.title()}**\n")
            out += [bullet(i) for i in group]

    if board.get("unreadable"):
        out.append("\n---\n\n## Could not read\n")
        out += [f"* {u}" for u in board["unreadable"]]

    return "\n".join(out)
