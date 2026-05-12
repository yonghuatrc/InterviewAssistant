#!/usr/bin/env python3
"""
experience_extractor.py — InterviewGenerator Phase 2
Rule-based claim extraction from ResumeScanner candidate JSON.
No LLM needed for extraction — keeps it fast and deterministic.
LLM is only used for question generation (in generate_interview.py).
"""
import re
import json
from typing import List, Dict, Any

# ─── Claim types ───────────────────────────────────────────────────────────────
CLAIM_TYPE_QUANTIFIED   = "quantified"
CLAIM_TYPE_PROJECT      = "project"
CLAIM_TYPE_SYSTEM       = "system"
CLAIM_TYPE_LEADERSHIP  = "leadership"
CLAIM_TYPE_INSTITUTION = "institution"
CLAIM_TYPE_GAP         = "gap"

CLAIM_PRIORITY = {
    CLAIM_TYPE_QUANTIFIED:   0,
    CLAIM_TYPE_LEADERSHIP:   1,
    CLAIM_TYPE_SYSTEM:       2,
    CLAIM_TYPE_PROJECT:      3,
    CLAIM_TYPE_INSTITUTION:  4,
    CLAIM_TYPE_GAP:          5,
}

# ─── Extraction helpers ────────────────────────────────────────────────────────

def clean(s: str) -> str:
    """Strip excess whitespace and pipe artifacts."""
    s = re.sub(r'[|\n\r\t]+', ' ', s)
    s = re.sub(r'\s{2,}', ' ', s)
    return s.strip()

def get_context(text: str, start: int, end: int, total_len: int, pad: int = 60) -> str:
    """Get surrounding context, then trim to nearest sentence/phrase boundary."""
    s = max(0, start - pad)
    e = min(total_len, end + pad)
    ctx = clean(text[s:e])

    # Trim to last sentence boundary before the match
    # Sentences end at . ! ? — but not in abbreviations
    before_match = text[s:start]
    last_punct = max(before_match.rfind('. '), before_match.rfind('! '), before_match.rfind('? '))
    if last_punct != -1 and start - s - last_punct < 60:
        ctx = clean(text[last_punct+1+s: e])

    # Trim to first sentence boundary after the match
    after_match = text[end:e]
    first_punct = min(
        after_match.find('. ') if after_match.find('. ') != -1 else 999,
        after_match.find('! ') if after_match.find('! ') != -1 else 999,
        after_match.find('? ') if after_match.find('? ') != -1 else 999,
    )
    if first_punct < 40:
        ctx = clean(text[s:end + first_punct + 2])

    # Hard cap
    if len(ctx) > 200:
        ctx = ctx[:200].rsplit(' ', 1)[0] + '...'

    return ctx

# ─── Patterns ─────────────────────────────────────────────────────────────────

# Quantified: percentages, ratios, team sizes, audience sizes
PCT_RE   = re.compile(r'(\d+(?:\.\d+)?)\s*%')
RATIO_RE = re.compile(r'(\d+(?:\.\d+)?)\s*(?:times|x)\b', re.I)
TEAM_RE  = re.compile(r'(?:team of|managed|supervised|led)\s*(\d+)\b', re.I)
PEOPLE_RE = re.compile(r'(\d+(?:\.\d+)?)\s*(?:people|staff|employees|participants|users|beds|clinics?)\b', re.I)

# Named EPIC/healthcare systems
SYSTEM_PATTERNS = [
    r'\b(EPIC)\b',
    r'\b(NGEMR)\b',
    r'\b(Cadence|Prelude|Bridges|Radiant|Canto|Welcome)\b',
    r'\b(ngemr|national health)\b',
    r'\bSingHealth|NHHS|NHG\b',
    r'\bNUHS\b',
    r'\b(Accenture|Deloitte)\b',
    r'\b(Athena|Cerner|Allscripts)\b',
    r'\b(Power BI|Tableau|SQL|Redcap)\b',
    r'\b(AI project|AI lead|artificial intelligence)\b',
    r'\b(go-live|command centre|UAT|configuration|build)\b',
    r'\b(change management)\b',
    r'\b(digital transformation)\b',
]
SYSTEM_RE = re.compile('|'.join(SYSTEM_PATTERNS), re.I)

# Institutions (named employers)
INSTITUTION_PATTERNS = [
    r'\b(NUHS|National University Health System)\b',
    r'\bSingHealth\b',
    r'\b(NHG|National Healthcare Group)\b',
    r'\bAccenture\b',
    r'\bDeloitte\b',
    r'\bLTA|Land Transport Authority\b',
    r'\bFlying Chalks\b',
    r'\bYishun Hospital|Khoo Teck Puat|IMH\b',
    r'\bAthena Health|Cerner|Allscripts\b',
    r'\bNUHS cluster|NHG cluster\b',
]
INSTITUTION_RE = re.compile('|'.join(INSTITUTION_PATTERNS), re.I)

# Action/achievement verbs
ACHIEVEMENT_VERBS = [
    'led', 'managed', 'co-supervised', 'reduced', 'saved',
    'increased', 'improved', 'streamlined', 'developed', 'designed',
    'implemented', 'executed', 'dispatched', 'supported', 'coordinated',
    'conducted', 'appointed as', 'enhanced', 'resolved', 'built',
    'configured', 'trained', 'presented', 'delivered', 'launched',
    'spearheaded', 'initiated', 'automated', 'digitized', 'transformed',
    'optimised', 'appointed', 'organised', 'secured',
]
ACHIEVEMENT_RE = re.compile(
    r'\b(' + '|'.join(ACHIEVEMENT_VERBS) + r')\b[^.!?]{0,150}',
    re.I
)

# Gap indicators
GAP_KEYWORDS = ['gap', 'break', 'unexplained', 'unaccounted', 'left between']


def extract_claims(candidate_json: Dict[str, Any]) -> List[Dict[str, str]]:
    """
    Extract structured claims from a ResumeScanner candidate JSON.
    Returns list of { id, type, text, source_field }.
    """
    claims  = []
    seen    = set()

    def add(type_, text, source_field):
        key = (type_, clean(text).lower()[:60])
        if key not in seen and len(clean(text)) > 15:
            seen.add(key)
            claims.append({
                "id":           f"claim_{len(claims)+1}",
                "type":         type_,
                "text":         clean(text),
                "source_field": source_field,
            })

    name = candidate_json.get("name", "Unknown")

    # Combine all summary fields for comprehensive scanning
    summaries = [
        candidate_json.get("healthcare_summary", ""),
        candidate_json.get("systems_summary", ""),
        candidate_json.get("communication_summary", ""),
        candidate_json.get("analytical_summary", ""),
    ]
    all_text = ' | '.join(filter(None, summaries))

    # ── 1. Quantified: percentages ─────────────────────────────────────────
    for m in PCT_RE.finditer(all_text):
        ctx = get_context(all_text, m.start(), m.end(), len(all_text), pad=50)
        add(CLAIM_TYPE_QUANTIFIED, ctx, "all_summaries")

    for m in RATIO_RE.finditer(all_text):
        ctx = get_context(all_text, m.start(), m.end(), len(all_text), pad=50)
        add(CLAIM_TYPE_QUANTIFIED, ctx, "all_summaries")

    for m in TEAM_RE.finditer(all_text):
        ctx = get_context(all_text, m.start(), m.end(), len(all_text), pad=50)
        add(CLAIM_TYPE_QUANTIFIED, ctx, "all_summaries")

    for m in PEOPLE_RE.finditer(all_text):
        ctx = get_context(all_text, m.start(), m.end(), len(all_text), pad=50)
        add(CLAIM_TYPE_QUANTIFIED, ctx, "all_summaries")

    # ── 2. Named systems / EPIC modules ───────────────────────────────────
    for m in SYSTEM_RE.finditer(all_text):
        ctx = get_context(all_text, m.start(), m.end(), len(all_text), pad=60)
        add(CLAIM_TYPE_SYSTEM, ctx, "systems_summary")

    # ── 3. Named institutions / employers ─────────────────────────────────
    for m in INSTITUTION_RE.finditer(all_text):
        ctx = get_context(all_text, m.start(), m.end(), len(all_text), pad=60)
        add(CLAIM_TYPE_INSTITUTION, ctx, "all_summaries")

    # ── 4. Achievement-verb sentences ─────────────────────────────────────
    for m in ACHIEVEMENT_RE.finditer(all_text):
        sentence = m.group(0).strip()
        if len(sentence) > 20:
            add(CLAIM_TYPE_LEADERSHIP, sentence, "all_summaries")

    # ── 5. Gap indicators ────────────────────────────────────────────────
    for kw in GAP_KEYWORDS:
        for m in re.finditer(rf'\b{re.escape(kw)}\b', all_text, re.I):
            ctx = get_context(all_text, m.start(), m.end(), len(all_text), pad=50)
            add(CLAIM_TYPE_GAP, ctx, "all_summaries")

    # ── 6. Work history scan ───────────────────────────────────────────────
    work_history = candidate_json.get("work_history", [])
    if isinstance(work_history, list):
        for entry in work_history:
            if not isinstance(entry, dict):
                continue
            for field in ("company", "title", "duration"):
                val = entry.get(field, "")
                if val and len(val) > 8:
                    add(CLAIM_TYPE_LEADERSHIP, val, f"work_history.{field}")

    return claims


def filter_top_claims(claims: List[Dict], max_claims: int = 8) -> List[Dict]:
    """Rank and return top claims by priority type."""
    ranked = sorted(claims, key=lambda c: CLAIM_PRIORITY.get(c["type"], 9))
    return ranked[:max_claims]


def claims_to_prompt_context(claims: List[Dict]) -> str:
    """Format claims for LLM prompt injection."""
    if not claims:
        return ""

    type_labels = {
        CLAIM_TYPE_QUANTIFIED:   "📊 Quantified",
        CLAIM_TYPE_PROJECT:      "🔧 Project",
        CLAIM_TYPE_SYSTEM:      "💻 System",
        CLAIM_TYPE_LEADERSHIP:  "👤 Leadership",
        CLAIM_TYPE_INSTITUTION: "🏢 Institution",
        CLAIM_TYPE_GAP:         "⚠️ Gap",
    }

    lines = []
    for c in claims:
        label = type_labels.get(c["type"], c["type"])
        text  = c["text"]
        lines.append(f"[{label}] {text}")

    return "\n".join(lines)


# ─── CLI test ─────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python3 experience_extractor.py <candidate_json_path>")
        sys.exit(1)

    with open(sys.argv[1]) as f:
        candidate = json.load(f)

    claims = extract_claims(candidate)
    top    = filter_top_claims(claims)

    print(f"=== {candidate.get('name','Unknown')} | {len(claims)} total claims, top {len(top)} ===\n")
    for c in top:
        print(f"[{c['type']:12}] {c['text']}")
        print()

    if top:
        print("=== LLM prompt context ===")
        print(claims_to_prompt_context(top))
