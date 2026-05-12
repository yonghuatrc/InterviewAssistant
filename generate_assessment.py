#!/usr/bin/env python3
"""
generate_assessment.py — InterviewAssistant Assessment Generator

Full pipeline: find candidate → extract questions → generate remarks → fill DOCX → email → Telegram.

Usage:
  python3 generate_assessment.py "<Candidate Name>" "<transcript text>"

Example:
  python3 generate_assessment.py "Daryl Chen" "Interviewer: Can you walk me through..."
"""
import json
import os
import re
import sys
import urllib.request
import urllib.error
import zipfile
import glob
import yaml
from pathlib import Path

# ── Paths ───────────────────────────────────────────────────────────────────────
PROJECT_ROOT  = Path("/mnt/d/Hermes/Project/InterviewAssistant")
RESUME_ROOT   = Path("/mnt/d/Hermes/Project/ResumeScanner")
ARCHIVE_ROOT  = RESUME_ROOT / "archive" / "processed_batches"
REPORTS_ROOT  = RESUME_ROOT / "output" / "reports"
OUTPUT_DIR    = PROJECT_ROOT / "output"
TEMPLATE_DIR  = PROJECT_ROOT / "templates"


# ─────────────────────────────────────────────────────────────────────────────
# 1. find_candidate(name) — locate candidate JSON + score + batch
# ─────────────────────────────────────────────────────────────────────────────

def find_candidate(name: str, require_exact: bool = False):
    """
    Search all ResumeScanner normalised JSONs for a candidate by name.
    Returns the FIRST match (most recent batch if duplicates exist).

    Returns: {
        "json_path":  str,   # absolute path to normalized JSON
        "batch_id":   str,   # e.g. "2026-04-30_090026_batch001"
        "score":      float, # total_score from scores.json
        "candidate_id": str,
        "recommendation": str,
    }
    Raises: FileNotFoundError if no match.
    """
    name_lower = name.lower().strip()
    # Search normalised JSONs
    pattern = str(ARCHIVE_ROOT / "*/processed/normalized/*.json")
    candidates = []

    for json_path in glob.glob(pattern):
        try:
            with open(json_path) as f:
                data = json.load(f)
        except (json.JSONDecodeError, IOError):
            continue

        cand_name = data.get("candidate_name") or data.get("name") or ""
        if name_lower not in cand_name.lower():
            continue

        # Found — now look up score from scores.json
        batch_id = Path(json_path).parts[-4]  # .../batch_id/processed/normalized/id.json
        scores_path = REPORTS_ROOT / batch_id / "scores.json"
        score = None
        recommendation = None

        if scores_path.exists():
            try:
                scores_data = json.load(open(scores_path))
                # scores.json is a list — match by name + source_file (candidate_id may differ across batches)
                norm_name = data.get("candidate_name") or data.get("name") or ""
                norm_source = data.get("source_file") or ""
                for entry in scores_data:
                    entry_name = entry.get("name", "")
                    entry_source = entry.get("source_file", "")
                    if (norm_name.lower() in entry_name.lower() or
                        entry_name.lower() in norm_name.lower()) and \
                       norm_source == entry_source:
                        score = entry.get("total_score")
                        recommendation = entry.get("recommendation")
                        break
            except (json.JSONDecodeError, IOError):
                pass

        candidates.append({
            "json_path":     json_path,
            "batch_id":      batch_id,
            "score":         score,
            "candidate_id":  data.get("candidate_id", ""),
            "recommendation": recommendation,
            "cand_name":     cand_name,
        })

    if not candidates:
        raise FileNotFoundError(f"Candidate not found: {name}")

    # If exact match required and we have ambiguous matches, filter further
    if require_exact:
        exact = [c for c in candidates if c["cand_name"].lower() == name_lower]
        if exact:
            candidates = exact

    # Prefer most recent batch (lexicographically descending by date+time string)
    candidates.sort(key=lambda c: c["batch_id"], reverse=True)
    best = candidates[0]
    return {
        "json_path":     best["json_path"],
        "batch_id":      best["batch_id"],
        "score":         best["score"],
        "candidate_id":  best["candidate_id"],
        "recommendation": best["recommendation"],
    }


# ─────────────────────────────────────────────────────────────────────────────
# 2. find_prep_doc(candidate_name) — locate InterviewGenerator prep DOCX
# ─────────────────────────────────────────────────────────────────────────────

def find_prep_doc(candidate_name: str):
    """
    Locate the InterviewGenerator prep DOCX for a candidate.
    Pattern: output/{Name}_Interview_Prep.docx
    Spaces in name are converted to underscores.

    Returns: absolute path string
    Raises: FileNotFoundError if not found.
    """
    safe_name = candidate_name.replace(" ", "_")
    doc_path = OUTPUT_DIR / f"{safe_name}_Interview_Prep.docx"
    if doc_path.exists():
        return str(doc_path)

    # Fallback: case-insensitive scan of output dir
    for f in OUTPUT_DIR.glob("*_Interview_Prep.docx"):
        if candidate_name.lower() in f.stem.lower().replace("_", " "):
            return str(f)

    raise FileNotFoundError(
        f"Prep DOCX not found for: {candidate_name}\n"
        f"Looked in: {OUTPUT_DIR}"
    )


# ─────────────────────────────────────────────────────────────────────────────
# 3. extract_interview_questions(docx_path) — pull question text from prep DOCX
# ─────────────────────────────────────────────────────────────────────────────

def extract_interview_questions(docx_path: str):
    """
    Read an InterviewGenerator prep DOCX and return a list of question strings.

    Returns: [
        {"section": "Section 1: Accountability", "q_num": 1,
         "text": "...", "whatToListen": [...]},
        ...
    ]
    Each entry has: section name, question number within that section,
    the full question text, and the good/red flag items.

    Returns at least 6 entries (3 sections × 2 questions) for a standard prep doc.
    """
    result = []

    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml").decode("utf-8", errors="replace")

    # Strip XML tags, normalise whitespace
    text = re.sub(r"<[^>]+>", " ", xml)
    text = re.sub(r"\s+", " ", text).strip()

    # Split on Q-numbered blocks — each question starts with Q1/Q2/etc. in the doc
    # The doc format is: "Q1: ...?" or "Q1 — ..." or "Q1." followed by question text
    question_pattern = re.compile(
        r"(Q\d+)[:\s\—]+(.*?)(?=Q\d+[:\s\—]|\Z)",
        re.DOTALL | re.IGNORECASE
    )

    for m in question_pattern.finditer(text):
        q_marker = m.group(1)          # "Q1" etc.
        q_body   = m.group(2).strip()   # everything after Q1: up to next Q

        if not q_body or len(q_body) < 15:
            continue

        # Extract the actual question sentence (ends at first ? or full stop that follows a question word)
        # Normalise HTML entities
        q_body = q_body.replace("&apos;", "'").replace("&amp;", "&").replace("&quot;", '"')

        # Try to isolate the question — find first '?' or use first sentence
        if "?" in q_body:
            end_idx = q_body.index("?")
            question_text = q_body[:end_idx + 1].strip()
        else:
            # Cut at first period after a reasonable length
            sentences = re.split(r"(?<=[.!?])\s+", q_body)
            question_text = sentences[0].strip() if sentences else q_body[:200]

        if len(question_text) < 20:
            continue

        # Determine section by looking at preceding text in the original XML
        # We also extract the "Good signals" / "Red flags" if present
        whatToListen = _extract_signals(q_body)

        q_num = int(re.search(r"\d+", q_marker).group())

        result.append({
            "q_num":       q_num,
            "text":        question_text,
            "whatToListen": whatToListen,
        })

    # Deduplicate and sort by q_num
    seen = set()
    deduped = []
    for item in result:
        key = (item["q_num"], item["text"][:60])
        if key not in seen:
            seen.add(key)
            deduped.append(item)

    deduped.sort(key=lambda x: x["q_num"])
    return deduped


def _extract_signals(text: str) -> list:
    """
    Extract ✅ good-signal and 🚩 red-flag items from a question body block.
    Returns list of {"check": True/False, "text": "..."}
    """
    signals = []
    # Good signals: look for "✅" or "Good signals:" or "Good signal:" patterns
    good_block = re.search(
        r"(?:Good signals?|✅|Positive)[:\s]*(.*?)(?:Probes?|Red flags?|🚩|$)",
        text, re.DOTALL | re.IGNORECASE
    )
    if good_block:
        items = re.findall(r"(?:^|[•\n])[\s]*([^•\n]+?)(?:(?=\n[^•])|$)", good_block.group(1))
        for item in items:
            item = item.strip()
            if item and len(item) > 5:
                signals.append({"check": True, "text": item})

    red_block = re.search(
        r"(?:Red flags?|🚩|Red flags?)[:\s]*(.*?)(?:Probes?|$)",
        text, re.DOTALL | re.IGNORECASE
    )
    if red_block:
        items = re.findall(r"(?:^|[•\n])[\s]*([^•\n]+?)(?:(?=\n[^•])|$)", red_block.group(1))
        for item in items:
            item = item.strip()
            if item and len(item) > 5:
                signals.append({"check": False, "text": item})

    return signals


# ─────────────────────────────────────────────────────────────────────────────
# 4. load_hr_template() — read current_version.txt → return template path
# ─────────────────────────────────────────────────────────────────────────────

def load_hr_template():
    """
    Read templates/current_version.txt to get the active template version,
    then return the path to that version's DOCX.

    Version file format: "v1" (no extension)
    Expected template: templates/Interview Assessment Form_GMIO.docx

    Returns: absolute path string
    Raises: FileNotFoundError if version file or template is missing.
    """
    version_file = TEMPLATE_DIR / "current_version.txt"
    if not version_file.exists():
        raise FileNotFoundError(f"Version file not found: {version_file}")

    version = version_file.read_text().strip()
    if not version:
        raise FileNotFoundError(f"Version file is empty: {version_file}")

    # Template filename pattern: the base template is the GMIO form
    # When HR updates the form, they add version suffix e.g. "HR_ASSESSMENT_v2.docx"
    # For now the base template IS the current version — check for versioned override first
    versioned_template = TEMPLATE_DIR / f"Interview Assessment Form_GMIO_{version}.docx"
    if versioned_template.exists():
        return str(versioned_template)

    base_template = TEMPLATE_DIR / "Interview Assessment Form_GMIO.docx"
    if base_template.exists():
        return str(base_template)

    raise FileNotFoundError(
        f"Template not found for version '{version}'.\n"
        f"Checked: {versioned_template} and {base_template}"
    )


# ─────────────────────────────────────────────────────────────────────────────
# LLM — Assessment content generation (provider-agnostic via config.yaml)
# ─────────────────────────────────────────────────────────────────────────────

def _load_llm_config():
    """Read current default provider/model/base_url from config.yaml.

    Also ensures ~/.hermes/.env is loaded so env vars (MINIMAX_BASE_URL,
    MINIMAX_API_KEY etc.) are available — this makes the function work
    correctly both when called from __main__ (where .env is pre-loaded)
    and when the module is imported by other tools.
    """
    import yaml
    # Load .env if not already present (handles import-time vs __main__ differences)
    if not os.environ.get("MINIMAX_BASE_URL"):
        env_path = os.path.expanduser("~/.hermes/.env")
        if os.path.exists(env_path):
            for line in open(env_path):
                line = line.strip()
                if "=" in line and not line.startswith("#"):
                    k, v = line.split("=", 1)
                    os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))

    cfg = yaml.safe_load(open(os.path.expanduser("~/.hermes/config.yaml")))
    model_cfg  = cfg.get("model", {})
    provider   = model_cfg.get("provider", "minimax")
    model      = model_cfg.get("default", "MiniMax-M2.7")
    providers  = cfg.get("providers", {})
    pcfg       = providers.get(provider, {})
    # Use MINIMAX_BASE_URL env var if set, otherwise fall back to config.
    # NOTE: config.yaml currently stores "anthropic" path for minimax, but the
    # working endpoint is /v1 — prefer env var override.
    base_url = os.environ.get("MINIMAX_BASE_URL", "").strip()
    if not base_url:
        base_url = pcfg.get("base_url") or "https://api.minimax.io/v1"
    api_key = os.environ.get("MINIMAX_API_KEY", pcfg.get("api_key", ""))
    return provider, model, base_url, api_key


def _truncate_transcript(transcript: str, max_chars: int = 8000) -> str:
    """Truncate transcript to max_chars, appending a truncation notice if cut.

    Args:
        transcript: Raw transcript text.
        max_chars:  Maximum character length before truncation (default 8000).

    Returns:
        Truncated transcript string with notice appended if truncated,
        or the original string unchanged if within limit.
    """
    if not transcript or len(transcript) <= max_chars:
        return transcript
    truncation_note = (
        f"\n\n[TRANSCRIPT TRUNCATED] "
        f"(remaining {len(transcript) - max_chars} chars omitted)"
    )
    return transcript[:max_chars] + truncation_note


def generate_assessment_content(transcript: str, questions: list, candidate_context: dict) -> dict:
    """
    Call the default LLM to generate a free-text remarks field from an interview transcript.

    Args:
        transcript:         Full interview transcript (string, from Boss pasted in chat)
        questions:          List of question dicts from extract_interview_questions()
        candidate_context:  Dict with keys: name, score, recommendation, education,
                            work_history (from ResumeScanner JSON)

    Returns:
        {"remarks": "150-200 word narrative assessment..."}
    Raises:
        RuntimeError if LLM call fails.
    """
    provider, model, base_url, api_key = _load_llm_config()

    # Build questions list for the prompt
    q_lines = []
    for q in questions:
        q_lines.append(f"  Q{q.get('q_num', '?')}: {q.get('text', '')}")
    questions_text = "\n".join(q_lines) if q_lines else "Not available"

    # Build candidate context summary
    name    = candidate_context.get("name", "the candidate")
    score   = candidate_context.get("score", "N/A")
    rec     = candidate_context.get("recommendation", "N/A")
    edu     = candidate_context.get("education", [])
    work    = candidate_context.get("work_history", [])
    edu_txt  = "; ".join(str(e) for e in edu) if edu else "Not available"
    work_txt = "; ".join(str(w) for w in work) if work else "Not available"

    prompt = f"""You are an expert HR assessor for NUHS (National University Health System), Singapore.
Your task is to write a professional assessment remark for a job interview assessment form.

## Candidate
Name: {name}
ResumeScanner Score: {score} / 100
ResumeScanner Recommendation: {rec}
Education: {edu_txt}
Work History: {work_txt}

## Interview Questions That Were Asked
(These are the questions prepared for this candidate before the interview)
{questions_text}

## Interview Transcript
§TRANSCRIPT§

## Your Task
Based on the interview transcript above, write a professional assessment remark in the following style used by NUHS GMIO HR assessors. The remark must:
1. Be between 150 and 200 words.
2. Be written as a single coherent paragraph (or two short paragraphs).
3. Cover the candidate's strengths demonstrated in the interview.
4. Cover areas for development or concerns observed.
5. Provide an overall impression of the candidate's suitability for the role.
6. Be grounded SPECIFICALLY in what the candidate said — quote or paraphrase actual transcript content to support your assessment. Do NOT write generic statements.
7. Use professional HR assessment language — formal but specific, not flowery.
8. Do NOT include numerical ratings or category scores.
9. Do NOT reference the ResumeScanner score or recommendation — only what was said in the interview.

Return your response as a valid JSON object with a single key "remarks":
{{"remarks": "your 150-200 word assessment paragraph here"}}

Output only valid JSON. No markdown. No explanation."""

    # Inject transcript via string replacement (avoid .format() which misinterprets JSON in prompt)
    transcript = _truncate_transcript(transcript)
    prompt = prompt.replace("§TRANSCRIPT§", transcript)

    payload = json.dumps({
        "model": model,
        "messages": [
            {"role": "system", "content": "You are a strict NUHS HR assessor. Output ONLY a valid JSON object with key 'remarks'. No markdown. No explanation."},
            {"role": "user",   "content": prompt},
        ],
        "temperature": 0.3,
        "max_tokens": 4096,
    }).encode("utf-8")

    req = urllib.request.Request(
        f"{base_url}/chat/completions",
        data=payload,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type":  "application/json",
        },
        method="POST",
    )

    content = None
    try:
        with urllib.request.urlopen(req, timeout=max(120, len(prompt) // 10)) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            content = data["choices"][0]["message"]["content"].strip()
            if not content:
                raise RuntimeError(
                    f"LLM returned empty response for transcript ({len(transcript)} chars). "
                    f"Transcript may be too long for model context window."
                )
    except Exception as e:
        raise RuntimeError(f"LLM API error ({provider}/{model}): {e}") from e

    # Strip thinking blocks and markdown
    content = re.sub(r'<think[^>]*>[\s\S]*?<\/think>', '', content, flags=re.IGNORECASE).strip()
    content = re.sub(r'^```(?:json)?\s*', '', content, flags=re.I).strip()
    content = re.sub(r'\s*```$', '', content).strip()

    try:
        result = json.loads(content)
        remarks = result.get("remarks", "")
        word_count = len(remarks.split())
        if word_count < 150:
            raise ValueError(f"Remarks too short ({word_count} words, expected 150-200)")
        return {"remarks": remarks, "_word_count": word_count, "_provider": provider, "_model": model}
    except (json.JSONDecodeError, ValueError) as e:
        raise RuntimeError(f"LLM output parse error: {e}\nRaw: {content[:500]}") from e


# ─────────────────────────────────────────────────────────────────────────────
# Phase 4 — DOCX Generation (fill HR template)
# ─────────────────────────────────────────────────────────────────────────────

def fill_hr_template(template_path: str, output_path: str, assessment_json: dict) -> str:
    """
    Fill the HR assessment DOCX template with assessment content.

    Approach: XML text substitution on the existing template DOCX.
    Specifically:
      - Candidate name + position: text replacement in the name/position cells
      - Remarks: paragraph replacement in the Remarks text area

    Args:
        template_path:    Path to the template DOCX (e.g. Interview Assessment Form_GMIO.docx)
        output_path:      Where to write the filled DOCX
        assessment_json:  Dict with keys: candidate_name, position, remarks, [optional: signature, date]

    Returns:
        Path to the output DOCX file.
    Raises:
        RuntimeError on failure.
    """
    import shutil
    candidate_name = assessment_json.get("candidate_name", "")
    position       = assessment_json.get("position", "Medical Informatics Specialist")
    remarks        = assessment_json.get("remarks", "")
    sign_by        = assessment_json.get("sign_by", "Dennis Ng")
    sign_date      = assessment_json.get("date", "")

    if not remarks:
        raise ValueError("assessment_json must contain non-empty 'remarks' field")

    # Copy template to output path first
    shutil.copy(template_path, output_path)

    # Read all zip contents
    with zipfile.ZipFile(output_path, "r") as zin:
        names    = zin.namelist()
        contents = {n: zin.read(n) for n in names}

    xml = contents["word/document.xml"].decode("utf-8")

    # ── Fill 1 & 2: Name and Position (new template has empty value cells) ──
    # Each label is followed by a gridSpan=20 value cell — inject text there.
    xml = _inject_after_label(xml, "Name of Candidate:", candidate_name.upper())
    xml = _inject_after_label(xml, "Position Applied:", position)

    # ── Fill 3: Remarks text — via <<Insert here>> placeholder ──────────────
    xml = _replace_insert_here(xml, remarks)

    # ── Fill 4: Sign-by name and date ───────────────────────────────────────
    # "Remarks By: Dennis Ng" is in one combined cell — replace just the name.
    # "Date:" label has an empty value cell after it — inject date there.
    xml = _replace_cell_text(xml, "Dennis Ng", sign_by)
    xml = _inject_after_label(xml, "Date:", sign_date)

    # Write back
    contents["word/document.xml"] = xml.encode("utf-8")

    # Fix tblGrid (same post-processor as generate_interview.py)
    xml_fixed = _fix_tblgrid(contents["word/document.xml"].decode("utf-8"))
    contents["word/document.xml"] = xml_fixed.encode("utf-8")

    # Rebuild zip
    os.remove(output_path)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, contents[name])

    return output_path


def _replace_cell_text(xml: str, old: str, new: str) -> str:
    """
    Replace text content inside the first matching cell in the document XML.
    Handles the cell content wrapped in <w:t> elements.
    Preserves run properties (rPr) of the containing run.
    """
    # Escape for XML
    old_esc = old.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    new_esc = new.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    # Find the <w:t> element containing the old text
    # Pattern: <w:t ...>OLD_TEXT</w:t>
    pattern = re.compile(
        r'(<w:t(?:[^>]*)>)([^<]*' + re.escape(old_esc) + r'[^<]*)(</w:t>)',
        re.IGNORECASE
    )
    if pattern.search(xml):
        xml = pattern.sub(lambda m: m.group(1) + m.group(2).replace(old_esc, new_esc) + m.group(3), xml, count=1)
    return xml


def _inject_after_label(xml: str, label: str, new_text: str) -> str:
    """
    Find a label cell by its text, then inject new_text into the
    immediately following sibling cell (same <w:tr> row).

    Handles the case where the value cell is empty (gridSpan=20, no text).
    """
    safe_text = new_text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    idx = xml.find(label)
    if idx == -1:
        raise RuntimeError(f"Label '{label}' not found in template XML")

    # Find the <w:tc> containing this label
    tc_start = xml.rfind("<w:tc>", 0, idx)
    tc_end = xml.find("</w:tc>", idx) + len("</w:tc>")

    # Find the next <w:tc> in the same row
    next_tc_start = xml.find("<w:tc>", tc_end)
    next_tc_end = xml.find("</w:tc>", next_tc_start) + len("</w:tc>")
    next_cell_xml = xml[next_tc_start:next_tc_end]

    # Build the replacement run (matching Arial 20pt style)
    rPr = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
    new_runs = f'<w:r>{rPr}<w:t xml:space="preserve">{safe_text}</w:t></w:r>'

    # The value cell has an empty <w:p> — insert runs before </w:p>
    if "</w:p>" in next_cell_xml:
        new_cell = next_cell_xml.replace("</w:p>", f"{new_runs}</w:p>")
    else:
        new_cell = next_cell_xml + new_runs

    xml = xml[:next_tc_start] + new_cell + xml[next_tc_end:]
    return xml


def _replace_remarks(xml: str, new_remarks: str) -> str:
    """
    Inject remarks text into the first empty gridSpan=19 content cell that follows
    the "share your assessment" label.

    Uses the same approach as _inject_after_label — finds the target cell by
    gridSpan=19, then injects the text into the existing empty <w:p> element,
    preserving all cell properties (gridSpan, borders, etc.).

    Template structure:
      - Row 1: [Label cell: gridSpan=19] | [Strong Candidate: gridSpan=3]
      - Row 2: [Remarks CONTENT cell: gridSpan=19 — EMPTY] | [Good Candidate: gridSpan=3]
    """
    safe_remarks = new_remarks.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    marker = "share your assessment "
    idx = xml.find(marker)
    if idx == -1:
        raise RuntimeError(f"Remarks marker '{marker}' not found in template XML")

    label_tc_end = xml.find("</w:tc>", idx)
    if label_tc_end == -1:
        raise RuntimeError("Could not find end of remarks label cell")

    # Find the first gridSpan=19 cell after the label cell
    search_start = label_tc_end
    while True:
        tc_start = xml.find("<w:tc>", search_start)
        if tc_start == -1:
            raise RuntimeError("Could not find remarks content cell (gridSpan=19)")
        tc_end = xml.find("</w:tc>", tc_start) + len("</w:tc>")
        cell_xml = xml[tc_start:tc_end]

        gridspan = re.search(r'<w:gridSpan w:val="(\d+)"/>', cell_xml)
        if gridspan and gridspan.group(1) == "19":
            # Found the remarks content cell
            break

        search_start = tc_end

    # Build the new text run (Arial 20pt, matching form style)
    rPr = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'

    # Split remarks into paragraphs at sentence boundaries
    para_texts = re.split(r'(?<=[.!?])\s+(?=[A-Z])', safe_remarks)
    para_texts = [p.strip() for p in para_texts if p.strip()]

    # Build <w:p> elements — one per paragraph
    new_paras = []
    for para_text in para_texts:
        text_run = f'<w:t xml:space="preserve">{para_text}</w:t>'
        new_paras.append(f'<w:p><w:r>{rPr}{text_run}</w:r></w:p>')

    if not new_paras:
        new_paras = [f'<w:p><w:r>{rPr}<w:t xml:space="preserve">{safe_remarks}</w:t></w:r></w:p>']

    new_runs = "".join(new_paras)

    # Inject into the existing empty <w:p> (before </w:p>) to preserve cell structure
    # CRITICAL: replace the LAST </w:p> in the cell, not the first.
    # Some template cells have multiple <w:p> open tags but only one </w:p> close tag
    # (e.g. a cell with a formcheckbox + empty paragraph has 2 opens, 1 close).
    # Replacing the last close avoids corrupting formcheckbox field codes.
    if "</w:p>" in cell_xml:
        # Find the last </w:p> position
        last_close = cell_xml.rfind("</w:p>")
        new_cell = cell_xml[:last_close] + new_runs + cell_xml[last_close:]
    else:
        new_cell = cell_xml + new_runs

    xml = xml[:tc_start] + new_cell + xml[tc_end:]

    if new_remarks[:20] not in xml:
        raise RuntimeError("Remarks text was not injected — XML substitution may have failed")

    return xml


# Import fix_tblgrid from generate_interview (same implementation, tested)
sys.path.insert(0, os.path.dirname(__file__))
try:
    from generate_interview import fix_tblgrid
except ImportError:
    def fix_tblgrid(path): pass  # no-op fallback

def _replace_insert_here(xml: str, assessment_text: str) -> str:
    """
    Replace the paragraph containing '<<Insert here>>' with the assessment text.

    The placeholder paragraph has 3 runs: << / Insert here / >>
    This function replaces the entire paragraph with a single run containing the
    assessment text in Arial 10pt (matching surrounding body text, no bold).
    """
    # Find "Insert here" in the XML
    idx = xml.find('>Insert here<')
    if idx == -1:
        raise RuntimeError("'<<Insert here>>' placeholder not found in template XML")

    # Find the paragraph boundaries
    para_start = xml.rfind('<w:p ', 0, idx)
    if para_start == -1:
        para_start = xml.rfind('<w:p>', 0, idx)
    para_end = xml.find('</w:p>', idx) + 6
    para_xml = xml[para_start:para_end]

    # Build replacement run properties (Arial 10pt, no bold — body text style)
    rPr = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    rPr += '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'

    safe_text = (assessment_text
                 .replace('&', '&amp;')
                 .replace('<', '&lt;')
                 .replace('>', '&gt;'))
    new_run = f'<w:r>{rPr}<w:t xml:space="preserve">{safe_text}</w:t></w:r>'

    # Extract paragraph attributes from original to preserve them
    para_id    = re.search(r'w14:paraId="([^"]+)"', para_xml)
    text_id    = re.search(r'w14:textId="([^"]+)"', para_xml)
    rsid_r     = re.search(r'w:rsidR="([^"]+)"', para_xml)
    rsid_rpr   = re.search(r'w:rsidRPr="([^"]+)"', para_xml)
    rsid_p     = re.search(r'w:rsidP="([^"]+)"', para_xml)
    rsid_rdef  = re.search(r'w:rsidRDefault="([^"]+)"', para_xml)

    # Build replacement paragraph tag
    new_attrs = '<w:p'
    for attr, match in [('w14:paraId', para_id), ('w14:textId', text_id),
                        ('w:rsidR', rsid_r), ('w:rsidRPr', rsid_rpr),
                        ('w:rsidP', rsid_p), ('w:rsidRDefault', rsid_rdef)]:
        if match:
            new_attrs += f' {attr}="{match.group(1)}"'
    new_attrs += '>'

    # Use paragraph properties from original (remove bold from rPr inside pPr)
    ppr_match = re.search(r'<w:pPr>.*?</w:pPr>', para_xml, re.DOTALL)
    ppr = ppr_match.group(0) if ppr_match else ''
    # Strip <w:b/> from paragraph rPr (we want normal weight)
    ppr = re.sub(r'<w:b/>', '', ppr)
    ppr = re.sub(r'<w:b\s*/>', '', ppr)
    ppr = re.sub(r'<w:bCs/>', '', ppr)
    ppr = re.sub(r'<w:bCs\s*/>', '', ppr)

    replacement = f'{new_attrs}{ppr}{new_run}</w:p>'

    new_xml = xml[:para_start] + replacement + xml[para_end:]

    # Verify
    if assessment_text[:40] not in new_xml:
        raise RuntimeError("Assessment text was not injected — XML substitution failed")

    return new_xml


def _fix_tblgrid(xml_str: str) -> str:
    """Standalone version — not used when generate_interview fix_tblgrid is available."""
    return xml_str  # passthrough


# ─────────────────────────────────────────────────────────────────────────────
# Main pipeline
# ─────────────────────────────────────────────────────────────────────────────

def main(candidate_name: str, transcript: str) -> str:
    """
    Run the full assessment pipeline for a candidate.

    Steps:
      1. Find candidate JSON + scores
      2. Find InterviewGenerator prep DOCX
      3. Extract interview questions from prep DOCX
      4. Load HR template
      5. Call LLM to generate remarks from transcript
      6. Fill HR template → output DOCX
      7. Email DOCX to Boss

    Returns:
        Path to the generated DOCX file.
    """
    from datetime import date
    import shutil
    from generate_interview import fix_tblgrid

    ASSESSMENT_OUTPUT = PROJECT_ROOT / "assessment_output"
    ASSESSMENT_OUTPUT.mkdir(exist_ok=True)

    safe_name  = candidate_name.replace(" ", "_")
    today_str = date.today().strftime("%d/%m/%Y")
    output_file = ASSESSMENT_OUTPUT / f"{safe_name}_Assessment.docx"

    # ── Step 1: Find candidate ─────────────────────────────────────────────
    print(f"[1] Finding candidate: {candidate_name}")
    cand = find_candidate(candidate_name)
    print(f"    batch={cand['batch_id']}, score={cand['score']}, id={cand['candidate_id']}")

    # ── Step 2: Load candidate JSON ───────────────────────────────────────
    print(f"[2] Loading candidate JSON")
    with open(cand["json_path"]) as f:
        cand_json = json.load(f)

    # ── Step 3: Find prep DOCX ──────────────────────────────────────────
    print(f"[3] Finding prep DOCX")
    prep_doc = find_prep_doc(candidate_name)
    print(f"    {prep_doc}")

    # ── Step 4: Extract interview questions ───────────────────────────────
    print(f"[4] Extracting interview questions")
    questions = extract_interview_questions(prep_doc)
    print(f"    {len(questions)} questions extracted")

    # ── Step 5: Load HR template ─────────────────────────────────────────
    print(f"[5] Loading HR template")
    template_path = load_hr_template()
    print(f"    {template_path}")

    # ── Step 6: Build context for LLM ───────────────────────────────────
    candidate_context = {
        "name":           cand_json.get("candidate_name", candidate_name),
        "score":          cand.get("score"),
        "recommendation": cand.get("recommendation"),
        "education":      cand_json.get("education", []),
        "work_history":   cand_json.get("work_history", []),
        "position":       cand_json.get("applied_position", "Medical Informatics Specialist"),
    }

    # ── Step 7: Generate remarks via LLM ──────────────────────────────────
    print(f"[6] Generating assessment remarks via LLM...")
    result = generate_assessment_content(transcript, questions, candidate_context)
    remarks  = result["remarks"]
    wc       = result["_word_count"]
    provider = result["_provider"]
    model    = result["_model"]
    print(f"    remarks: {wc} words ({provider}/{model})")

    # ── Step 8: Fill HR template ─────────────────────────────────────────
    print(f"[7] Filling HR template")
    assessment_json = {
        "candidate_name": candidate_context["name"].upper(),
        "position":       candidate_context["position"],
        "remarks":       remarks,
        "sign_by":       "Dennis Ng",
        "date":          today_str,
    }
    tmp_output = str(output_file) + ".tmp"
    fill_hr_template(template_path, tmp_output, assessment_json)

    # Apply tblGrid fix (generates fresh correct ZIP at tmp_output)
    fix_tblgrid(tmp_output)

    # Move to final output path
    shutil.move(tmp_output, str(output_file))
    print(f"    {output_file}")

    # ── Step 9: Send email ───────────────────────────────────────────────
    print(f"[8] Sending email...")
    try:
        sys.path.insert(0, str(PROJECT_ROOT))
        from send_email import send as send_email
        send_email(
            candidate_name  = candidate_context["name"],
            docx_path      = str(output_file),
            batch_id       = cand["batch_id"],
            score          = str(cand["score"] or "N/A"),
            additional_info = (
                f"Assessment generated by Hermes InterviewAssistant.\n"
                f"LLM: {provider}/{model} ({wc} words in remarks)"
            ),
        )
    except Exception as e:
        print(f"    WARNING: Email failed ({e}) — DOCX still saved to {output_file}")

    print(f"\n✓ DONE — {output_file}")
    return str(output_file)


# ─────────────────────────────────────────────────────────────────────────────
# CLI entry point
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="InterviewAssistant — generate HR assessment DOCX")
    parser.add_argument("candidate_name", nargs="?", help="Candidate name (as appears in ResumeScanner)")
    parser.add_argument("transcript",    nargs="?", help="Full interview transcript text")
    parser.add_argument("--test", action="store_true", help="Run Phase 2 unit tests only")
    args = parser.parse_args()

    if args.test:
        import pprint
        print("=" * 60)
        print("Phase 2 — Context Gathering Tests")
        print("=" * 60)
        test_name = "Daryl Chen"
        # [Phase 2 tests — run inline]
        print(f"\n[1] find_candidate('{test_name}')")
        try:
            result = find_candidate(test_name)
            pprint.pprint(result)
            assert result["batch_id"] == "2026-04-30_090026_batch001"
            assert result["score"] == 79.0
            print("  ✓ PASS")
        except Exception as e:
            print(f"  ✗ FAIL: {e}")

        print(f"\n[2] find_prep_doc('{test_name}')")
        try:
            path = find_prep_doc(test_name)
            assert "Daryl_Chen_Interview_Prep.docx" in path
            print(f"  ✓ {path}")
        except Exception as e:
            print(f"  ✗ FAIL: {e}")

        print(f"\n[3] extract_interview_questions")
        try:
            prep = find_prep_doc(test_name)
            qs = extract_interview_questions(prep)
            print(f"  ✓ {len(qs)} questions")
        except Exception as e:
            print(f"  ✗ FAIL: {e}")

        print(f"\n[4] load_hr_template()")
        try:
            tpl = load_hr_template()
            print(f"  ✓ {tpl}")
        except Exception as e:
            print(f"  ✗ FAIL: {e}")

        print("\n" + "=" * 60)
        print("Phase 2 tests complete.")
        print("=" * 60)
    else:
        # If transcript arg is missing, read from stdin (allows piping large transcript text)
        if args.transcript is None:
            args.transcript = sys.stdin.read()
        output = main(args.candidate_name, args.transcript)
        print(f"\nOUTPUT:{output}")