#!/usr/bin/env python3
"""
InterviewGenerator — Phase 2 Orchestrator
Reads candidate JSON, extracts claims, generates personalized questions via LLM,
builds DOCX with both generic + personalized questions, emails to Boss.

Usage: python3 generate_interview.py <candidate_name> <json_path> <batch_id> <score> [role]
 role is optional: pass "PT" to include PT-specific training questions in Section 2
"""
import json
import os
import subprocess
import smtplib
import sys
import tempfile
import re
import zipfile
import io
import urllib.request
import urllib.error
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# ── Load .env ─────────────────────────────────────────────────────────────────
ENV_PATH = os.path.expanduser("~/.hermes/.env")
if os.path.exists(ENV_PATH):
    for line in open(ENV_PATH):
        line = line.strip()
        if "=" in line and not line.startswith("#"):
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))

# ── Config ─────────────────────────────────────────────────────────────────────
GMAIL_USER        = "yonghuatrc@gmail.com"
GMAIL_APP_PASSWORD= os.environ.get("SMTP_PASSWORD", "")
TO_EMAIL          = "dennis_ng@nuhs.edu.sg"
OUTPUT_DIR        = "/mnt/d/Hermes/Project/InterviewGenerator/output"
MINIMAX_API_KEY   = os.environ.get("MINIMAX_API_KEY", "")
MINIMAX_BASE_URL = os.environ.get("MINIMAX_BASE_URL", "https://api.minimax.io/v1")

# ── Total experience calculator ──────────────────────────────────────────────────
from datetime import datetime

MONTH_MAP = {
    "january": 1, "jan": 1,
    "february": 2, "feb": 2,
    "march": 3, "mar": 3,
    "april": 4, "apr": 4,
    "may": 5,
    "june": 6, "jun": 6,
    "july": 7, "jul": 7,
    "august": 8, "aug": 8,
    "september": 9, "sep": 9,
    "october": 10, "oct": 10,
    "november": 11, "nov": 11,
    "december": 12, "dec": 12,
}


def _parse_date(token: str):
    """Parse 'June 2021', '(Dec 2025)' or 'Present' into a datetime object."""
    token = token.strip().strip("()").strip()
    if token.lower() == "present":
        return datetime.now()
    parts = token.split()
    if len(parts) == 2 and parts[0].lower() in MONTH_MAP:
        try:
            return datetime(int(parts[1]), MONTH_MAP[parts[0].lower()], 1)
        except (ValueError, IndexError):
            return None
    return None


def compute_total_experience(work_experience: list) -> float:
    """Parse work_experience date ranges and compute total overlapping years.
    Entries follow format: '<role>, <company>, <month> <year> - <month> <year>'
    """
    ranges = []
    for entry in work_experience:
        if " - " not in entry:
            continue
        left, right = entry.split(" - ", 1)
        # Start date: last two words of left part (e.g. 'June 2021')
        left_words = left.strip().split()
        if len(left_words) < 2:
            continue
        start = _parse_date(" ".join(left_words[-2:]))
        # End date: right part is 'Month Year' or 'Present'
        end = _parse_date(right.strip())
        if start and end and end > start:
            ranges.append((start, end))
    ranges.sort(key=lambda r: r[0])
    merged = []
    for start, end in ranges:
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    total_days = sum((end - start).days for start, end in merged)
    return round(total_days / 365.25, 1)


# ── Import Phase 2 claim extractor ──────────────────────────────────────────────
sys.path.insert(0, os.path.dirname(__file__))
try:
    from experience_extractor import extract_claims, filter_top_claims, claims_to_prompt_context
except ImportError:
    # Fallback if module import fails
    def extract_claims(c):   return []
    def filter_top_claims(l, n=8): return l
    def claims_to_prompt_context(l): return ""

# ─────────────────────────────────────────────────────────────────────────────
# MiniMax LLM — personalized question generation
# ─────────────────────────────────────────────────────────────────────────────

GENERIC_MOTIVATION = [
    {
        "question": "Walk me through your resume — what is it about this specific role at NUHS that drew you to apply?",
        "whatToListen": [
            {"check": True,  "text": "References specific NUHS context, NGEMR, or healthcare mission"},
            {"check": True,  "text": "Articulates a clear 'why now' — something in their career prompted this move"},
            {"check": True,  "text": "Shows research — mentions specific programmes, systems, or initiatives"},
            {"cross": True, "text": "'I just saw a job ad and applied' — no specific motivation"},
            {"cross": True, "text": "Generic 'I want to help patients' with no tie to their skills"},
        ],
        "probes": [],
    },
    {
        "question": "Where do you see yourself in 3 years — and does this role sit on that path?",
        "whatToListen": [
            {"check": True,  "text": "Realistic career trajectory — not over-claiming"},
            {"check": True,  "text": "Specific skills or growth areas named that this role would develop"},
            {"check": True,  "text": "Shows self-awareness — knows what they don't know yet"},
            {"cross": True, "text": "Vague — 'I want to grow' without specifics"},
            {"cross": True, "text": "The role doesn't connect to their stated path — mismatch"},
        ],
        "probes": [],
    },
]

# PT-specific training questions — appended to Section 2 for Principal Trainer candidates
PT_TRAINING_QUESTIONS = [
    {
        "question": "Walk me through how you design a training session from scratch — from needs analysis to post-session evaluation. What tools or frameworks do you use?",
        "whatToListen": [
            {"check": True,  "text": "Starts with needs analysis — identifies who needs training and why"},
            {"check": True,  "text": "Structures content logically: objectives → delivery → practice → assessment"},
            {"check": True,  "text": "Includes post-session evaluation — collects feedback and迭代 improves"},
            {"check": True,  "text": "Names specific methods: hands-on labs, simulations, job aids, microlearning"},
            {"cross": True, "text": "Jumps straight to content without analysing who the learners are"},
            {"cross": True, "text": "No evaluation step — 'I just ran the session'"},
        ],
        "probes": ["How do you identify if a training programme actually changed behaviour, not just test scores?"],
    },
    {
        "question": "How do you train end-users on a new system update or workflow change without disrupting daily clinical operations?",
        "whatToListen": [
            {"check": True,  "text": "Uses bite-sized sessions or staggered rollouts to avoid overwhelming staff"},
            {"check": True,  "text": "Leverages super-users / floor champions who can support peers"},
            {"check": True,  "text": "Provides just-in-time reference materials (job aids, quick guides) at the point of care"},
            {"check": True,  "text": "Schedules training around clinical rosters — not during peak hours"},
            {"cross": True, "text": "Runs one-size-fits-all marathon sessions with no regard for clinical schedules"},
            {"cross": True, "text": "No plan for post-go-live support — leaves users to figure it out alone"},
        ],
        "probes": ["What do you do when a senior clinician refuses to attend training?"],
    },
]


def generate_personalized_questions(candidate_name: str, claims_context: str, max_questions: int = 10) -> list:
    """
    Call MiniMax API to generate personalized interview questions
    based on extracted resume claims.

    Returns a list of dicts:
      { question: str, probes: [str], category: "accountability" | "technical" | "motivation",
        whatToListen: [{check|text}|{cross|text}] }
    Falls back to [] on any error — never crashes document generation.
    """
    if not MINIMAX_API_KEY:
        print("WARNING: MINIMAX_API_KEY not set — skipping personalized questions")
        return []
    if not claims_context.strip():
        print("INFO: No claims extracted — skipping personalized questions")
        return []

    prompt = f"""You are an expert healthcare IT interviewer. A candidate named {candidate_name} has submitted their resume. Your task is to generate deeply personalized interview questions based on the specific claims they made in their resume.

For each question you generate, you MUST:
1. Reference a specific claim from the candidate's resume (by quoting or paraphrasing it)
2. Ask a question that the candidate CANNOT answer generically — they must know their specific project/role
3. Add 1-2 follow-up probes that dig into specifics (numbers, names, outcomes, your role vs team's role)

Return your response as a valid JSON array of objects. Each object must have:
- "question": the full interview question (string)
- "probes": array of probe strings (1-2 items)
- "category": "accountability" (Section 1: STAR/responsibility) OR "technical" (Section 2: EPIC/NGEMR/healthcare IT) OR "motivation" (Section 3: Why this role, will they stay?)
- "whatToListen": array of listen-guide objects, each with either:
  - {{"check": true, "text": "✅ what to listen for — good signal"}}
  - {{"cross": true, "text": "🚩 what to listen for — red flag signal"}}
  Each question must have at least 1 check and 1 cross item in whatToListen.

Rules:
- Keep each question UNDER 60 words — concise and direct, not rambling
- Each question should be answerable in 2-4 minutes in an interview
- Aim to generate 2-3 accountability questions, 2-3 technical questions, 2-3 motivation questions (up to {max_questions} total)
- Motivation questions MUST be anchored to the candidate's own career history, recent achievements, or stated goals from their resume — NOT generic questions
- For motivation: ask why this role fits THEIR specific path, not just 'why healthcare'
- Reference exact claims: "You wrote X..." not "Tell me about a time when..."
- If a claim has numbers/percentages, ask how they measured it
- If a claim has a named system/project, ask for their specific role
- If a candidate has career gaps, job changes, or the timing of an MBA/certificate — probe that specifically
- Output ONLY valid JSON array — no markdown, no explanation, no thinking

Candidate resume claims:
{claims_context}
"""

    payload = json.dumps({
        "model": "MiniMax-M2",
        "messages": [
            {"role": "system", "content": "You are a strict healthcare IT interview question generator. Output ONLY a valid JSON array, no markdown, no explanation."},
            {"role": "user",   "content": prompt},
        ],
        "temperature": 0.7,
        "max_tokens": 4096,  # Increased from 2048 to avoid truncation mid-JSON (2026-05-09)
    }).encode("utf-8")

    req = urllib.request.Request(
        f"{MINIMAX_BASE_URL}/chat/completions",
        data=payload,
        headers={
            "Authorization": f"Bearer {MINIMAX_API_KEY}",
            "Content-Type":  "application/json",
        },
        method="POST",
    )
    _content_defined = False
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            content = data["choices"][0]["message"]["content"].strip()
        _content_defined = True
    except (urllib.error.HTTPError, urllib.error.URLError) as e:
        print(f"WARNING: MiniMax API error ({e}) — skipping personalized questions")
        return []
    except Exception as e:
        print(f"WARNING: Unexpected LLM error ({e}) — skipping personalized questions")
        return []

    # Process content outside try/except — so JSON parse errors can be handled
    try:
        # Strip thinking blocks and markdown fences
        content = re.sub(r'<think[^>]*>[\s\S]*?</think>', '', content, flags=re.IGNORECASE).strip()
        content = re.sub(r'^```(?:json)?\s*', '', content, flags=re.I).strip()
        content = re.sub(r'\s*```$', '', content).strip()
        questions = json.loads(content)
        if isinstance(questions, list):
            valid = [q for q in questions
                     if q.get("category") in ("accountability", "technical", "motivation")
                     and q.get("question")]
            for q in valid:
                if "whatToListen" not in q or not isinstance(q["whatToListen"], list):
                    q["whatToListen"] = []
            print(f"LLM generated {len(questions)} personalized questions ({len(valid)} valid)")
            return valid
        else:
            print(f"WARNING: LLM returned non-list: {type(questions)}")
            return []
    except json.JSONDecodeError as e:
        print(f"WARNING: LLM response parse error ({e}) — skipping personalized questions")
        print(f"Raw response: {content[:500] if _content_defined else 'N/A'}")
        return []
    except Exception as e:
        print(f"WARNING: Unexpected LLM error ({e}) — skipping personalized questions")
        return []

# ─────────────────────────────────────────────────────────────────────────────
# DOCX tblGrid bug fix (from Phase 1)
# ─────────────────────────────────────────────────────────────────────────────

def fix_tblgrid(docx_path):
    """docx.js bug: tblGrid column widths set to 100 twips regardless of actual widths."""
    with zipfile.ZipFile(docx_path, 'r') as zin:
        names = zin.namelist()
        files = {n: zin.read(n) for n in names}

    if 'word/document.xml' not in files:
        return

    xml = files['word/document.xml'].decode('utf-8')
    TBL_TOTAL = 8640  # twips for 6-inch content width

    def fix_table(tbl_match):
        tbl = tbl_match.group(0)
        if '<w:tblGrid>' not in tbl:
            return tbl
        first_row = re.search(r'<w:tr>(.*?)</w:tr>', tbl, re.DOTALL)
        if not first_row:
            return tbl
        row_content = first_row.group(1)
        col_count = row_content.count('<w:tc>')
        if col_count == 0:
            return tbl

        raw_widths = re.findall(r'<w:tcW[^>]+w:w="(\d+)%"', row_content)
        if len(raw_widths) >= col_count:
            col_widths = [int(round(TBL_TOTAL * int(pct) / 100)) for pct in raw_widths]
        else:
            col_widths = [TBL_TOTAL // col_count] * col_count

        diff = TBL_TOTAL - sum(col_widths)
        col_widths[0] += diff

        grid_cols = ''.join(f'<w:gridCol w:w="{w}"/>' for w in col_widths)
        total_w = sum(col_widths)
        tbl = re.sub(r'<w:tblW[^>]+>', f'<w:tblW w:type="dxa" w:w="{total_w}"/>', tbl, count=1)
        tbl = re.sub(r'<w:tblGrid>.*?</w:tblGrid>', f'<w:tblGrid>{grid_cols}</w:tblGrid>', tbl, count=1)
        return tbl

    xml = re.sub(r'<w:tbl>[\s\S]*?</w:tbl>', fix_table, xml)
    files['word/document.xml'] = xml.encode('utf-8')

    tmp_path = docx_path + '.tmp'
    with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, files[name])
    os.replace(tmp_path, docx_path)

# ─────────────────────────────────────────────────────────────────────────────
# JS template — Phase 2 (reads from docx_generator_phase2.js)
# Uses CANDIDATE_JSON, OUTPUT_PATH, SCORE_VAL, BATCH_ID_VAL, PERSONALIZED_QUESTIONS replacements
# ─────────────────────────────────────────────────────────────────────────────

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

def build_js_template(section1_questions: list, section2_questions: list, section3_questions: list) -> str:
    """
    Read the docx_generator_phase2.js template and inject merged section questions.
    Uses simple string replacement — no f-string brace conflicts.
    """
    template_path = os.path.join(SCRIPT_DIR, "docx_generator_phase2.js")
    with open(template_path, "r", encoding="utf-8") as f:
        tmpl = f.read()

    s1_json = json.dumps(section1_questions, ensure_ascii=False)
    s2_json = json.dumps(section2_questions, ensure_ascii=False)
    s3_json = json.dumps(section3_questions, ensure_ascii=False)
    tmpl = tmpl.replace("SECTION1_QUESTIONS", s1_json)
    tmpl = tmpl.replace("SECTION2_QUESTIONS", s2_json)
    tmpl = tmpl.replace("SECTION3_QUESTIONS", s3_json)

    return tmpl


def main():
    candidate_name = sys.argv[1]
    json_path      = sys.argv[2]
    batch_id       = sys.argv[3]
    score          = sys.argv[4]
    role           = sys.argv[5].upper() if len(sys.argv) > 5 else ""  # e.g. "PT"

    with open(json_path) as f:
        cand_data = json.load(f)

    # ── Load ResumeScanner criteria from scores.json ────────────────────────────
    criteria_data = {}
    scores_path = f"/mnt/d/Hermes/Project/ResumeScanner/output/reports/{batch_id}/scores.json"
    if os.path.exists(scores_path):
        with open(scores_path) as f:
            all_scores = json.load(f)
        cid = cand_data.get("candidate_id", "")
        for entry in all_scores:
            if entry.get("candidate_id") == cid:
                criteria_data = {
                    "total_score":  entry.get("total_score", 0),
                    "recommendation": entry.get("recommendation", ""),
                    "criteria": entry.get("criteria", {}),
                    "warnings": entry.get("warnings", []),
                }
                break
    if not criteria_data:
        print(f"WARNING: No criteria found for candidate '{candidate_name}' in batch '{batch_id}' — recommendation will show N/A")

    # ── Compute total experience years from work_experience ──────────────────────
    work_exp = cand_data.get("work_experience", [])
    total_exp_raw = compute_total_experience(work_exp) if work_exp else 0
    if total_exp_raw > 0:
        total_experience_str = str(total_exp_raw)
        print(f"Total experience: {total_exp_raw} years")
    else:
        print(f"WARNING: Could not compute total experience from work history — showing 'See resume'")
        total_experience_str = '"See resume"'

    # ── Generic question banks ─────────────────────────────────────────────────
    GENERIC_SECTION1 = [
        {
            "question": "Tell me about a time when a task you were responsible for didn't go as planned. What did you do?",
            "whatToListen": [
                {"check": True,  "text": "Owns it fully — 'I realised I missed X, so I did Y to fix it'"},
                {"check": True,  "text": "Traces root cause without blaming others or circumstances"},
                {"check": True,  "text": "Shows learning — 'Next time I would...'"},
                {"cross": True, "text": "Deflects — 'The system was slow so I couldn't finish'"},
                {"cross": True, "text": "'We' throughout with no clarification of their specific role"},
            ],
            "probes": ["Were there moments where a project deliverable slipped? How did you handle it?"],
        },
        {
            "question": "Describe a situation where you had multiple deadlines or competing priorities. How did you manage?",
            "whatToListen": [
                {"check": True,  "text": "Structured approach — triage, communicate early if slipping"},
                {"check": True,  "text": "Finishes what they start even when it gets difficult"},
                {"cross": True, "text": "Drops things when they get hard or complicated"},
                {"cross": True, "text": "Says yes to everything then fails to deliver"},
            ],
            "probes": [],
        },
        {
            "question": "Tell me about a time you had to hold someone (or yourself) accountable for a standard or deadline.",
            "whatToListen": [
                {"check": True,  "text": "Doesn't let small issues become big problems"},
                {"check": True,  "text": "Has honest conversations early before the situation worsens"},
                {"check": True,  "text": "Follows through even when the conversation is uncomfortable"},
                {"cross": True, "text": "Avoids confrontation and lets problems fester"},
                {"cross": True, "text": "Passes problems upward without trying to resolve first"},
            ],
            "probes": [],
        },
        {
            "question": "Describe a situation where you realised a project or deliverable was at risk of missing its target. What did you do?",
            "whatToListen": [
                {"check": True,  "text": "Surfaces problems early before they become crises"},
                {"check": True,  "text": "Proposes solutions, not just reports problems upward"},
                {"check": True,  "text": "Keeps stakeholders informed so they can make informed decisions"},
                {"cross": True, "text": "Hides problems until they blow up"},
                {"cross": True, "text": "Only escalates after the deadline has already been missed"},
            ],
            "probes": [],
        },
        {
            "question": "Tell me about the biggest mistake you made in your career and what you did about it.",
            "whatToListen": [
                {"check": True,  "text": "Takes full ownership — no excuses"},
                {"check": True,  "text": "No blame on external factors or other people"},
                {"check": True,  "text": "Concrete action taken to fix the mistake"},
                {"check": True,  "text": "Clear learning applied since the incident"},
                {"cross": True, "text": "'It wasn't really a mistake, more like...' (deflection)"},
                {"cross": True, "text": "Blaming circumstances or other people"},
            ],
            "probes": [],
        },
    ]

    GENERIC_SECTION2 = [
        {
            "question": "The NGEMR project is Singapore's national healthtech initiative. Based on your experience, what do you understand about how EPIC fits into the broader NGEMR landscape?",
            "whatToListen": [
                {"check": True,  "text": "Understands NGEMR as a national-level, multi-cluster integration effort"},
                {"check": True,  "text": "Knows EPIC is the EMR system within NGEMR deployed at RHS level"},
                {"check": True,  "text": "Mentions interoperability or data exchange between clusters"},
                {"cross": True, "text": "'I'm not sure what NGEMR is' — significant gap for a Medical Informatics role"},
                {"cross": True, "text": "Only knows EPIC as a local institutional system with no national context"},
            ],
            "probes": ["How do you think patient data should flow between different clusters under NGEMR?", "What challenges do you foresee in connecting different RHS EMR systems?"],
        },
        {
            "question": "What EPIC modules or functionality are you familiar with? Walk me through what you specifically did — not just using the system, but configuring, building, or testing it.",
            "whatToListen": [
                {"check": True,  "text": "Can describe specific modules with concrete configuration or build tasks"},
                {"check": True,  "text": "Training environment build experience — set up from scratch or maintained"},
                {"check": True,  "text": "Can explain the difference between user training and technical build/configuration"},
                {"cross": True, "text": "Only describes end-user tasks with no build or configuration knowledge"},
                {"cross": True, "text": "Says 'I used Epic' but cannot describe any specific module functionality"},
            ],
            "probes": ["What was the most complex EPIC configuration you were involved in?", "Did you work with the EPIC team directly or just use the system?"],
        },
        {
            "question": "Healthcare IT implementations follow stages: Requirements → Design/Build → Testing (UAT) → Deployment → Support/Optimization. Walk me through your work and map your activities to these stages.",
            "whatToListen": [
                {"check": True,  "text": "Maps to 4+ stages with specific, concrete examples"},
                {"check": True,  "text": "Can clearly articulate their role in each stage"},
                {"check": True,  "text": "Understands the difference between deployment support and ongoing optimization"},
                {"cross": True, "text": "Vague — 'I did a bit of everything' without structure"},
                {"cross": True, "text": "Can only describe activities in one stage"},
                {"cross": True, "text": "'They assigned me where needed' without showing personal ownership"},
            ],
            "probes": ["Which stage did you spend the most time in? Why?", "Were you involved in UAT? What was your specific role?"],
        },
        {
            "question": "What do you think is the biggest challenge in implementing an EMR system like EPIC in a hospital environment? Give an example from your experience.",
            "whatToListen": [
                {"check": True,  "text": "User adoption — clinicians resistant to change, workflow disruption"},
                {"check": True,  "text": "Clinical workflow complexity — cannot break existing safe practices"},
                {"check": True,  "text": "Stakeholder management — doctors, nurses, admins all have different needs"},
                {"cross": True, "text": "Only gives technical answers, ignores human/clinical factors"},
                {"cross": True, "text": "'It's just software' — misses the healthcare context entirely"},
            ],
            "probes": ["What strategies have you used to get clinician buy-in for new systems or changes?"],
        },
        {
            "question": "Tell me about a time you had to work across different teams, institutions, or stakeholder groups. What was challenging about it?",
            "whatToListen": [
                {"check": True,  "text": "Different priorities across stakeholder groups (clinical vs admin vs IT)"},
                {"check": True,  "text": "Different levels of technology readiness across teams"},
                {"check": True,  "text": "Had to adapt communication style or approach per audience"},
                {"check": True,  "text": "Named specific people or teams — not generic 'we worked with others'"},
                {"cross": True, "text": "'It was fine' — no acknowledgment of complexity"},
                {"cross": True, "text": "No specific examples, just general statements"},
            ],
            "probes": [],
        },
    ]

    # ── Phase 2: Extract claims + generate personalized questions ───────────
    personalized_questions = []
    try:
        claims = extract_claims(cand_data)
        top_claims = filter_top_claims(claims, max_claims=8)
        claims_context = claims_to_prompt_context(top_claims)
        if claims_context:
            personalized_questions = generate_personalized_questions(candidate_name, claims_context, max_questions=7)
        else:
            print("INFO: No claims extracted from candidate JSON — using generic questions only")
    except Exception as e:
        print(f"WARNING: Claim extraction failed ({e}) — using generic questions only")

    # Split personalized by category
    personalized_s1 = [q for q in personalized_questions if q.get("category") == "accountability"]
    personalized_s2 = [q for q in personalized_questions if q.get("category") == "technical"]
    personalized_s3 = [q for q in personalized_questions if q.get("category") == "motivation"]

    # Merge: personalized first, then generic fallback
    # Passes through whatToListen from personalized questions (for 🎯 questions)
    def merge_section(personalized, generic, min_count=5):
        result = []
        for pq in personalized[:min_count]:
            # Personalized questions: do NOT include whatToListen — keep them focused and short
            # The generic fallback questions already carry full whatToListen for each section
            result.append({
                "question": pq["question"],
                "probes": pq.get("probes", []),
                "whatToListen": pq.get("whatToListen", []),
                "isPersonalized": True,
            })
        generic_idx = 0
        while len(result) < min_count and generic_idx < len(generic):
            gq = generic[generic_idx]
            if not any(gq["question"] == r["question"] for r in result):
                result.append({
                    "question": gq["question"],
                    "whatToListen": gq.get("whatToListen", []),
                    "probes": gq.get("probes", []),
                    "isPersonalized": False,
                })
            generic_idx += 1
        return result

    section1_questions = merge_section(personalized_s1, GENERIC_SECTION1, min_count=2)
    section2_questions = merge_section(personalized_s2, GENERIC_SECTION2, min_count=2)
    section3_questions = merge_section(personalized_s3, GENERIC_MOTIVATION, min_count=2)

    # Append PT-specific training questions to Section 2
    if role == "PT":
        for q in PT_TRAINING_QUESTIONS:
            section2_questions.append({
                "question": q["question"],
                "whatToListen": q.get("whatToListen", []),
                "probes": q.get("probes", []),
                "isPersonalized": False,
            })

    personalized_count = sum(1 for q in section1_questions if q.get("isPersonalized")) + \
                          sum(1 for q in section2_questions if q.get("isPersonalized")) + \
                          sum(1 for q in section3_questions if q.get("isPersonalized"))

    # Build JS template with merged sections
    js_content = build_js_template(section1_questions, section2_questions, section3_questions)
    js_content = js_content.replace("CANDIDATE_JSON", json.dumps(cand_data, indent=2))
    js_content = js_content.replace("OUTPUT_PATH", f"'{os.path.join(OUTPUT_DIR, candidate_name.replace(' ', '_') + '_Interview_Prep.docx')}'")
    js_content = js_content.replace('"SCORE_VAL"', f'"{score}"')
    js_content = js_content.replace("BATCH_ID_VAL", f"'{batch_id}'")
    js_content = js_content.replace("CRITERIA_DATA", json.dumps(criteria_data) if criteria_data else "{}")
    js_content = js_content.replace("TOTAL_EXP_YEARS", total_experience_str)

    # If criteria_data has no recommendation, fall back to score argument
    if not criteria_data.get("recommendation"):
        js_content = js_content.replace('criteriaData.recommendation || "N/A"', f'"Score: {score}/100 — see breakdown"')

    safe_name     = candidate_name.replace(" ", "_")
    output_path   = os.path.join(OUTPUT_DIR, f"{safe_name}_Interview_Prep.docx")

    # Ensure node_modules/docx available in /tmp
    node_modules = "/tmp/node_modules"
    if not os.path.exists(node_modules):
        os.makedirs(node_modules, exist_ok=True)
        subprocess.run(["cp", "-r",
                        "/home/dennis/.hermes/node/lib/node_modules/",
                        "/tmp/"],
                       capture_output=True)

    tmp_js = tempfile.NamedTemporaryFile(suffix=".js", mode="w", delete=False)
    tmp_js.write(js_content)
    tmp_js.flush()
    tmp_js_path = tmp_js.name

    try:
        result = subprocess.run(
            ["node", tmp_js_path],
            cwd="/tmp",
            capture_output=True, text=True, timeout=60
        )
        if result.returncode != 0:
            print("JS ERROR:", result.stderr[:2000])
            sys.exit(1)
        print("JS:", result.stdout.strip())
    finally:
        os.unlink(tmp_js_path)

    # Fix docx.js tblGrid bug
    fix_tblgrid(output_path)

    # Verify DOCX
    try:
        with zipfile.ZipFile(output_path) as z:
            assert "word/document.xml" in z.namelist()
        print(f"VERIFIED: Valid DOCX ({os.path.getsize(output_path)} bytes)")
    except Exception as e:
        print(f"VERIFICATION FAILED: {e}")
        sys.exit(1)

    # ── Email to Boss ───────────────────────────────────────────────────────
    if personalized_count > 0:
        pq_note = f"\n- Sections 1, 2 & 3: {personalized_count} personalized questions (🎯) woven in from resume claims — each with contextual red/green flags"
    else:
        pq_note = ""

    msg = MIMEMultipart()
    msg["From"]    = GMAIL_USER
    msg["To"]      = TO_EMAIL
    msg["Subject"] = f"Interview Prep — {candidate_name}"

    body = f"""Hi Boss,

Please find attached the Interview Preparation document for {candidate_name}.

Source: ResumeScanner (Batch {batch_id})
ResumeScanner Score: {score}/100

This document includes:
- Candidate summary & scoring
- Flag signals (HIRE / NO-HIRE)
- Section 1: Accountability & Responsibility (STAR-format questions)
- Section 2: Technical — Epic, NGEMR, IT Project Stages
- Section 3: Motivation, Self-Awareness & Retention (why this role, will they stay?)
{pq_note}
- Scoring rubric & decision guide

Generated by Hermes InterviewGenerator (Phase 2 — personalized questions with contextual flags).

Best,
Hermes Agent"""

    msg.attach(MIMEText(body, "plain"))

    with open(output_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={safe_name}_Interview_Prep.docx")
        msg.attach(part)

    if not GMAIL_APP_PASSWORD:
        print(f"EMAIL SKIPPED: GMAIL_APP_PASSWORD not set — doc saved to {output_path}")
    else:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
            server.send_message(msg)
        print(f"EMAIL SENT: to {TO_EMAIL}")
        print(f"OUTPUT: {output_path}")

if __name__ == "__main__":
    main()
