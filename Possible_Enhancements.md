# ResumeScanner & InterviewGenerator — Possible Enhancements Assessment

**Date:** 2026-04-26
**Author:** Hermes Agent
**Status:** Assessed — held for later implementation

---

## 1. LLM-Based Scoring for ResumeScanner

### Proposed Change
Replace keyword/heuristic scoring (`scorer.py`) with MiniMax LLM scoring, where the LLM reads the Job Description, priority rubric, and resume text to produce structured per-criterion scores with reasoning.

### Current System (scorer.py)
- Pure keyword/pattern matching — counts occurrences of strong keywords (e.g. "emr", "epic", "SAP") vs weak keywords (e.g. "intern", "trainee")
- Score 0–5 per criterion, weighted total, max 100
- Evidence = keyword list (e.g. `"Matched keywords: emr, epic, SAP"`)
- Runs in milliseconds, costs $0, deterministic, fully reproducible
- No LLM involved at all

### Proposed LLM System
```
System prompt: "You are a strict healthcare IT resume scorer. Score 0–5 per criterion
using this rubric. Return structured JSON."

User prompt: JD text + Priority YAML (weights, guidance, bands) + Resume text

Expected JSON output:
{
  "criteria_scores": [
    {"criterion": "systems_emr_exposure", "score": 4,
     "reasoning": "Candidate led EPIC Bed Management rollout at NTFGH...",
     "evidence": "EPIC certification, 2 years NGEMR exposure, SAP integration project"}
  ],
  "total_score": 74,
  "recommendation_band": "hold_further_review"
}
```

### Scoring Rubric (from priority YAML — already structured)
Each criterion has explicit guidance per score level:
- Score 4: "Strong healthcare systems or EMR exposure"
- Score 5: "Deep directly relevant NGEMR, Epic, EMR, or healthcare platform experience"

The LLM would apply this guidance to actual resume claims.

---

## 2. Options

### Option A: Hybrid Scoring (Recommended Starting Point)

```
Heuristic scorer runs first (fast, free)
  → Produces initial scores + ranking immediately

Borderline candidates (heuristic score 45–80) flagged
  → LLM re-scores ONLY those candidates
  → Final ranking uses LLM scores for borderline, heuristic for clear accept/reject

Candidates with extraction_confidence = "low"
  → LLM re-scores regardless of score
```

| Pros | Cons |
|---|---|
| LLM called for ~30% of candidates — cost stays negligible | Two scoring passes for borderline — added complexity |
| Heuristic is fallback if MiniMax API is down | Borderline threshold needs tuning (45–80?) |
| Can A/B test hybrid vs heuristic-only rankings | Different candidates scored by different methods — needs disclosure in report |
| Faster candidates skip LLM wait entirely | |

### Option B: Full LLM Replacement

Replace `scorer.py` entirely. Every candidate scored by MiniMax.

| Pros | Cons |
|---|---|
| Single scoring system — cleaner | Hard dependency on MiniMax API — pipeline fails if API is down |
| Genuinely understands depth, context, quality | 4–8 sec/candidate sequentially. 10 candidates = 40–80 sec extra wait |
| Evidence becomes actual reasoning text | PII risk: resume (names, emails, phone, work history) sent to MiniMax API |
| Better nuance for role relevance and achievement quality | No fallback if API key is missing |
| | Harder to debug non-deterministic score changes |

---

## 3. Token & Cost Estimate

Per candidate LLM call:
- JD text (MI role, ~300 words) ≈ 400 tokens
- Priority YAML (rubric, guidance, bands, ~200 words) ≈ 250 tokens
- Resume text (average, ~600 words) ≈ 800 tokens
- System prompt + scoring instructions ≈ 200 tokens
- LLM reasoning output ≈ 300 tokens
- **Total: ~2,000 tokens/candidate**

| Batch Size | Heuristic Cost | Hybrid Cost (~30% LLM) | Full LLM Cost |
|---|---|---|---|
| 3 candidates | $0 | ~$0.01 | ~$0.04 |
| 10 candidates | $0 | ~$0.04 | ~$0.13 |
| 20 candidates | $0 | ~$0.08 | ~$0.26 |

**Cost is negligible.** Real costs are latency and API reliability.

---

## 4. MiniMax Structured Output — Technical Feasibility

**Current InterviewGenerator MiniMax usage:**
- Temperature 0.7 — creative but inconsistent
- No `response_format` parameter — relies on regex stripping of markdown fences
- Freeform JSON parsed with try/except fallback — ~90% success rate

**Required for scoring (needs higher reliability):**
- Structured JSON with consistent schema per candidate
- Per-criterion: `{score: int, reasoning: str, evidence: str}`
- Must parse reliably every time

**Options for structured output:**
1. **Careful prompting + JSON parsing** — works ~90% with temperature=0.5; 10% need fallback handling
2. **`response_format: {"type": "json_object"}`** — OpenAI-compatible; MiniMax-M2 may support this; forces valid JSON
3. **Tool calling / function calling** — define a `score_resume` function with JSON schema; most reliable; requires API support verification
4. **Keep both scores in scores.json** — heuristic score + LLM score side-by-side; display LLM in report, keep heuristic as backup audit trail

**Recommendation:** Test `response_format: {"type": "json_object"}` with MiniMax-M2 first. If unreliable, fall back to function calling which has stronger output guarantees.

---

## 5. Risk Analysis

| Risk | Level | Mitigation |
|---|---|---|
| MiniMax API down mid-batch | High | Keep heuristic scorer as fallback; if API fails mid-batch, use heuristic scores |
| LLM scoring inconsistency (same resume = different score on rerun) | Medium | temperature=0.0; structured output; cache scores in scores.json |
| PII leaves system (resume → MiniMax API) | High | Candidate PII (name, email, phone, work history) transmitted to third-party API. Requires Boss's explicit decision. May require candidate privacy notice in hiring process. |
| Latency: 10 candidates = 40–80 sec extra wait | Low | Acceptable; background processing; Boss already tolerates pipeline wait |
| Prompt injection / malformed resume breaks LLM | Low | Try/except per candidate; log and skip if parsing fails |
| Borderline threshold tuning needed | Medium | Start with 45–80 range, adjust based on manual review of a few batches |
| scores.json format change breaks report_generator.py | Low | LLM scorer should output same ScoreResult dataclass — no format change needed for display |

---

## 6. Unknowns — Requires Boss's Decision

1. **PII sensitivity** — Candidate resume data (names, emails, phone numbers, work history) would be transmitted to MiniMax API (third-party). Is this acceptable given it is a private internal tool? Does the hiring process need a candidate privacy notice?

2. **Hybrid vs full replacement** — Does Boss want LLM to replace heuristic entirely, or only supplement it for borderline cases?

3. **Structured output verification** — Should test MiniMax-M2's `response_format` parameter before committing to an approach.

4. **scores.json format** — Should it contain both heuristic AND LLM scores side-by-side (for audit), or replace heuristic entirely?

---

## 7. Implementation Plan (When Approved)

### Phase 1 — Hybrid Scorer (lowest risk)
- Add `llm_scorer.py` alongside `scorer.py` (existing scorer is NOT deleted)
- New function `score_with_llm(candidate_json, priority_yaml, jd_text)` → returns ScoreResult
- Called only for: heuristic score in 45–80 range OR extraction_confidence = "low"
- `scores.json` gains a `scoring_method: "llm"` or `"heuristic"` field per candidate
- scores.json output schema unchanged — ScoreResult dataclass remains compatible
- report_generator.py unchanged — it reads scores.json the same way

### Phase 2 — Validation
- Run 3–5 historical batches with hybrid scoring
- Compare rankings to heuristic-only output
- Boss manually reviews borderline candidates to validate LLM quality

### Phase 3 — Optional Full Replacement
- Only if Boss confirms Phase 2 results are meaningfully better
- Add `use_llm_scoring: true` flag to priority YAML
- Remove heuristic-only path once LLM is validated

---

## 8. Files That Would Change

| File | Change |
|---|---|
| `src/scorer.py` | No change — kept as fallback |
| `src/llm_scorer.py` | **New** — LLM scoring function |
| `src/run_stage3.py` | Call `llm_scorer` for borderline/low-confidence candidates |
| `src/paths.py` | No change |
| `input/priorities/<ROLE>.yaml` | Optional: add `use_llm_scoring: true` flag per role |
| `output/reports/<batch>/scores.json` | New field: `scoring_method` per candidate |

---

## 9. Priority

**Low — not urgent.** Current heuristic scorer works and is fit for purpose. LLM scoring is an enhancement that improves scoring nuance but is not required for the system to function.

**Recommended next step:** Boss decides on PII sensitivity stance. If acceptable, run a proof-of-concept test with MiniMax-M2 using `response_format: json_object` on 3 historical resumes before committing to implementation.
