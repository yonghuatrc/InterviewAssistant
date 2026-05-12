# InterviewAssistant

NUHS interview automation pipeline — generates personalized interview prep documents and post-interview assessment forms via MiniMax LLM.

## Pipelines

### 1. Interview Prep Document
- Takes candidate resume JSON (from ResumeScanner)
- Extracts claims via `experience_extractor.py` (regex-based)
- Generates personalized interview questions via MiniMax LLM
- Produces a styled DOCX using Node.js `docx` library
- Emails to `dennis_ng@nuhs.edu.sg`

### 2. Interview Assessment Form
- Takes interview transcript + prep DOCX
- Calls MiniMax LLM to generate 150-200 word narrative assessment
- Fills standardised GMIO HR template DOCX
- Emails to `dennis_ng@nuhs.edu.sg`

## Usage

### Generate Interview Prep
```bash
python3 generate_interview.py "<Name>" <json_path> <batch_id> <score> [role]
```

### Generate Assessment
```bash
python3 generate_assessment.py "<Candidate Name>" "<transcript text>"
```

Or pipe transcript via stdin:
```bash
echo "<transcript>" | python3 generate_assessment.py "<Candidate Name>"
```

## Dependencies

### Python 3.12+
- `requests` (for LLM API calls)

### Node.js
- `docx` library (for DOCX generation)

Install Node deps:
```bash
npm install docx
```

## Environment Variables

| Variable | Purpose |
|----------|---------|
| `MINIMAX_API_KEY` | MiniMax LLM API key |
| `SMTP_PASSWORD` | Gmail app password for sending emails |

## Project Structure

```
├── generate_assessment.py      # Assessment pipeline
├── generate_interview.py       # Interview prep pipeline
├── experience_extractor.py     # Resume claim extraction (regex)
├── send_email.py               # Email sender (Gmail SMTP)
├── docx_generator_phase2.js    # DOCX template engine
├── docx_generator.js           # Phase 1 template (static)
├── templates/                  # HR assessment form templates
├── output/                     # Generated prep documents (gitignored)
├── assessment_output/          # Generated assessment forms (gitignored)
└── config.yaml                 # Project configuration
```

## License

Internal — NUHS Medical Informatics
