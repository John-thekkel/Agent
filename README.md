# Job Application Agent

A Python CLI that finds remote jobs across several boards, scores fit, drafts a
tailored resume, cover letter and outreach email, then has an LLM critic grade
the draft against a rubric. Failing drafts get one revision; anything still
below the bar is held back from sending. Every email needs manual approval.

## Architecture

Scout → Matcher → Writer → Critic ──(fail)──▶ Writer (one revision) ──▶ Critic
                                      │
                                   (pass) ──▶ output/ drafts ──▶ `send` (manual approval)

Four stages run as sequential model calls in one program. They are roles in a
pipeline, not independently deployed agents.

## Evaluation

The critic scores each draft 0–100: grounding 30%, role relevance 25%, evidence
specificity 20%, completeness 15%, professional quality 10%. A draft passes at
≥80 overall and ≥4/5 on grounding. `python agent.py eval` rescores saved drafts
into `output/evaluations.csv`.

| Metric | Result |
|---|---|
| Drafts evaluated | [N] |
| Mean score, first draft | [X] |
| Mean score after revision | [Y] |
| Pass rate | [Z%] |

## Security and privacy

- The matcher flags prompt injection in job listings before they reach the writer.
- Profile data and job descriptions are sent to the Anthropic API.
- Langfuse traces contain stage names, model, command, source and job ID only.
- Credentials, drafts and job history are git-ignored. Nothing sends without approval.

## Known limitations

- Relevance filtering uses job titles only, not full descriptions.
- The "up to 3 years" preference is guidance to the LLM, not a parsed filter.
- No retry or backoff logic for job-board APIs. [edit to match the code]

## Built with

Python, Anthropic API, Langfuse. Developed with AI coding assistance (GitHub Copilot).
