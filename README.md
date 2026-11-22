# Job Application Agent

## Setup
```
pip install anthropic requests python-docx python-dotenv langfuse
```
Copy `.env.example` to `.env` and add your own credentials:
```
Copy-Item .env.example .env
```
- Gmail app password: Google Account > Security > 2-Step Verification > App passwords.
- Adzuna (free key at developer.adzuna.com) adds India listings; it is not restricted to remote roles.
- Jobicy, Himalayas, and We Work Remotely are fetched from their public API/RSS feeds without extra credentials.
- Jobicy asks API clients to link to and credit its listings; the agent preserves each source job URL.
- Langfuse tracing is enabled only when both Langfuse keys are set. Traces include stage names,
  model, command, source, and job ID; prompts, candidate details, job descriptions, and generated
  application text are not sent to Langfuse.
- `config.json` contains placeholder profile fields. Replace them with your own accurate resume
  facts and preferences before running `find`. Do not commit personal profile data or credentials.
- `.gitignore` excludes `.env`, generated applications, and local job history.

## Use
```
python agent.py find   # scout, match, write, and critique; drafts saved under ./output/
python agent.py eval   # rescore saved drafts; write rubric scores to ./output/evaluations.csv
python agent.py send   # show approved-quality drafts; you approve each email before sending
```
Run `find` daily. `output/applications.csv` is your tracker.

## Pipeline and evaluation
- **Scout** gathers and deduplicates listings from Remotive, RemoteOK, Arbeitnow, Jobicy,
  Himalayas, and We Work Remotely. Adzuna is added when credentials are configured.
- **Matcher** scores experience, preferences, pay fit, and possible prompt injection.
- **Writer** creates the tailored resume content, cover letter, and outreach email.
- **Critic** scores factual grounding (30%), role relevance (25%), evidence specificity (20%),
  completeness (15%), and professional quality (10%). A score below 80/100 or grounding below
  4/5 triggers one writer revision. Drafts that still fail are saved as `needs_review` and are
  not eligible for `send`.
- `eval` reruns the critic on saved drafts and stores aggregate results in `evaluations.csv`.
- Search terms include .NET/C#, Python, Java, backend/software development, AWS, Flask, and
  machine learning. Roles requiring up to three years are in scope. Generated resume details
  come from the profile and project entries you provide in `config.json`.

## Notes
- Edit `config.json` to change roles, pay goals, score threshold, or your resume facts.
- The agent only rewords your real experience. Read each draft before approving.
- Many jobs have no application email. For those, `send` shows the apply link and you upload the generated resume there.
- Add accurate, measurable results to your experience bullets for stronger tailored resumes.
