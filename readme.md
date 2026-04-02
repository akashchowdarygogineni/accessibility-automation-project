# ♿ Accessibility Automation Project

## Overview
The Accessibility Automation Project streamlines PDF accessibility validation using a robust 3-stage pipeline:

1. PAC stage: desktop automation with `pywinauto`
2. PREP stage: Java JAR/API-based processing
3. Comparison stage: result consolidation and report generation

The workflow is designed for bulk PDF processing, version-based execution, and operational reporting with Slack notifications.

## Architecture
Add architecture diagram here

<!-- Example: docs/architecture-diagram.png -->

## Workflow
![Workflow Diagram](docs/workflow.png)

1. Input PDFs are collected from a target version folder (for example, `v1`, `v1.5`).
2. `pac.py` runs PAC desktop checks and generates PAC outputs.
3. `prep.py` processes files through the PREP JAR/API.
4. `slack.py` compares PAC and PREP outcomes.
5. Consolidated Excel reports are generated.
6. Slack notifications are sent for status and summary.

## Project Structure
```text
.
|-- pipeline.py
|-- pac.py
|-- prep.py
|-- slack.py
|-- prep_slack.py
|-- slack_only_runner.py
|-- s3_uploader.py
|-- readme.md
|-- docs/
|   `-- workflow.png
|-- s3/
|-- v1/
|   |-- pac_processed/
|   |-- prep_processed/
|   |-- prep_results/
|   `-- prep_skipped/
`-- v1.1/
```

## Features
- Version-based execution (`v1`, `v1.5`, etc.)
- End-to-end orchestration across PAC, PREP, and comparison
- Bulk PDF processing support
- Excel report generation for audit and review
- Slack notifications for execution updates
- Error handling and retry mechanisms for stability

## Tech Stack
- Language: Python
- Libraries:
  - `pandas`
  - `openpyxl`
  - `pywinauto`
  - `requests`
  - `python-dotenv`
  - `slack_sdk`
- Integrations:
  - PAC desktop tool
  - Java PREP JAR/API
  - Slack API

## Installation Steps
1. Clone the repository.
2. Create a virtual environment.
3. Activate the environment.
4. Install dependencies.
5. Configure `.env` values.
6. Ensure PAC desktop tool and Java runtime are installed.

```bash
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install pandas openpyxl pywinauto requests python-dotenv slack_sdk
```

## Execution Commands
```bash
# Full pipeline: PAC + PREP + comparison
python pipeline.py

# PAC stage only
python pac.py

# PREP stage only
python prep.py

# Comparison and reporting only
python slack.py

# PREP + Slack flow
python prep_slack.py

# Comparison-only runner
python slack_only_runner.py
```

## Input and Output
### Input
- Source PDF files placed in version folders
- Version identifier (`v1`, `v1.5`, etc.)
- Environment configuration from `.env`

### Output
- Creates stage-wise output folders (for example: `pac_processed`, `prep_processed`, `prep_results`, `prep_skipped`) inside the selected version directory
- Generates multiple Excel output files for PAC summary, PREP summary, comparison reports, and final consolidated accessibility reporting
- Stores consolidated artifacts for review and audit tracking

Add output screenshots here

## Slack Integration
Use a `.env` file for Slack configuration:

```env
SLACK_BOT_TOKEN=xoxb-your-token-here
SLACK_CHANNEL_ID=C0123456789
SLACK_ENABLED=true
```

## Test Cases
- Validate single-file pipeline execution.
- Validate bulk PDF processing flow.
- Validate version-based execution (`v1`, `v1.5`).
- Validate retry behavior for transient failures.
- Validate Slack notification and report generation.

## Known Limitations
- PAC cannot run in server or headless mode.
- PAC does not provide native CLI/API support for direct automation.
- PAC requires a Windows UI environment and active desktop session.
- PAC stage cannot run reliably inside standard CI/CD pipelines.
- Overall workflow behavior can be affected by third-party tool changes.

## Future Improvements
- Add a centralized monitoring dashboard.
- Introduce configurable workflow profiles per version/client.
- Expand automated validation coverage with regression checks.
- Improve scalability with controlled parallel processing.
- Add richer analytics and trend summaries in reports.

## Author
- Name: akash gogineni
- Role: Accessibility Automation Engineer
- Contact: www.linkedin.com/in/akash-gogineni-b68102300

