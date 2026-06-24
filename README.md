[README.md](https://github.com/user-attachments/files/29302208/README.md)
# Weekly Project Dashboard Generator

An internal Professional Services tool that helps PMs generate standardized
weekly PowerPoint project dashboards from ADI List, Clarizen project plan,
meeting transcript, and optional client logo uploads.

## What The App Does

- Upload an ADI List CSV to populate the RAID table.
- Upload a Clarizen Project Plan Excel export to populate the Gantt chart.
- Upload one or more meeting transcripts to suggest a Project Status narrative.
- Optionally upload a client logo for the bottom-left footer.
- Review and edit selected RAID items, deliverables, and status content.
- Generate and download a one-slide PowerPoint dashboard.

The ADI List remains the source of truth for RAID items. Clarizen remains the
source of truth for deliverables. Transcript content only suggests a Project
Status narrative and review items.

## Local Setup

Open Windows PowerShell in the project folder.

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run app.py
```

Streamlit normally opens `http://localhost:8501` in the browser.

## Using The App

1. Enter the Project Information fields.
2. Optionally upload a PNG, JPG, or JPEG client logo.
3. Upload the ADI List CSV and review the selected RAID items.
4. Optionally upload the Clarizen Excel export and review deliverables.
5. Optionally upload one or more meeting transcripts.
6. Review or edit all suggested content.
7. Select **Generate PowerPoint**, then **Download PowerPoint**.

PNG logos with transparent backgrounds generally look best. If no client logo
is uploaded, the bottom-left footer remains blank. The Cogsdale logo remains in
the bottom-right.

## Input Files

### ADI List CSV

Required columns:

- `Action Items`
- `Status`
- `Owner`
- `Priority`
- `Notes`
- `Modified`
- `Due Date`
- `TestRail/JIRA Link`

`Attachments` is optional. Extra whitespace in column headers is accepted.

### Clarizen Project Plan

Supported formats: `.xlsx` and `.xls`.

Required columns:

- `State`
- `Name`
- `Start Date`
- `Level`
- Either `Due Date` or `End Date`

`Resource` is preferred for the displayed Gantt owner. `Owner` is used as a
fallback when `Resource` is blank or unavailable.

### Meeting Transcripts

Supported formats:

- TXT
- VTT
- DOCX

Multiple transcripts may be uploaded and combined. Suggestions require PM
review and do not automatically modify RAID items or Clarizen deliverables.

### Client Logos

Supported formats:

- PNG
- JPG
- JPEG

The logo is used only while generating the PowerPoint and is not permanently
stored by the app.

## Output

The generated PowerPoint is downloaded through the Streamlit app. The generator
saves and validates a temporary completed file before exposing the download.
Generated PowerPoints must not be committed to GitHub.

For a command-line sample:

```powershell
.\.venv\Scripts\python.exe generate_dashboard.py `
  --data sample_inputs\dashboard_sample.json
```

The default command-line output is written under `generated/`, which is created
automatically and ignored by Git.

## Testing

```powershell
.\.venv\Scripts\python.exe -m pytest
```

The tests cover parsers, timeline behavior, PowerPoint layout, client logos,
and generated-file integrity.

## Repository Structure

```text
weekly-dashboard-generator/
|-- app.py
|-- dashboard_defaults.py
|-- generate_dashboard.py
|-- parse_adi_export.py
|-- parse_clarizen_plan.py
|-- parse_meeting_transcript.py
|-- requirements.txt
|-- README.md
|-- .gitignore
|-- templates/
|   `-- weekly-status-template.pptx
|-- sample_inputs/
|   |-- README.md
|   `-- dashboard_sample.json
`-- tests/
    `-- test_*.py
```

## Security And Data Handling

This tool may process confidential client and project data. Do not commit real
client files, ADI exports, Clarizen exports, transcripts, logos, generated
dashboards, credentials, or secrets to the repository.

Uploads are processed in memory where practical. Temporary PowerPoint files use
the operating system's temporary storage and are removed after the completed
bytes are read. The app does not include live SharePoint or Clarizen
integrations.

Only sanitized, fictional samples belong in `sample_inputs/`.

## Pre-commit Checklist

- The app runs locally.
- All tests pass.
- No real client files are staged.
- No generated PowerPoints are staged.
- No `.venv` folder is staged.
- No credentials or secrets are staged.
