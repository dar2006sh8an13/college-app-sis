# Student Information System (SIS)

A Python desktop application for managing student information, built with Tkinter. The application source is centered around `collegeapp.py`, with Excel export support through `openpyxl`.

## Features and Project Files

- Tkinter-based desktop interface.
- Student information workflows implemented in `collegeapp.py`.
- Excel export support using `openpyxl`.
- Local JSON data files for application state and records.

| File | Purpose |
|---|---|
| `collegeapp.py` | Main application interface and logic |
| `requirements.txt` | Python package dependencies |
| `users.json`, `student_data.json`, `session.json` | Local runtime data; keep real data out of version control |

## Requirements

- Python 3.8 or later
- Tkinter (included with most standard Python installations)
- `openpyxl`

## Setup

```bash
git clone https://github.com/dar2006sh8an13/college-app-sis.git
cd college-app-sis
python -m venv .venv
```

Activate the environment:

```powershell
.\.venv\Scripts\Activate.ps1
```

Install dependencies:

```bash
pip install -r requirements.txt
```

## Run

```bash
python collegeapp.py
```

## Data and Security

The application may create local JSON files for user accounts, sessions, and student records. Use synthetic data for development. Do not commit real student information, passwords, session tokens, or database exports. For production use, add appropriate authentication, access controls, backups, and privacy safeguards.

## Contributing

1. Create a feature branch.
2. Make a focused change.
3. Test the desktop application and any export workflows.
4. Open a pull request with verification steps.

## License

No license is specified in this repository yet.
