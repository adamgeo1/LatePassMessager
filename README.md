# CS-270 Late Pass Google Sheet Scraper & Emailer

Script to automatically email all students who requested to use one of their late passes on a homework assignment for CS 270 at Drexel University.
Automated to run every Saturday at `12:01 AM`, the cutoff for late pass requests. Updates late pass Google sheet as well.

## Requirements
- Python 3
- Logged into Outlook
- Google Cloud credentials
- ### Install all requirements:
```bash
pip install -r requirements.txt
```

## Files
- `Setup.py`: Script to setup Windows Task Scheduler to automatically run `Main.py` at `12:01 AM` on Saturdays
- `Main.py`: Main script, scrapes Google sheets and generates messages for students who requested late pass uses
- Google Cloud credentials file (must be given)

## Usage/Setup
>⚠️ Must be ran on Windows to automate, can be ran manually on Unix
1. Run `Setup.py` to create scheduled task to run `Main.py` every Saturday at `12:01 AM`
2. Create (or be given) `.env` file with Google sheets ids (string between `/d/` and `/edit` within URL) and Google Cloud credentials JSON

## Arguments
### `Setup.py`
- `--python-path`: Specify path to python executable, defaults to executable used when running the script