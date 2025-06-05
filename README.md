# CS-270 Late Pass Google Sheet Scraper & Emailer

Script to automatically email all students who requested to use one of their late passes on a homework assignment for CS
270 at Drexel University.
Automated to run every Saturday at `12:01 AM`, the cutoff for late pass requests. Updates late pass Google sheet as
well.

## Requirements

- Python 3
- Logged into Outlook
- Google Cloud credentials
- On Mac/Linux: `at` daemon, may not be running/installed by default
- ### Install all requirements:

```bash
pip install -r requirements.txt
```

## Files

- `Setup.py`: Script to setup Windows Task Scheduler to automatically run `Main.py` at `12:01 AM` on Saturdays
- `Main.py`: Main script, scrapes Google sheets and generates messages for students who requested late pass uses
- `.env`: Environment file with Google Sheets IDs and path to Google Cloud credentials (must be given/created)
- Google Cloud credentials file (must be given)

## Usage/Setup

> ⚠️ On Windows, `Setup.py` must be ran as administrator, or will not successfully create weekly task\
> ⚠️ Mac/Linux emailing currently not supported

1. Create (or be given) production and test Google Sheets for both the Google form responses and the log of late pass
   uses.  
   Here is the formatting for the test response sheet. For each set of test cases, make sure there is an empty row
   between them in the table.

   | Timestamp         | Enter your full name (first then last) | user ID (initials followed by digits, you don't need the "@drexel.edu") | Section | Type of Assignment (note: late pass must be requested BEFORE deadline) | for what asessment do you want the late pass? | Choose Lab | Choose Homework Assignment      |
                        |-------------------|----------------------------------------|-------------------------------------------------------------------------|---------|------------------------------------------------------------------------|-----------------------------------------------|------------|---------------------------------|
   | 4/17/2025 21:23:34 | John Doe                               | jd1234                                                                  | prof1   | Homework                                                               |                                               |            | HW1: Homework 1 (Due January 1) |
   |          ⠀         |                                        |                                                                         |         |                                                                        |                                               |            |                                 |
   | 4/17/2025 21:23:34 | Jane Doe                               | jd5678                                                                  | prof2   | Homework                                                               |                                               |            | HW2: Homework 2 (Due March 31)  |

   Here is the formatting for the test late pass uses sheet. For each set of test cases, make sure there is an empty row
   between them in the table.

   | Last | First | email  | instructor | P1  | P2  | other notes            |
         |-----|-------|--------|------------|-----|-----|------------------------|
   | aaa | aaa   | aaa    | aaa        | aaa | aaa | last email: hw4        |
   | Doe | John  | jd1234 | prof1      | hw1 |     |                        |
   |  ⠀   |       |        |            |     |     |                        |
   | Doe | Jane  | jd5678 | prof2      | hw2 | hw3 | came in for lab makeup |

   The second row in this table must have `last email: hw#` in the `other notes` column, as it denotes what the last
   homework emails were sent for.


2. Take note of all the IDs for each Google sheet and the path/filename for the Google Cloud credentials file.
   > **Google Sheet ID Example**: `https://docs.google.com/spreadsheets/d/abc123/edit?gid=0#gid=0` The ID is the string
   between `/d/` and `/edit?`

3. Run `Setup.py` with the `--setup` flag. The program will prompt you to enter each sheet ID and the Google Cloud
   credentials file so they can be stored in the `.env` file. This should be the only time you ever need to use the
   `--setup` flag, unless you want to change the `.env` at a later date.
   > If this is not the first time `Setup.py` has been ran with the `--setup` flag, you will be prompted for each `.env`
   variable you wish to override.

4. If you wish to test the program before scheduling it to run with the production sheets, you may run `Setup.py` with
   the `--test` flag, which schedules `Main.py` to be ran with the test sheets in `1` minute from now.
5. If you wish to test the program without scheduling, run `Main.py` with the `--test` flag
6. To schedule `Main.py` to run every Saturday at `12:01 AM`, run `Setup.py` with no arguments

## Arguments

### `Setup.py`

- `--setup`: Has user enter Google Sheet IDs for the production and tests.
- `--python-path`: Specify path to python executable, defaults to executable used when running the script
- `--test`: Schedules `Main.py` to run 1 minute after `Setup.py` execution for testing purposes, passes through `--test`
  flag to `Main.py`

### `Main.py`

- `--test`: Runs `Main.py` with test Google Sheets instead of real ones