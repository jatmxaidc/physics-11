# Orbit — Class 11 NEET Physics Study App

A fully local demo application for the Class 11 CBSE + NEET Physics programme.

## What is included

- Student and Teacher login
- SQLite database stored locally
- Nine original demo lessons covering Mathematical Tools and Units and Measurements
- Forty-five original daily-test questions
- Automatic objective marking
- 70% lesson mastery gate with remediation below 70%
- Error log and 48-hour correction workflow
- Notes and highlight categories
- Full 138-lecture schedule imported from the programme CSV
- Locked roadmap for all remaining Class 11 chapters
- Teacher dashboard with:
  - Test attempts
  - Topic accuracy
  - Weak-topic analysis
  - Next-day review queue
  - Question miss rates
  - Individual student progress

## Demo accounts

### Student

- Email: `student@physics.local`
- Password: `student123`

### Teacher

- Email: `teacher@physics.local`
- Password: `teacher123`

## Start the demo

Double-click:

`run_demo.bat`

The app opens at:

`http://127.0.0.1:8000`

The interactive lecture board opens at:

`http://127.0.0.1:8000/static/lecture.html`

The board includes ten presentation slides, presenter cues, fullscreen mode, pen, highlighter, eraser, undo and per-slide drawing memory.

## Manual startup

```powershell
python -m pip install -r requirements.txt
python -m uvicorn app:app --host 127.0.0.1 --port 8000
```

Then open `http://127.0.0.1:8000`.

## Suggested live-demo flow

1. Sign in as the Student.
2. Show the full-year roadmap and 138-lecture schedule.
3. Open “Trigonometry & Graph Interpretation.”
4. Select lesson text and save a note.
5. Complete the five-question daily test.
6. Show the score, explanations and error log.
7. Sign out.
8. Sign in as the Teacher.
9. Show the next-day review queue and weak-topic analytics.
10. Open the individual student progress page.

## Reset demo data

Stop the server and delete:

`physics_demo.db`

Restart the app. The database, demo users, lessons, questions and example activity will be recreated automatically.

## Local-only limitations

- The Student and Teacher must use the same computer/database.
- No public internet URL is created.
- Email delivery and password reset are not included.
- This is a single-student demo, not a classroom LMS.
- Objective questions are auto-marked; subjective CBSE marking is represented in the programme but not automated.

## Data files

- Application database: `physics_demo.db`
- Backend: `app.py`
- Browser UI: `static/`
- Programme source: `../syllabus/Class_11_Physics_Detailed_Lecture_Planner_Aug2026-Jan2027.csv`

