from __future__ import annotations

import csv
import hashlib
import hmac
import json
import secrets
import sqlite3
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException, Request, Response
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles


BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "physics_demo.db"
STATIC_DIR = BASE_DIR / "static"
SCHEDULE_CSV = BASE_DIR.parent / "syllabus" / "Class_11_Physics_Detailed_Lecture_Planner_Aug2026-Jan2027.csv"
IST = timezone(timedelta(hours=5, minutes=30))

app = FastAPI(title="Class 11 NEET Physics Study App", version="0.1.0")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


def db() -> sqlite3.Connection:
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


def now_iso() -> str:
    return datetime.now(IST).replace(microsecond=0).isoformat()


def hash_password(password: str, salt: str | None = None) -> str:
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode(), bytes.fromhex(salt), 120_000)
    return f"{salt}${digest.hex()}"


def verify_password(password: str, stored: str) -> bool:
    salt, expected = stored.split("$", 1)
    actual = hash_password(password, salt).split("$", 1)[1]
    return hmac.compare_digest(actual, expected)


def require_user(request: Request, role: str | None = None) -> sqlite3.Row:
    token = request.cookies.get("physics_session")
    if not token:
        raise HTTPException(401, "Please sign in")
    with db() as connection:
        user = connection.execute(
            """
            SELECT u.* FROM sessions s
            JOIN users u ON u.id = s.user_id
            WHERE s.token = ? AND s.expires_at > ?
            """,
            (token, now_iso()),
        ).fetchone()
    if not user:
        raise HTTPException(401, "Session expired")
    if role and user["role"] != role:
        raise HTTPException(403, "You do not have access to this area")
    return user


def lesson_content() -> list[dict[str, Any]]:
    return [
        {
            "lecture_no": 1,
            "chapter": "Mathematical Tools",
            "title": "Algebra, Ratios & Scientific Notation",
            "date": "2026-08-24",
            "summary": "Build the number skills used in every Physics calculation.",
            "minutes": 60,
            "sections": [
                {"type": "recall", "title": "Before you begin", "body": "What does 10³ mean? How would you compare 3:5 and 9:15?"},
                {"type": "common", "title": "Scientific notation", "body": "Write a non-zero number as a × 10ⁿ where 1 ≤ |a| < 10. Keep the coefficient in the standard range and move the decimal carefully.", "formula": "6,400,000 = 6.4 × 10⁶"},
                {"type": "neet", "title": "NEET speed habit", "body": "Combine powers of ten separately from coefficients. Estimate the order of magnitude before calculating to catch impossible answers."},
                {"type": "example", "title": "Worked example", "body": "(3 × 10⁴)(2 × 10⁻²) = 6 × 10². Multiply 3 × 2 and add the exponents 4 + (−2)."},
                {"type": "cbse", "title": "CBSE presentation", "body": "Show conversion steps and units. A correct final number without units is an incomplete Physics answer."},
            ],
        },
        {
            "lecture_no": 2,
            "chapter": "Mathematical Tools",
            "title": "Trigonometry & Graph Interpretation",
            "date": "2026-08-25",
            "summary": "Read slopes, components and standard trigonometric values.",
            "minutes": 60,
            "sections": [
                {"type": "recall", "title": "Recall strip", "body": "Convert 0.00045 to scientific notation. Simplify the ratio 12:18."},
                {"type": "common", "title": "Essential ratios", "body": "For a right triangle: sin θ = opposite/hypotenuse, cos θ = adjacent/hypotenuse and tan θ = opposite/adjacent.", "formula": "sin²θ + cos²θ = 1"},
                {"type": "common", "title": "Reading a graph", "body": "The slope is change in vertical quantity divided by change in horizontal quantity. Always include the units of both axes.", "formula": "slope = Δy / Δx"},
                {"type": "example", "title": "Physics connection", "body": "On a position–time graph the slope gives velocity. A steeper line means a greater magnitude of velocity."},
                {"type": "neet", "title": "NEET trap", "body": "Do not assume every upward graph means acceleration. First read the axis labels."},
            ],
        },
        {
            "lecture_no": 3,
            "chapter": "Mathematical Tools",
            "title": "Differentiation for Motion",
            "date": "2026-08-26",
            "summary": "Use rate of change to connect position, velocity and acceleration.",
            "minutes": 60,
            "sections": [
                {"type": "recall", "title": "Recall strip", "body": "State sin 30°, cos 60° and the meaning of graph slope."},
                {"type": "common", "title": "Derivative as rate", "body": "A derivative measures instantaneous rate of change. In motion, velocity is the derivative of position.", "formula": "v = dx/dt; a = dv/dt"},
                {"type": "example", "title": "Worked example", "body": "If x = 3t² + 2t, then v = 6t + 2 and a = 6."},
                {"type": "cbse", "title": "CBSE emphasis", "body": "Write the original function, differentiate term by term and state the units of the result."},
                {"type": "neet", "title": "NEET trap", "body": "At a turning point velocity may be zero while acceleration is not zero."},
            ],
        },
        {
            "lecture_no": 4,
            "chapter": "Mathematical Tools",
            "title": "Integration for Displacement & Work",
            "date": "2026-08-27",
            "summary": "Use accumulation and area under a graph in Physics.",
            "minutes": 60,
            "sections": [
                {"type": "recall", "title": "Recall strip", "body": "Differentiate x = 2t³. What does dx/dt represent?"},
                {"type": "common", "title": "Integration as accumulation", "body": "Integration reverses differentiation and adds many small contributions.", "formula": "Δx = ∫ v dt; W = ∫ F dx"},
                {"type": "example", "title": "Worked example", "body": "For v = 2t from t = 0 to 3 s, displacement is ∫₀³2t dt = 9 m."},
                {"type": "common", "title": "Graph meaning", "body": "Area under a velocity–time graph gives displacement. Signed area below the time axis is negative."},
                {"type": "neet", "title": "NEET trap", "body": "Distance and displacement are not always the same when velocity changes sign."},
            ],
        },
        {
            "lecture_no": 5,
            "chapter": "Units and Measurements",
            "title": "Physical Quantities, SI Units & Conversion",
            "date": "2026-08-28",
            "summary": "Use a consistent language of measurement.",
            "minutes": 60,
            "sections": [
                {"type": "recall", "title": "Recall strip", "body": "What does area under a velocity–time graph represent?"},
                {"type": "common", "title": "Base and derived quantities", "body": "SI has seven base quantities. Derived quantities are built from them by multiplication or division."},
                {"type": "definition", "title": "SI examples", "body": "Length: metre (m), mass: kilogram (kg), time: second (s), electric current: ampere (A), temperature: kelvin (K)."},
                {"type": "example", "title": "Conversion", "body": "72 km h⁻¹ = 72 × 1000/3600 = 20 m s⁻¹."},
                {"type": "cbse", "title": "Written-answer rule", "body": "Use standard symbols: m for metre, s for second and kg for kilogram. Unit symbols are not plural."},
            ],
        },
        {
            "lecture_no": 6,
            "chapter": "Units and Measurements",
            "title": "Significant Figures, Least Count & Errors",
            "date": "2026-08-29",
            "summary": "Report measurements honestly and combine uncertainty correctly.",
            "minutes": 60,
            "sections": [
                {"type": "recall", "title": "Recall strip", "body": "Convert 54 km h⁻¹ to m s⁻¹. Name three SI base units."},
                {"type": "common", "title": "Significant figures", "body": "Significant figures communicate measurement precision. Leading zeros are not significant; zeros between non-zero digits are significant."},
                {"type": "definition", "title": "Least count", "body": "The least count is the smallest value an instrument can reliably measure."},
                {"type": "common", "title": "Percentage error", "body": "Relative error is absolute error divided by measured value.", "formula": "percentage error = (Δx/x) × 100"},
                {"type": "neet", "title": "NEET trap", "body": "A calculator may show many digits, but the measurement does not automatically gain precision."},
            ],
        },
        {
            "lecture_no": 7,
            "chapter": "Units and Measurements",
            "title": "Dimensions & Dimensional Equations",
            "date": "2026-08-31",
            "summary": "Express physical quantities using mass, length and time.",
            "minutes": 60,
            "sections": [
                {"type": "recall", "title": "Recall strip", "body": "How many significant figures are in 0.00450? Define least count."},
                {"type": "common", "title": "Dimensional formula", "body": "Dimensions show how a quantity depends on base quantities. They are independent of the chosen unit system.", "formula": "[velocity] = [LT⁻¹]"},
                {"type": "example", "title": "Force", "body": "From F = ma, [F] = [M][LT⁻²] = [MLT⁻²]."},
                {"type": "example", "title": "Energy", "body": "Work = force × displacement, so [W] = [ML²T⁻²]."},
                {"type": "cbse", "title": "Presentation", "body": "Use square brackets for dimensions and write powers clearly."},
            ],
        },
        {
            "lecture_no": 8,
            "chapter": "Units and Measurements",
            "title": "Dimensional Analysis: Uses & Limits",
            "date": "2026-09-01",
            "summary": "Check equations and infer relations without overusing the method.",
            "minutes": 60,
            "sections": [
                {"type": "recall", "title": "Recall strip", "body": "Write the dimensions of force, velocity and energy."},
                {"type": "common", "title": "Principle of homogeneity", "body": "Every term added or subtracted in a valid physical equation must have the same dimensions."},
                {"type": "example", "title": "Checking an equation", "body": "In s = ut + ½at², both ut and at² have dimension L, matching displacement."},
                {"type": "neet", "title": "What analysis can do", "body": "It can test dimensional consistency, convert units and suggest power-law relations."},
                {"type": "neet", "title": "Limitation", "body": "It cannot determine dimensionless numerical constants such as 2, π or ½, and dimensional correctness does not guarantee physical correctness."},
            ],
        },
        {
            "lecture_no": 9,
            "chapter": "Units and Measurements",
            "title": "Vernier Calipers & Screw Gauge",
            "date": "2026-09-02",
            "summary": "Read common instruments and correct zero error.",
            "minutes": 60,
            "sections": [
                {"type": "recall", "title": "Recall strip", "body": "What does dimensional homogeneity test? Name one limitation."},
                {"type": "common", "title": "Vernier calipers", "body": "A Vernier scale improves precision by comparing divisions on a main scale and a sliding scale.", "formula": "reading = main scale + (coinciding division × least count)"},
                {"type": "common", "title": "Screw gauge", "body": "Pitch is the distance moved in one complete rotation.", "formula": "least count = pitch / circular-scale divisions"},
                {"type": "cbse", "title": "Practical focus", "body": "Record main scale, coinciding division, least count, zero correction and final reading separately."},
                {"type": "neet", "title": "NEET trap", "body": "Zero correction has the opposite sign to zero error."},
            ],
        },
    ]


QUESTION_BANK: dict[int, list[dict[str, Any]]] = {
    1: [
        {"prompt": "Write 0.00072 in scientific notation.", "options": ["7.2 × 10⁻⁴", "72 × 10⁻⁴", "7.2 × 10⁴", "0.72 × 10⁻³"], "answer": 0, "explanation": "Move the decimal four places right: 7.2 × 10⁻⁴.", "topic": "Scientific notation", "difficulty": "Easy", "tag": "Common"},
        {"prompt": "(2 × 10³)(4 × 10⁻²) equals:", "options": ["8 × 10¹", "8 × 10⁵", "6 × 10¹", "8 × 10⁻⁶"], "answer": 0, "explanation": "Multiply coefficients and add exponents: 8 × 10¹.", "topic": "Powers", "difficulty": "Easy", "tag": "NEET"},
        {"prompt": "The ratio 18:24 in simplest form is:", "options": ["9:12", "3:4", "2:3", "4:3"], "answer": 1, "explanation": "Divide both terms by 6.", "topic": "Ratios", "difficulty": "Easy", "tag": "Common"},
        {"prompt": "The order of magnitude of 6.2 × 10⁵ is:", "options": ["10⁴", "10⁵", "10⁶", "6.2"], "answer": 2, "explanation": "Because 6.2 is above √10, the nearest power of ten is 10⁶.", "topic": "Order of magnitude", "difficulty": "Medium", "tag": "NEET"},
        {"prompt": "If x/5 = 6, x is:", "options": ["1.2", "11", "30", "25"], "answer": 2, "explanation": "Multiply both sides by 5.", "topic": "Algebra", "difficulty": "Easy", "tag": "Common"},
    ],
    2: [
        {"prompt": "sin 30° equals:", "options": ["0", "1/2", "1/√2", "1"], "answer": 1, "explanation": "sin 30° = 1/2.", "topic": "Trigonometry", "difficulty": "Easy", "tag": "Common"},
        {"prompt": "tan 45° equals:", "options": ["0", "1/2", "1", "√3"], "answer": 2, "explanation": "tan 45° = 1.", "topic": "Trigonometry", "difficulty": "Easy", "tag": "Common"},
        {"prompt": "The slope of a position–time graph represents:", "options": ["Acceleration", "Velocity", "Force", "Distance"], "answer": 1, "explanation": "Slope Δx/Δt is velocity.", "topic": "Graphs", "difficulty": "Easy", "tag": "CBSE"},
        {"prompt": "A horizontal line on a position–time graph means:", "options": ["Constant acceleration", "Object at rest", "Uniform speed", "Negative velocity"], "answer": 1, "explanation": "Position is unchanged, so velocity is zero.", "topic": "Graphs", "difficulty": "Medium", "tag": "NEET"},
        {"prompt": "Which identity is correct?", "options": ["sin θ + cos θ = 1", "sin²θ + cos²θ = 1", "tan θ = sin θ cos θ", "cos²θ − sin²θ = 1 always"], "answer": 1, "explanation": "The fundamental Pythagorean identity is sin²θ + cos²θ = 1.", "topic": "Trigonometry", "difficulty": "Easy", "tag": "Common"},
    ],
    3: [
        {"prompt": "If x = 4t², velocity is:", "options": ["4t", "8t", "8", "2t"], "answer": 1, "explanation": "dx/dt = 8t.", "topic": "Differentiation", "difficulty": "Easy", "tag": "Common"},
        {"prompt": "If v = 3t² + 2, acceleration is:", "options": ["3t", "6t", "6", "3t²"], "answer": 1, "explanation": "dv/dt = 6t.", "topic": "Differentiation", "difficulty": "Easy", "tag": "Common"},
        {"prompt": "The derivative of a constant is:", "options": ["1", "The constant", "0", "Undefined"], "answer": 2, "explanation": "A constant does not change.", "topic": "Differentiation", "difficulty": "Easy", "tag": "Common"},
        {"prompt": "Instantaneous velocity is geometrically the:", "options": ["Area under x–t graph", "Slope of tangent to x–t graph", "Slope of v–t graph", "Area under a–t graph"], "answer": 1, "explanation": "Instantaneous rate is the tangent slope.", "topic": "Graphs", "difficulty": "Medium", "tag": "NEET"},
        {"prompt": "At a turning point, which can be true?", "options": ["Velocity and acceleration must both be zero", "Velocity is zero but acceleration is non-zero", "Acceleration must be zero", "Position must be zero"], "answer": 1, "explanation": "A body can reverse direction with non-zero acceleration.", "topic": "Motion", "difficulty": "Medium", "tag": "NEET"},
    ],
    4: [
        {"prompt": "If v = 2t, displacement from 0 to 2 s is:", "options": ["2 m", "4 m", "6 m", "8 m"], "answer": 1, "explanation": "∫₀²2t dt = [t²]₀² = 4 m.", "topic": "Integration", "difficulty": "Easy", "tag": "Common"},
        {"prompt": "Area under a velocity–time graph gives:", "options": ["Acceleration", "Displacement", "Force", "Power"], "answer": 1, "explanation": "Integrating velocity over time gives displacement.", "topic": "Graphs", "difficulty": "Easy", "tag": "CBSE"},
        {"prompt": "Work done by a variable force is:", "options": ["∫F dx", "dF/dx", "F/x", "∫x dF only"], "answer": 0, "explanation": "Work is the area under the force–displacement graph.", "topic": "Integration", "difficulty": "Easy", "tag": "NEET"},
        {"prompt": "Signed area below the time axis on a v–t graph is:", "options": ["Always distance", "Negative displacement", "Positive acceleration", "Zero speed"], "answer": 1, "explanation": "Negative velocity contributes negative displacement.", "topic": "Graphs", "difficulty": "Medium", "tag": "NEET"},
        {"prompt": "Integration is most closely related to:", "options": ["Instantaneous rate only", "Accumulation", "Unit conversion", "Rounding"], "answer": 1, "explanation": "Integration sums many small contributions.", "topic": "Integration", "difficulty": "Easy", "tag": "Common"},
    ],
    5: [
        {"prompt": "The SI base unit of mass is:", "options": ["gram", "kilogram", "tonne", "newton"], "answer": 1, "explanation": "Kilogram is the SI base unit of mass.", "topic": "SI units", "difficulty": "Easy", "tag": "Common"},
        {"prompt": "72 km h⁻¹ equals:", "options": ["10 m s⁻¹", "20 m s⁻¹", "36 m s⁻¹", "72 m s⁻¹"], "answer": 1, "explanation": "Multiply by 5/18: 72 × 5/18 = 20.", "topic": "Conversion", "difficulty": "Easy", "tag": "NEET"},
        {"prompt": "Which is a derived SI unit?", "options": ["metre", "second", "newton", "kelvin"], "answer": 2, "explanation": "Newton is kg m s⁻².", "topic": "SI units", "difficulty": "Easy", "tag": "Common"},
        {"prompt": "The correct symbol for seconds is:", "options": ["sec", "S", "s", "secs"], "answer": 2, "explanation": "The SI symbol is lowercase s.", "topic": "Notation", "difficulty": "Easy", "tag": "CBSE"},
        {"prompt": "54 km h⁻¹ equals:", "options": ["12 m s⁻¹", "15 m s⁻¹", "18 m s⁻¹", "20 m s⁻¹"], "answer": 1, "explanation": "54 × 5/18 = 15 m s⁻¹.", "topic": "Conversion", "difficulty": "Easy", "tag": "NEET"},
    ],
    6: [
        {"prompt": "How many significant figures are in 0.00450?", "options": ["2", "3", "4", "5"], "answer": 1, "explanation": "Leading zeros do not count; 4, 5 and the trailing zero count.", "topic": "Significant figures", "difficulty": "Easy", "tag": "Common"},
        {"prompt": "Least count means:", "options": ["Largest reading", "Smallest measurable value", "Average error", "Zero correction"], "answer": 1, "explanation": "It is the smallest instrument division or reliable increment.", "topic": "Least count", "difficulty": "Easy", "tag": "Common"},
        {"prompt": "For x = 50 ± 1, percentage error is:", "options": ["1%", "2%", "5%", "50%"], "answer": 1, "explanation": "(1/50) × 100 = 2%.", "topic": "Errors", "difficulty": "Easy", "tag": "NEET"},
        {"prompt": "A calculator shows 2.345678 for a two-significant-figure measurement. Report:", "options": ["2.345678", "2.35", "2.3", "2"], "answer": 2, "explanation": "2.345678 rounds to 2.3 at two significant figures.", "topic": "Rounding", "difficulty": "Medium", "tag": "CBSE"},
        {"prompt": "Zero correction is:", "options": ["Same sign as zero error", "Opposite sign to zero error", "Always positive", "Always zero"], "answer": 1, "explanation": "Correction removes the error and therefore has the opposite sign.", "topic": "Errors", "difficulty": "Easy", "tag": "NEET"},
    ],
    7: [
        {"prompt": "Dimensions of velocity are:", "options": ["[LT⁻¹]", "[LT⁻²]", "[MLT⁻¹]", "[L²T⁻¹]"], "answer": 0, "explanation": "Velocity is displacement divided by time.", "topic": "Dimensions", "difficulty": "Easy", "tag": "Common"},
        {"prompt": "Dimensions of force are:", "options": ["[MLT⁻¹]", "[MLT⁻²]", "[ML²T⁻²]", "[LT⁻²]"], "answer": 1, "explanation": "Force = mass × acceleration.", "topic": "Dimensions", "difficulty": "Easy", "tag": "Common"},
        {"prompt": "Dimensions of energy are:", "options": ["[MLT⁻²]", "[ML²T⁻²]", "[ML²T⁻¹]", "[M⁰LT⁻¹]"], "answer": 1, "explanation": "Energy = force × displacement.", "topic": "Dimensions", "difficulty": "Easy", "tag": "Common"},
        {"prompt": "Dimensions of density are:", "options": ["[ML⁻³]", "[M⁻¹L³]", "[ML³]", "[L⁻³]"], "answer": 0, "explanation": "Density is mass/volume.", "topic": "Dimensions", "difficulty": "Medium", "tag": "NEET"},
        {"prompt": "Dimensions depend on the chosen unit system:", "options": ["Always true", "False", "Only in SI", "Only for derived quantities"], "answer": 1, "explanation": "Dimensions describe physical dependence, not the unit system.", "topic": "Dimensions", "difficulty": "Easy", "tag": "CBSE"},
    ],
    8: [
        {"prompt": "In a valid equation, added terms must have:", "options": ["Same numerical value", "Same dimensions", "Same sign", "No units"], "answer": 1, "explanation": "This is the principle of dimensional homogeneity.", "topic": "Homogeneity", "difficulty": "Easy", "tag": "Common"},
        {"prompt": "Dimensional analysis cannot determine:", "options": ["Power-law exponents", "Unit conversion", "Dimensionless numerical constants", "Dimensional consistency"], "answer": 2, "explanation": "Constants such as 2 and π cannot be found.", "topic": "Limitations", "difficulty": "Easy", "tag": "NEET"},
        {"prompt": "Is s = ut + ½at² dimensionally consistent?", "options": ["Yes", "No", "Only if a = 0", "Only in SI"], "answer": 0, "explanation": "Both terms on the right have dimension L.", "topic": "Homogeneity", "difficulty": "Medium", "tag": "CBSE"},
        {"prompt": "A dimensionally correct equation is:", "options": ["Always physically correct", "Never physically correct", "Not necessarily physically correct", "Correct only in SI"], "answer": 2, "explanation": "Dimensional consistency is necessary but not sufficient.", "topic": "Limitations", "difficulty": "Medium", "tag": "NEET"},
        {"prompt": "Dimensional analysis can help with:", "options": ["Deriving exact trigonometric dependence", "Checking an equation", "Finding additive constants", "Proving a physical law completely"], "answer": 1, "explanation": "Checking consistency is a core use.", "topic": "Applications", "difficulty": "Easy", "tag": "Common"},
    ],
    9: [
        {"prompt": "Vernier reading is main scale plus:", "options": ["Pitch × divisions", "Coinciding division × least count", "Zero error only", "Main scale/least count"], "answer": 1, "explanation": "Add the Vernier contribution to the main-scale reading.", "topic": "Vernier calipers", "difficulty": "Easy", "tag": "Practical"},
        {"prompt": "Screw gauge least count is:", "options": ["Pitch × divisions", "Pitch/divisions", "Divisions/pitch", "Pitch + divisions"], "answer": 1, "explanation": "Least count = pitch divided by circular-scale divisions.", "topic": "Screw gauge", "difficulty": "Easy", "tag": "Practical"},
        {"prompt": "For a positive zero error, zero correction is:", "options": ["Positive", "Negative", "Zero", "Undefined"], "answer": 1, "explanation": "Correction is opposite in sign to the error.", "topic": "Zero error", "difficulty": "Easy", "tag": "NEET"},
        {"prompt": "Pitch of a screw gauge is the distance moved in:", "options": ["One scale division", "One complete rotation", "Half rotation", "Ten rotations"], "answer": 1, "explanation": "Pitch is linear movement per complete rotation.", "topic": "Screw gauge", "difficulty": "Easy", "tag": "Practical"},
        {"prompt": "Which should be recorded separately in a practical?", "options": ["Only final reading", "Main scale, coincidence, least count and correction", "Only zero error", "Only instrument name"], "answer": 1, "explanation": "Separate observations make the result verifiable.", "topic": "Measurement", "difficulty": "Medium", "tag": "CBSE"},
    ],
}


def create_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('student','teacher'))
        );
        CREATE TABLE IF NOT EXISTS sessions (
            token TEXT PRIMARY KEY,
            user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            expires_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS chapters (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL UNIQUE,
            sequence INTEGER NOT NULL,
            status TEXT NOT NULL DEFAULT 'locked'
        );
        CREATE TABLE IF NOT EXISTS lessons (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            lecture_no INTEGER UNIQUE,
            chapter_id INTEGER NOT NULL REFERENCES chapters(id),
            title TEXT NOT NULL,
            lesson_date TEXT,
            sequence INTEGER NOT NULL,
            summary TEXT NOT NULL,
            minutes INTEGER NOT NULL DEFAULT 60,
            sections_json TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS questions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            lesson_id INTEGER NOT NULL REFERENCES lessons(id) ON DELETE CASCADE,
            prompt TEXT NOT NULL,
            options_json TEXT NOT NULL,
            correct_index INTEGER NOT NULL,
            explanation TEXT NOT NULL,
            topic TEXT NOT NULL,
            difficulty TEXT NOT NULL,
            exam_tag TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS progress (
            user_id INTEGER NOT NULL REFERENCES users(id),
            lesson_id INTEGER NOT NULL REFERENCES lessons(id),
            status TEXT NOT NULL DEFAULT 'not_started',
            mastery REAL NOT NULL DEFAULT 0,
            last_opened TEXT,
            PRIMARY KEY(user_id, lesson_id)
        );
        CREATE TABLE IF NOT EXISTS attempts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL REFERENCES users(id),
            lesson_id INTEGER NOT NULL REFERENCES lessons(id),
            score INTEGER NOT NULL,
            total INTEGER NOT NULL,
            duration_seconds INTEGER NOT NULL DEFAULT 0,
            submitted_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS answers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            attempt_id INTEGER NOT NULL REFERENCES attempts(id) ON DELETE CASCADE,
            question_id INTEGER NOT NULL REFERENCES questions(id),
            selected_index INTEGER,
            is_correct INTEGER NOT NULL,
            retry_resolved INTEGER NOT NULL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS notes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL REFERENCES users(id),
            lesson_id INTEGER NOT NULL REFERENCES lessons(id),
            kind TEXT NOT NULL,
            selected_text TEXT,
            note TEXT NOT NULL,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS schedule (
            lecture_no INTEGER PRIMARY KEY,
            lecture_date TEXT NOT NULL,
            day TEXT NOT NULL,
            phase TEXT NOT NULL,
            chapter TEXT NOT NULL,
            topic TEXT NOT NULL,
            focus TEXT,
            homework TEXT
        );
        CREATE TABLE IF NOT EXISTS teacher_reviews (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            teacher_id INTEGER NOT NULL REFERENCES users(id),
            question_id INTEGER NOT NULL REFERENCES questions(id),
            reviewed_at TEXT NOT NULL,
            UNIQUE(teacher_id, question_id)
        );
        """
    )


def seed_database() -> None:
    with db() as connection:
        create_schema(connection)
        if connection.execute("SELECT COUNT(*) FROM users").fetchone()[0] == 0:
            connection.executemany(
                "INSERT INTO users(name,email,password_hash,role) VALUES(?,?,?,?)",
                [
                    ("Aarav Sharma", "student@physics.local", hash_password("student123"), "student"),
                    ("Dr. Meera Singh", "teacher@physics.local", hash_password("teacher123"), "teacher"),
                ],
            )

        chapter_names = [
            "Mathematical Tools",
            "Units and Measurements",
            "Motion in a Straight Line",
            "Motion in a Plane",
            "Laws of Motion",
            "Work, Energy and Power",
            "System of Particles and Rotational Motion",
            "Gravitation",
            "Mechanical Properties of Solids",
            "Mechanical Properties of Fluids",
            "Thermal Properties of Matter",
            "Thermodynamics",
            "Kinetic Theory",
            "Oscillations",
            "Waves",
        ]
        for index, title in enumerate(chapter_names, start=1):
            connection.execute(
                "INSERT OR IGNORE INTO chapters(title,sequence,status) VALUES(?,?,?)",
                (title, index, "active" if index <= 2 else "locked"),
            )

        if connection.execute("SELECT COUNT(*) FROM lessons").fetchone()[0] == 0:
            chapter_ids = {row["title"]: row["id"] for row in connection.execute("SELECT id,title FROM chapters")}
            for sequence, lesson in enumerate(lesson_content(), start=1):
                cursor = connection.execute(
                    """
                    INSERT INTO lessons(lecture_no,chapter_id,title,lesson_date,sequence,summary,minutes,sections_json)
                    VALUES(?,?,?,?,?,?,?,?)
                    """,
                    (
                        lesson["lecture_no"],
                        chapter_ids[lesson["chapter"]],
                        lesson["title"],
                        lesson["date"],
                        sequence,
                        lesson["summary"],
                        lesson["minutes"],
                        json.dumps(lesson["sections"]),
                    ),
                )
                lesson_id = cursor.lastrowid
                for question in QUESTION_BANK[lesson["lecture_no"]]:
                    connection.execute(
                        """
                        INSERT INTO questions(lesson_id,prompt,options_json,correct_index,explanation,topic,difficulty,exam_tag)
                        VALUES(?,?,?,?,?,?,?,?)
                        """,
                        (
                            lesson_id,
                            question["prompt"],
                            json.dumps(question["options"]),
                            question["answer"],
                            question["explanation"],
                            question["topic"],
                            question["difficulty"],
                            question["tag"],
                        ),
                    )

        if SCHEDULE_CSV.exists() and connection.execute("SELECT COUNT(*) FROM schedule").fetchone()[0] == 0:
            with SCHEDULE_CSV.open("r", encoding="utf-8-sig", newline="") as handle:
                for row in csv.DictReader(handle):
                    connection.execute(
                        """
                        INSERT OR IGNORE INTO schedule(lecture_no,lecture_date,day,phase,chapter,topic,focus,homework)
                        VALUES(?,?,?,?,?,?,?,?)
                        """,
                        (
                            int(row["Lecture"]),
                            row["Date"],
                            row["Day"],
                            row["Phase"],
                            row["Chapter"],
                            row["Topic"],
                            row["Delivery Focus"],
                            row["Assessment or Homework"],
                        ),
                    )

        student = connection.execute("SELECT id FROM users WHERE role='student'").fetchone()
        lesson_one = connection.execute("SELECT id FROM lessons WHERE lecture_no=1").fetchone()
        lesson_two = connection.execute("SELECT id FROM lessons WHERE lecture_no=2").fetchone()
        if student and connection.execute("SELECT COUNT(*) FROM progress WHERE user_id=?", (student["id"],)).fetchone()[0] == 0:
            connection.execute(
                "INSERT INTO progress(user_id,lesson_id,status,mastery,last_opened) VALUES(?,?,?,?,?)",
                (student["id"], lesson_one["id"], "mastered", 80, now_iso()),
            )
            connection.execute(
                "INSERT INTO progress(user_id,lesson_id,status,mastery,last_opened) VALUES(?,?,?,?,?)",
                (student["id"], lesson_two["id"], "reading", 20, now_iso()),
            )
            questions = connection.execute("SELECT * FROM questions WHERE lesson_id=? ORDER BY id", (lesson_one["id"],)).fetchall()
            cursor = connection.execute(
                "INSERT INTO attempts(user_id,lesson_id,score,total,duration_seconds,submitted_at) VALUES(?,?,?,?,?,?)",
                (student["id"], lesson_one["id"], 4, 5, 338, (datetime.now(IST) - timedelta(days=1)).replace(microsecond=0).isoformat()),
            )
            attempt_id = cursor.lastrowid
            for index, question in enumerate(questions):
                is_correct = 0 if index == 3 else 1
                selected = question["correct_index"] if is_correct else 1
                connection.execute(
                    "INSERT INTO answers(attempt_id,question_id,selected_index,is_correct) VALUES(?,?,?,?)",
                    (attempt_id, question["id"], selected, is_correct),
                )

        connection.commit()


seed_database()


@app.get("/")
def index() -> FileResponse:
    return FileResponse(STATIC_DIR / "index.html")


@app.post("/api/login")
async def login(request: Request, response: Response) -> dict[str, Any]:
    payload = await request.json()
    email = str(payload.get("email", "")).strip().lower()
    password = str(payload.get("password", ""))
    with db() as connection:
        user = connection.execute("SELECT * FROM users WHERE email=?", (email,)).fetchone()
        if not user or not verify_password(password, user["password_hash"]):
            raise HTTPException(401, "Incorrect email or password")
        token = secrets.token_urlsafe(32)
        expires = (datetime.now(IST) + timedelta(days=7)).replace(microsecond=0).isoformat()
        connection.execute("INSERT INTO sessions(token,user_id,expires_at) VALUES(?,?,?)", (token, user["id"], expires))
        connection.commit()
    response.set_cookie("physics_session", token, httponly=True, samesite="lax", max_age=7 * 24 * 3600)
    return {"user": {"id": user["id"], "name": user["name"], "email": user["email"], "role": user["role"]}}


@app.post("/api/logout")
def logout(request: Request, response: Response) -> dict[str, bool]:
    token = request.cookies.get("physics_session")
    if token:
        with db() as connection:
            connection.execute("DELETE FROM sessions WHERE token=?", (token,))
            connection.commit()
    response.delete_cookie("physics_session")
    return {"ok": True}


@app.get("/api/me")
def me(request: Request) -> dict[str, Any]:
    user = require_user(request)
    return {"id": user["id"], "name": user["name"], "email": user["email"], "role": user["role"]}


@app.get("/api/roadmap")
def roadmap(request: Request) -> dict[str, Any]:
    user = require_user(request)
    with db() as connection:
        rows = connection.execute(
            """
            SELECT c.id,c.title,c.sequence,c.status,
                   COUNT(l.id) lesson_count,
                   SUM(CASE WHEN p.status='mastered' THEN 1 ELSE 0 END) mastered_count
            FROM chapters c
            LEFT JOIN lessons l ON l.chapter_id=c.id
            LEFT JOIN progress p ON p.lesson_id=l.id AND p.user_id=?
            GROUP BY c.id ORDER BY c.sequence
            """,
            (user["id"],),
        ).fetchall()
    return {"chapters": [dict(row) for row in rows]}


@app.get("/api/student/dashboard")
def student_dashboard(request: Request) -> dict[str, Any]:
    user = require_user(request, "student")
    with db() as connection:
        lesson_rows = connection.execute(
            """
            SELECT l.id,l.lecture_no,l.title,l.lesson_date,l.summary,c.title chapter,
                   COALESCE(p.status,'not_started') status,COALESCE(p.mastery,0) mastery
            FROM lessons l JOIN chapters c ON c.id=l.chapter_id
            LEFT JOIN progress p ON p.lesson_id=l.id AND p.user_id=?
            ORDER BY l.sequence
            """,
            (user["id"],),
        ).fetchall()
        attempts = connection.execute(
            """
            SELECT a.id,a.score,a.total,a.submitted_at,l.title,c.title chapter
            FROM attempts a JOIN lessons l ON l.id=a.lesson_id JOIN chapters c ON c.id=l.chapter_id
            WHERE a.user_id=? ORDER BY a.id DESC LIMIT 5
            """,
            (user["id"],),
        ).fetchall()
        stats = connection.execute(
            """
            SELECT COUNT(*) attempt_count,
                   COALESCE(ROUND(AVG(score*100.0/total)),0) average_score
            FROM attempts WHERE user_id=?
            """,
            (user["id"],),
        ).fetchone()
        unresolved = connection.execute(
            """
            SELECT COUNT(*) FROM answers an
            JOIN attempts a ON a.id=an.attempt_id
            WHERE a.user_id=? AND an.is_correct=0 AND an.retry_resolved=0
            """,
            (user["id"],),
        ).fetchone()[0]
        upcoming = connection.execute(
            "SELECT * FROM schedule ORDER BY lecture_no LIMIT 8"
        ).fetchall()
    next_lesson = next((dict(row) for row in lesson_rows if row["status"] != "mastered"), dict(lesson_rows[-1]))
    return {
        "user": {"name": user["name"]},
        "stats": {
            "lessonsMastered": sum(1 for row in lesson_rows if row["status"] == "mastered"),
            "totalAvailable": len(lesson_rows),
            "averageScore": stats["average_score"],
            "unresolvedErrors": unresolved,
            "streak": 4,
        },
        "nextLesson": next_lesson,
        "lessons": [dict(row) for row in lesson_rows],
        "attempts": [dict(row) for row in attempts],
        "upcoming": [dict(row) for row in upcoming],
    }


@app.get("/api/lessons/{lesson_id}")
def get_lesson(lesson_id: int, request: Request) -> dict[str, Any]:
    user = require_user(request)
    with db() as connection:
        lesson = connection.execute(
            """
            SELECT l.*,c.title chapter FROM lessons l
            JOIN chapters c ON c.id=l.chapter_id WHERE l.id=?
            """,
            (lesson_id,),
        ).fetchone()
        if not lesson:
            raise HTTPException(404, "Lesson not found")
        if user["role"] == "student":
            connection.execute(
                """
                INSERT INTO progress(user_id,lesson_id,status,mastery,last_opened)
                VALUES(?,?,?,0,?)
                ON CONFLICT(user_id,lesson_id) DO UPDATE SET
                    status=CASE WHEN progress.status='not_started' THEN 'reading' ELSE progress.status END,
                    last_opened=excluded.last_opened
                """,
                (user["id"], lesson_id, "reading", now_iso()),
            )
            connection.commit()
        progress = connection.execute(
            "SELECT * FROM progress WHERE user_id=? AND lesson_id=?",
            (user["id"], lesson_id),
        ).fetchone()
        notes = connection.execute(
            "SELECT * FROM notes WHERE user_id=? AND lesson_id=? ORDER BY id DESC",
            (user["id"], lesson_id),
        ).fetchall()
    result = dict(lesson)
    result["sections"] = json.loads(result.pop("sections_json"))
    result["progress"] = dict(progress) if progress else {"status": "not_started", "mastery": 0}
    result["notes"] = [dict(note) for note in notes]
    return result


@app.post("/api/lessons/{lesson_id}/notes")
async def add_note(lesson_id: int, request: Request) -> dict[str, Any]:
    user = require_user(request, "student")
    payload = await request.json()
    kind = str(payload.get("kind", "Key concept"))[:40]
    note = str(payload.get("note", "")).strip()
    selected = str(payload.get("selectedText", "")).strip()
    if not note:
        raise HTTPException(400, "Note cannot be empty")
    with db() as connection:
        cursor = connection.execute(
            "INSERT INTO notes(user_id,lesson_id,kind,selected_text,note,created_at) VALUES(?,?,?,?,?,?)",
            (user["id"], lesson_id, kind, selected, note[:1000], now_iso()),
        )
        connection.commit()
        row = connection.execute("SELECT * FROM notes WHERE id=?", (cursor.lastrowid,)).fetchone()
    return dict(row)


@app.get("/api/lessons/{lesson_id}/quiz")
def get_quiz(lesson_id: int, request: Request) -> dict[str, Any]:
    require_user(request)
    with db() as connection:
        lesson = connection.execute("SELECT id,title FROM lessons WHERE id=?", (lesson_id,)).fetchone()
        if not lesson:
            raise HTTPException(404, "Lesson not found")
        questions = connection.execute(
            "SELECT id,prompt,options_json,topic,difficulty,exam_tag FROM questions WHERE lesson_id=? ORDER BY id",
            (lesson_id,),
        ).fetchall()
    return {
        "lesson": dict(lesson),
        "questions": [
            {
                "id": row["id"],
                "prompt": row["prompt"],
                "options": json.loads(row["options_json"]),
                "topic": row["topic"],
                "difficulty": row["difficulty"],
                "examTag": row["exam_tag"],
            }
            for row in questions
        ],
    }


@app.post("/api/lessons/{lesson_id}/quiz")
async def submit_quiz(lesson_id: int, request: Request) -> dict[str, Any]:
    user = require_user(request, "student")
    payload = await request.json()
    submitted_answers = payload.get("answers", {})
    duration = max(0, min(int(payload.get("durationSeconds", 0)), 7200))
    with db() as connection:
        questions = connection.execute("SELECT * FROM questions WHERE lesson_id=? ORDER BY id", (lesson_id,)).fetchall()
        if not questions:
            raise HTTPException(404, "Quiz not found")
        score = 0
        evaluated = []
        for question in questions:
            raw = submitted_answers.get(str(question["id"]))
            selected = int(raw) if raw is not None else None
            correct = selected == question["correct_index"]
            score += int(correct)
            evaluated.append((question, selected, correct))
        cursor = connection.execute(
            "INSERT INTO attempts(user_id,lesson_id,score,total,duration_seconds,submitted_at) VALUES(?,?,?,?,?,?)",
            (user["id"], lesson_id, score, len(questions), duration, now_iso()),
        )
        attempt_id = cursor.lastrowid
        for question, selected, correct in evaluated:
            connection.execute(
                "INSERT INTO answers(attempt_id,question_id,selected_index,is_correct) VALUES(?,?,?,?)",
                (attempt_id, question["id"], selected, int(correct)),
            )
        percentage = round(score * 100 / len(questions))
        status = "mastered" if percentage >= 70 else "practiced"
        connection.execute(
            """
            INSERT INTO progress(user_id,lesson_id,status,mastery,last_opened)
            VALUES(?,?,?,?,?)
            ON CONFLICT(user_id,lesson_id) DO UPDATE SET
                status=excluded.status,mastery=MAX(progress.mastery,excluded.mastery),last_opened=excluded.last_opened
            """,
            (user["id"], lesson_id, status, percentage, now_iso()),
        )
        connection.commit()
    return {
        "attemptId": attempt_id,
        "score": score,
        "total": len(questions),
        "percentage": percentage,
        "status": status,
        "needsRemediation": percentage < 70,
        "retryDue": (datetime.now(IST) + timedelta(hours=48)).replace(microsecond=0).isoformat(),
        "results": [
            {
                "questionId": question["id"],
                "selectedIndex": selected,
                "correctIndex": question["correct_index"],
                "isCorrect": correct,
                "explanation": question["explanation"],
            }
            for question, selected, correct in evaluated
        ],
    }


@app.get("/api/student/error-log")
def error_log(request: Request) -> dict[str, Any]:
    user = require_user(request, "student")
    with db() as connection:
        rows = connection.execute(
            """
            SELECT an.id answer_id,q.prompt,q.topic,q.explanation,q.options_json,q.correct_index,
                   an.selected_index,an.retry_resolved,a.submitted_at,l.title lesson
            FROM answers an
            JOIN attempts a ON a.id=an.attempt_id
            JOIN questions q ON q.id=an.question_id
            JOIN lessons l ON l.id=q.lesson_id
            WHERE a.user_id=? AND an.is_correct=0
            ORDER BY an.id DESC
            """,
            (user["id"],),
        ).fetchall()
    errors = []
    for row in rows:
        item = dict(row)
        item["options"] = json.loads(item.pop("options_json"))
        errors.append(item)
    return {"errors": errors}


@app.get("/api/student/notebook")
def notebook(request: Request) -> dict[str, Any]:
    user = require_user(request, "student")
    with db() as connection:
        rows = connection.execute(
            """
            SELECT n.*,l.title lesson,c.title chapter FROM notes n
            JOIN lessons l ON l.id=n.lesson_id JOIN chapters c ON c.id=l.chapter_id
            WHERE n.user_id=? ORDER BY n.id DESC
            """,
            (user["id"],),
        ).fetchall()
    return {"notes": [dict(row) for row in rows]}


@app.get("/api/teacher/dashboard")
def teacher_dashboard(request: Request) -> dict[str, Any]:
    require_user(request, "teacher")
    with db() as connection:
        students = connection.execute(
            """
            SELECT u.id,u.name,u.email,
                   COUNT(DISTINCT a.id) attempts,
                   COALESCE(ROUND(AVG(a.score*100.0/a.total)),0) average_score,
                   MAX(a.submitted_at) last_active,
                   COUNT(DISTINCT CASE WHEN p.status='mastered' THEN p.lesson_id END) mastered
            FROM users u
            LEFT JOIN attempts a ON a.user_id=u.id
            LEFT JOIN progress p ON p.user_id=u.id
            WHERE u.role='student' GROUP BY u.id
            """
        ).fetchall()
        weak_topics = connection.execute(
            """
            SELECT q.topic,COUNT(*) attempts,
                   SUM(CASE WHEN an.is_correct=0 THEN 1 ELSE 0 END) wrong,
                   ROUND(100.0*SUM(CASE WHEN an.is_correct=1 THEN 1 ELSE 0 END)/COUNT(*)) accuracy
            FROM answers an JOIN questions q ON q.id=an.question_id
            GROUP BY q.topic ORDER BY accuracy ASC,attempts DESC LIMIT 8
            """
        ).fetchall()
        review_queue = connection.execute(
            """
            SELECT q.id question_id,q.prompt,q.topic,l.title lesson,
                   COUNT(*) attempts,
                   SUM(CASE WHEN an.is_correct=0 THEN 1 ELSE 0 END) wrong,
                   ROUND(100.0*SUM(CASE WHEN an.is_correct=0 THEN 1 ELSE 0 END)/COUNT(*)) miss_rate,
                   CASE WHEN tr.id IS NULL THEN 0 ELSE 1 END reviewed
            FROM answers an
            JOIN questions q ON q.id=an.question_id
            JOIN lessons l ON l.id=q.lesson_id
            LEFT JOIN teacher_reviews tr ON tr.question_id=q.id
            GROUP BY q.id ORDER BY reviewed ASC,miss_rate DESC,attempts DESC LIMIT 10
            """
        ).fetchall()
        recent = connection.execute(
            """
            SELECT a.id,u.name student,l.title lesson,a.score,a.total,a.submitted_at
            FROM attempts a JOIN users u ON u.id=a.user_id JOIN lessons l ON l.id=a.lesson_id
            ORDER BY a.id DESC LIMIT 8
            """
        ).fetchall()
        notes_summary = connection.execute(
            """
            SELECT u.name,COUNT(n.id) note_count,MAX(n.created_at) latest
            FROM users u LEFT JOIN notes n ON n.user_id=u.id
            WHERE u.role='student' GROUP BY u.id
            """
        ).fetchall()
    return {
        "students": [dict(row) for row in students],
        "weakTopics": [dict(row) for row in weak_topics],
        "reviewQueue": [dict(row) for row in review_queue],
        "recentAttempts": [dict(row) for row in recent],
        "notesSummary": [dict(row) for row in notes_summary],
    }


@app.post("/api/teacher/review/{question_id}")
def mark_reviewed(question_id: int, request: Request) -> dict[str, bool]:
    teacher = require_user(request, "teacher")
    with db() as connection:
        connection.execute(
            "INSERT OR REPLACE INTO teacher_reviews(teacher_id,question_id,reviewed_at) VALUES(?,?,?)",
            (teacher["id"], question_id, now_iso()),
        )
        connection.commit()
    return {"ok": True}


@app.get("/api/teacher/student/{student_id}")
def teacher_student(student_id: int, request: Request) -> dict[str, Any]:
    require_user(request, "teacher")
    with db() as connection:
        student = connection.execute("SELECT id,name,email FROM users WHERE id=? AND role='student'", (student_id,)).fetchone()
        if not student:
            raise HTTPException(404, "Student not found")
        attempts = connection.execute(
            """
            SELECT a.*,l.title lesson,c.title chapter FROM attempts a
            JOIN lessons l ON l.id=a.lesson_id JOIN chapters c ON c.id=l.chapter_id
            WHERE a.user_id=? ORDER BY a.id DESC
            """,
            (student_id,),
        ).fetchall()
        topics = connection.execute(
            """
            SELECT q.topic,COUNT(*) attempts,
                   ROUND(100.0*SUM(CASE WHEN an.is_correct=1 THEN 1 ELSE 0 END)/COUNT(*)) accuracy
            FROM answers an JOIN attempts a ON a.id=an.attempt_id JOIN questions q ON q.id=an.question_id
            WHERE a.user_id=? GROUP BY q.topic ORDER BY accuracy
            """,
            (student_id,),
        ).fetchall()
        progress = connection.execute(
            """
            SELECT l.title lesson,c.title chapter,COALESCE(p.status,'not_started') status,COALESCE(p.mastery,0) mastery
            FROM lessons l JOIN chapters c ON c.id=l.chapter_id
            LEFT JOIN progress p ON p.lesson_id=l.id AND p.user_id=?
            ORDER BY l.sequence
            """,
            (student_id,),
        ).fetchall()
    return {
        "student": dict(student),
        "attempts": [dict(row) for row in attempts],
        "topics": [dict(row) for row in topics],
        "progress": [dict(row) for row in progress],
    }


@app.get("/api/health")
def health() -> dict[str, Any]:
    with db() as connection:
        counts = {
            "users": connection.execute("SELECT COUNT(*) FROM users").fetchone()[0],
            "lessons": connection.execute("SELECT COUNT(*) FROM lessons").fetchone()[0],
            "questions": connection.execute("SELECT COUNT(*) FROM questions").fetchone()[0],
            "schedule": connection.execute("SELECT COUNT(*) FROM schedule").fetchone()[0],
        }
    return {"ok": True, "database": str(DB_PATH), "counts": counts}
