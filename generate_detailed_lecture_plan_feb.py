"""
Class 11 lecture plan — Monday–Friday, ends 10 Feb 2027.
All 133 PW topics kept. Saturday is not a class day.
117 weekdays: 111 teaching + 6 lock.
"""
from __future__ import annotations

from collections import OrderedDict
from datetime import date, timedelta
from pathlib import Path

from generate_detailed_lecture_plan import (
    PW_TOPICS,
    md_day,
    pace,
    session_type,
    topics_cell,
    week_key,
)

OUT = Path(__file__).resolve().parent / "syllabus" / "Class_11_Physics_Detailed_Lecture_Plan_to_10_Feb_2027.md"
START = date(2026, 9, 1)
END = date(2027, 2, 10)


def teaching_days_mon_fri(start: date, end: date) -> list[date]:
    out = []
    d = start
    while d <= end:
        if d.weekday() < 5:
            out.append(d)
        d += timedelta(days=1)
    return out


HOLIDAYS = {
    date(2026, 10, 2): "Gandhi Jayanti — if closed, swap with the nearest practice day",
    date(2026, 11, 9): "Diwali week — keep the topic; use a recording if school is closed",
    date(2026, 11, 10): "Diwali week — keep the topic; use a recording if school is closed",
    date(2026, 12, 25): "Christmas — if closed, keep the lecture number and teach Mon 28 Dec",
    date(2027, 1, 1): "New Year — if closed, keep the lecture number and teach Mon 4 Jan",
    date(2027, 1, 26): "Republic Day — if closed, keep the lecture number and teach 27 Jan",
    date(2027, 2, 10): "Class 11 complete",
}

# 111 days / 133 topics. Hard titles get their own hour.
# Related PW pairs still share when 60 min is enough (5-day week to 10 Feb).
SESSIONS: list[tuple[list[int], str]] = [
    # Mathematical Tools
    ([1], "Full. First trig lecture; ratios and identities need a slow start."),
    ([2], "Full. ASTC + radian/degree + height-distance is already a packed hour."),
    ([3, 4], "Shared 30+30. Both are algebraic tools; binomial is one extra formula set."),
    ([5], "Full. A.P./G.P. is new and used later in decay, series, and NEET numericals."),
    ([6], "Full. Straight-line graphs run through all of kinematics."),
    ([7], "Full. Log laws look short; students lose marks on change-of-base."),
    ([8, 9], "Shared 30+30. Coordinate geometry I is the straight line from graphs; II is parabola."),
    ([10], "Full. First genuine calculus for physics."),
    ([11], "Practice. Full hour of mixed maths."),
    # Vectors
    ([12], "Full. Scalar vs vector and types."),
    ([13, 14], "Shared 40+15. Subtraction is addition of −A."),
    ([15], "Full. Components are used in almost every later chapter."),
    ([16], "Full. Scalar product and projection. Own hour on this 5-day track."),
    ([17], "Full. Cross product and torque. Own hour."),
    ([18], "Practice."),
    # Units
    ([19], "Full. SI, prefixes, conversion."),
    ([20], "Full. Writing dimensional formulae by hand."),
    ([21], "Full. Checking an equation by dimensions."),
    ([22, 23], "Shared 35+20. More analysis problems, then uses and limits."),
    ([24], "Full. Significant figures: rules + rounding traps."),
    ([25, 26], "Shared 30+25. Same error formulae; day 2 of PW is more cases."),
    ([27], "Full. Vernier and screw gauge."),
    ([28], "Practice + chapter gate."),
    # 1D
    ([29], "Full. Frame, position, distance, displacement."),
    ([30], "Full. Average speed vs average velocity."),
    ([31], "Full. Instantaneous velocity from calculus and from x–t slope."),
    ([32, 33], "Shared 25+30. Acceleration, then SUVAT as the uniform-a special case."),
    ([34, 35], "Shared 30+25. Same free-fall equations; PW II is extra cases."),
    ([36], "Full. x–t graphs. Own hour on this 5-day track."),
    ([37], "Full. v–t graphs. Own hour."),
    ([38, 39], "Shared 30+25. One calculus idea: a variable, integrate for v and x."),
    ([40, 41], "Shared practice. PW III and IV are both mixed-graph drills."),
    ([42], "Practice + chapter gate."),
    # Plane
    ([43], "Full. Level-ground projectile. Own hour."),
    ([44], "Full. Equation of trajectory. Own hour."),
    ([45, 46], "Shared 35+20. From-a-height setup, then numericals on the same setup."),
    ([47], "Full. Relative velocity in 1D and 2D."),
    ([48], "Full. UCM and centripetal acceleration."),
    ([49], "Practice + chapter gate."),
    # Laws
    ([50], "Full. Inertia and N1."),
    ([51, 52], "Shared 25+30. Momentum, then N2 as F = dp/dt."),
    ([53], "Full. N3 and FBD. Never share this hour."),
    ([54, 55], "Shared 30+25. Application hour after the three laws are already taught."),
    ([56], "Full. Connected bodies / pulleys. Own hour."),
    ([57], "Full. Spring force, non-inertial feel, rocket. Own hour."),
    ([58, 59], "Shared 40+15. Friction types, then the F–f graph."),
    ([60, 61], "Shared 25+30. Repose, then rough-surface acceleration."),
    ([62], "Full. Level-road circular dynamics."),
    ([63], "Full. Banking. Own hour."),
    ([64], "Full. Vertical circle. Own hour."),
    ([65], "Practice + chapter gate."),
    # WEP
    ([66], "Full. Work by constant and variable force."),
    ([67], "Full. Energy. Own hour on this 5-day track."),
    ([68], "Full. Work–energy theorem. Own hour."),
    ([69], "Full. Conservative vs non-conservative force. Own hour."),
    ([70, 71], "Shared 30+25. Gravitational U, then spring U."),
    ([72], "Full. Stable / unstable / neutral equilibrium."),
    ([73], "Practice + chapter gate."),
    # CoM
    ([74, 75], "Shared 30+25. Two-particle CoM, then rod / rigid body."),
    ([76], "Full. Head-on collision I. Own hour."),
    ([77], "Full. Head-on collision II. Own hour."),
    ([78], "Full. Coefficient of restitution numericals."),
    ([79], "Full. 2D / oblique collision."),
    ([80], "Practice + chapter gate."),
    # Rotation
    ([81], "Full. Rigid body and rotational kinematics language."),
    ([82], "Full. Moment of inertia defined. Own hour."),
    ([83], "Full. Standard objects (ring, disc, rod, plate). Own hour."),
    ([84], "Full. Parallel- and perpendicular-axis theorems."),
    ([85], "Full. k, torque, τ = Iα."),
    ([86], "Full. Angular momentum and conservation."),
    ([87], "Full. Rotational kinetic energy. Own hour."),
    ([88], "Full. Conservation of mechanical energy including rotation. Own hour."),
    ([89], "Full. Rolling and incline."),
    ([90, 91], "Shared 30+25. Translation vs rotation table."),
    ([92], "Practice + chapter gate."),
    # Gravitation
    ([93], "Full. Newton’s law and superposition."),
    ([94], "Full. g, g(h), g(d)."),
    ([95], "Full. V and U."),
    ([96, 97], "Shared 35+20. Escape and satellite first; Kepler as the orbit laws."),
    ([98], "Practice + chapter gate."),
    # Solids
    ([99, 100], "Shared 25+30. Hooke, then Y/B/η."),
    ([101], "Practice + chapter gate."),
    # Fluids
    ([102], "Full. Hydrostatics. Own hour."),
    ([103], "Full. Pressure. Own hour."),
    ([104], "Full. Pascal’s law. Own hour."),
    ([105], "Full. Buoyancy."),
    ([106], "Full. Continuity. Own hour."),
    ([107], "Full. Bernoulli’s theorem. Own hour."),
    ([108], "Full. Viscosity and surface tension."),
    ([109], "Practice + chapter gate."),
    # Thermal
    ([110, 111], "Shared 35+20. Expansion formulae, then applications."),
    ([112], "Full. Specific heat, Cp, Cv."),
    ([113], "Full. Latent heat and calorimetry."),
    ([114], "Full. Conduction, convection, radiation."),
    ([115], "Practice + chapter gate."),
    # KTG
    ([116], "Full. Ideal gases; kinetic theory."),
    ([117], "Practice + chapter gate."),
    # Thermo
    ([118], "Full. Zeroth law, Q, W, U, first law."),
    ([119], "Full. Isothermal, adiabatic, PV graphs."),
    ([120], "Full. Second law, Carnot, refrigerator."),
    ([121], "Practice + chapter gate."),
    # Oscillations
    ([122, 123], "Shared 25+30. Periodic motion, then A, ω, φ."),
    ([124], "Full. SHM equation, v and a."),
    ([125], "Full. Energy in SHM."),
    ([126], "Full. Spring–block and pendulum derivations."),
    ([127], "Practice + chapter gate."),
    # Waves
    ([128, 129], "Shared 30+25. Wave equation, then v = fλ and characteristics."),
    ([130], "Full. String wave speed and sound."),
    ([131], "Full. Superposition, reflection, refraction."),
    ([132], "Full. Stationary waves, pipes and beats."),
    ([133], "Practice + chapter gate. Class 11 theory complete."),
]

flat = [p for s, _ in SESSIONS for p in s]
assert flat == list(range(1, 134)), f"Topic mismatch: {flat}"
assert len(SESSIONS) == 111, len(SESSIONS)

LOCK = [
    ("—", "Class 11 lock", "Mechanics mixed NEET MCQ — kinematics to rotation, every topic represented", "Mock", "Mechanics MCQ"),
    ("—", "Class 11 lock", "Mechanics CBSE written — derivations and numericals", "Mock", "Mechanics written"),
    ("—", "Class 11 lock", "Gravitation + solids + fluids mixed", "Practice", "Part II mixed"),
    ("—", "Class 11 lock", "Thermal + KTG + thermodynamics mixed", "Practice", "Thermal block"),
    ("—", "Class 11 lock", "Oscillations + waves mixed  •  Full Class 11 NEET mock", "Mock", "No extra DT — mock day"),
    ("—", "Class 11 lock", "Formula / graph / derivation lock  •  Class 11 complete", "Revision", "No DT — closing review"),
]

days = teaching_days_mon_fri(START, END)
assert len(days) == 117, len(days)


def pw_label(pws: list[int]) -> str:
    return "+".join(str(p) for p in pws)


LEC = []
for pws, why in SESSIONS:
    ch = PW_TOPICS[pws[0]][0]
    LEC.append((
        pw_label(pws), ch, topics_cell(pws), session_type(pws, why),
        topics_cell(pws),
        pws, why, pace(pws),
    ))
for row in LOCK:
    LEC.append((row[0], row[1], row[2], row[3], row[4], [], "", "Lock"))
assert len(LEC) == 117, len(LEC)

n_teach = 111
shared = sum(1 for pws, _ in SESSIONS if len(pws) > 1)
full = n_teach - shared


def write_markdown() -> None:
    lines: list[str] = []
    A = lines.append

    A("# Class 11 Physics — Detailed Lecture Plan (to 10 February 2027)")
    A("")
    A("## CBSE 2026–27 + NEET Foundation  ·  Monday–Friday")
    A("")
    A("**All 133 PW topics are kept.** Saturday is **not** a class day.")
    A("Hard titles get a full hour. Related PW pairs share a day only when 60 minutes is enough.")
    A("")
    A("| Setting | Value |")
    A("|---|---|")
    A("| Start | **Tue 1 Sep 2026** |")
    A("| PW topics | **133 / 133** |")
    A(f"| Our teaching days | **{n_teach}** |")
    A(f"| Full-hour days | {full} |")
    A(f"| Shared days | {shared} |")
    A("| Lock | 6 days, 3–10 Feb 2027 (weekdays only) |")
    A(f"| New theory ends | **{md_day(days[n_teach - 1])}** (Waves practice) |")
    A(f"| Class 11 complete | **{md_day(days[-1])}** |")
    A("| Lecture | 60 min · **Monday–Friday** (Week 1 starts Tuesday; no Saturday class) |")
    A("| Saturday | Daily-test catch-up, NCERT, error log — not a lecture |")
    A("| After class | Daily test of 5–10 questions |")
    A("| Chapter gate | **70%** |")
    A("")
    A("## How a day is planned")
    A("")
    A("| Kind of day | When we use it | Clock |")
    A("|---|---|---|")
    A("| **Full 60 min** | New idea, student trap, or already-dense PW lecture | 8 min DT review · 32 min concept · 15 min numerical · 5 min close |")
    A("| **Shared 60 min** | PW split one idea across two days, or a short add-on | 8 min DT · 22–30 min title (1) · 15–25 min title (2) · 10 min mixed numerical |")
    A("| **Practice** | End of each chapter | Full hour of CBSE + NEET items + 70% gate |")
    A("| **Lock** | 3–10 Feb, Mon–Fri | Mixed papers, mock, formula lock |")
    A("")
    A("If a shared day overruns, **finish title (1) properly** and move title (2)’s numericals to the next-day DT.")
    A("")
    A("## Why some PW days are shared")
    A("")
    A("| Shared PW titles | Why they share one hour |")
    A("|---|---|")
    for pws, why in SESSIONS:
        if len(pws) > 1:
            labels = ", ".join(f"PW {p}" for p in pws)
            A(f"| {labels} | {why} |")
    A("")
    A("Dot product, cross product, x–t, v–t, projectile, pulleys, energy / WET / conservative, collisions, MI, rotational energy, hydrostatics / pressure / Pascal, continuity and Bernoulli each get **their own hour** on this track.")
    A("")
    A("## Chapter allocation")
    A("")
    A("| Block | Lectures | Dates | Days | PW topics |")
    A("|---|---|---|---|---|")

    ranges = OrderedDict()
    for i, row in enumerate(LEC):
        ch = row[1]
        if ch not in ranges:
            ranges[ch] = [i, i]
        else:
            ranges[ch][1] = i
    for ch, (a, b) in ranges.items():
        da, db = days[a], days[b]
        if da.month == db.month and da.year == db.year:
            span = f"{da.strftime('%d')}–{db.strftime('%d %b %Y')}"
        else:
            span = f"{da.strftime('%d %b %Y')} – {db.strftime('%d %b %Y')}"
        pws = []
        for j in range(a, b + 1):
            pws.extend(LEC[j][5])
        pw_span = f"{min(pws)}–{max(pws)}" if pws else "—"
        A(f"| {ch} | L{a+1:03d}–L{b+1:03d} | {span} | {b-a+1} | {pw_span} |")

    A("")
    A("## Holiday rule")
    A("")
    A("Keep the lecture number. If the day is closed, teach **the listed topic** on the next working day before a new chapter.")
    A("")
    A("| Date | Note |")
    A("|---|---|")
    for hd, note in HOLIDAYS.items():
        A(f"| {hd.strftime('%a %d %b %Y')} | {note} |")

    A("")
    A("---")
    A("")
    A("# Lecture calendar")
    A("")
    A("**Topics covered** = every PW title that day. **Pace** says whether the hour is full or shared. No Saturday rows.")
    A("")

    weeks: OrderedDict[date, list[int]] = OrderedDict()
    for i, dte in enumerate(days):
        weeks.setdefault(week_key(dte), []).append(i)

    for w_i, (_monday, idxs) in enumerate(weeks.items(), start=1):
        a, b = days[idxs[0]], days[idxs[-1]]
        if a.month == b.month and a.year == b.year:
            wr = f"{a.strftime('%d')}–{b.strftime('%d %b %Y')}"
        else:
            wr = f"{a.strftime('%d %b')} – {b.strftime('%d %b %Y')}"
        ch0, ch1 = LEC[idxs[0]][1], LEC[idxs[-1]][1]
        title = ch0 if ch0 == ch1 else f"{ch0} → {ch1}"
        A(f"## Week {w_i}: {wr} — {title}")
        A("")
        A("| L | Date | PW | Pace | Topics covered | Type | Daily test |")
        A("|---|---|---|---|---|---|---|")
        for i in idxs:
            dte = days[i]
            pw, ch, topics, typ, dt, _pws, _why, pc = LEC[i]
            date_s = md_day(dte)
            if dte in HOLIDAYS:
                date_s += " *"
            A(f"| {i+1:03d} | {date_s} | {pw} | {pc} | {topics} | {typ} | DT-{i+1:03d} · {dt} |")
        A("")
        flagged = [i for i in idxs if days[i] in HOLIDAYS]
        for i in flagged:
            A(f"> \\* {md_day(days[i])}: {HOLIDAYS[days[i]]}")
        if flagged:
            A("")

    A("---")
    A("")
    A("# Daily-test and progress rules")
    A("")
    A("1. Daily test is due before the next lecture. A shared day tests **every title listed**. Saturday may be used to catch up.")
    A("2. Open the next class with weak items only.")
    A("3. Tag errors: concept / formula / calculation / reading / time.")
    A("4. Reattempt wrong numericals within 48 hours.")
    A("5. Chapter complete: NCERT examples + exercises + formula sheet + practice lecture + **≥ 70%**.")
    A("6. Below 70%: remediation within 7 days before adding a new chapter load.")
    A("")
    A("# After 10 February")
    A("")
    A("All 133 PW Class 11 topics are done. Mid-February onward is for school exams or error-log repair. Class 12 electrostatics is not in this file.")
    A("")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Wrote {OUT}")
    print(f"Teach {n_teach}  shared {shared}  full {full}  total {len(LEC)}")
    print(f"Theory {days[n_teach - 1]}  Complete {days[-1]}")


if __name__ == "__main__":
    write_markdown()
