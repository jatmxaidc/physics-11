"""
Class 11 lecture plan — all 133 PW topics kept.
Shared days only where teaching time actually allows it.
Start Tue 1 Sep 2026. Theory Thu 24 Dec. Class 11 complete Thu 31 Dec.
"""
from __future__ import annotations
from datetime import date, timedelta
from collections import OrderedDict
from pathlib import Path

OUT = Path(__file__).resolve().parent / "syllabus" / "Class_11_Physics_Detailed_Lecture_Plan_2026-27.md"
START = date(2026, 9, 1)
END = date(2026, 12, 31)

HOLIDAYS = {
    date(2026, 10, 2): "Gandhi Jayanti — if closed, swap with the nearest practice day",
    date(2026, 11, 9): "Diwali week — keep the topic; use a recording if school is closed",
    date(2026, 11, 10): "Diwali week — keep the topic; use a recording if school is closed",
    date(2026, 12, 25): "Christmas — if closed, sit the lock paper at home; teach 26 Dec",
    date(2026, 12, 31): "Class 11 complete",
}

PW_TOPICS = {
    1: ("Mathematical Tools", "Trigonometry — ratios, identities, right-triangle problems"),
    2: ("Mathematical Tools", "Trigonometry — ASTC, radian–degree, height and distance"),
    3: ("Mathematical Tools", "Algebra for physics"),
    4: ("Mathematical Tools", "Binomial expansion"),
    5: ("Mathematical Tools", "A.P. and G.P."),
    6: ("Mathematical Tools", "Graphs — straight line and slope"),
    7: ("Mathematical Tools", "Logarithms"),
    8: ("Mathematical Tools", "Coordinate geometry I"),
    9: ("Mathematical Tools", "Coordinate geometry II — parabola and related curves"),
    10: ("Mathematical Tools", "Differentiation (for v and a)"),
    11: ("Mathematical Tools", "Mathematical tools — question practice"),
    12: ("Vectors", "Introduction of vector and scalar; types of vectors"),
    13: ("Vectors", "Vector addition"),
    14: ("Vectors", "Vector subtraction"),
    15: ("Vectors", "Resolution of vectors; multiplication by a scalar"),
    16: ("Vectors", "Scalar product, vector product and projection"),
    17: ("Vectors", "Vector product (cross product) in depth"),
    18: ("Vectors", "Vectors — question practice"),
    19: ("Units and Measurements", "Physical quantities and units"),
    20: ("Units and Measurements", "Dimensions and dimensional formula"),
    21: ("Units and Measurements", "Dimensional analysis I"),
    22: ("Units and Measurements", "Dimensional analysis II"),
    23: ("Units and Measurements", "Dimensional analysis III — uses and limits"),
    24: ("Units and Measurements", "Significant figures"),
    25: ("Units and Measurements", "Errors in measurement I"),
    26: ("Units and Measurements", "Errors in measurement II"),
    27: ("Units and Measurements", "Measuring instruments — Vernier and screw gauge"),
    28: ("Units and Measurements", "Units and measurements — question practice"),
    29: ("Motion in a Straight Line", "Introduction of mechanics; frame of reference; distance and displacement"),
    30: ("Motion in a Straight Line", "Average speed and average velocity"),
    31: ("Motion in a Straight Line", "Instantaneous velocity and its speed"),
    32: ("Motion in a Straight Line", "Acceleration"),
    33: ("Motion in a Straight Line", "Uniform motion and uniform acceleration"),
    34: ("Motion in a Straight Line", "Motion under gravity (free fall) I"),
    35: ("Motion in a Straight Line", "Motion under gravity (free fall) II"),
    36: ("Motion in a Straight Line", "Graphs of 1D motion I — x–t"),
    37: ("Motion in a Straight Line", "Graphs of 1D motion II — v–t"),
    38: ("Motion in a Straight Line", "Variable acceleration I"),
    39: ("Motion in a Straight Line", "Variable acceleration II"),
    40: ("Motion in a Straight Line", "Graphs of 1D motion III — mixed"),
    41: ("Motion in a Straight Line", "Graphs of 1D motion IV — mixed"),
    42: ("Motion in a Straight Line", "Motion in a straight line — question practice"),
    43: ("Motion in a Plane", "Projectile motion I — level ground"),
    44: ("Motion in a Plane", "Projectile motion II — equation of trajectory"),
    45: ("Motion in a Plane", "Projectile motion III — from a height"),
    46: ("Motion in a Plane", "Projectile motion IV — numericals"),
    47: ("Motion in a Plane", "Relative motion"),
    48: ("Motion in a Plane", "Circular motion I — UCM"),
    49: ("Motion in a Plane", "Motion in a plane — question practice"),
    50: ("Laws of Motion", "Force and Newton’s first law"),
    51: ("Laws of Motion", "Linear momentum"),
    52: ("Laws of Motion", "Newton’s second law"),
    53: ("Laws of Motion", "Newton’s third law and free-body diagram"),
    54: ("Laws of Motion", "Working with Newton’s first law"),
    55: ("Laws of Motion", "Working with Newton’s second law"),
    56: ("Laws of Motion", "Calculation of acceleration"),
    57: ("Laws of Motion", "Spring force, non-inertial feel, rocket propulsion"),
    58: ("Laws of Motion", "Friction — types of friction"),
    59: ("Laws of Motion", "Graph between applied force and friction"),
    60: ("Laws of Motion", "Angle of friction and angle of repose"),
    61: ("Laws of Motion", "Acceleration of a body on a rough surface"),
    62: ("Laws of Motion", "Dynamics of circular motion I"),
    63: ("Laws of Motion", "Dynamics of circular motion II"),
    64: ("Laws of Motion", "Dynamics of circular motion III"),
    65: ("Laws of Motion", "Laws of motion — question practice"),
    66: ("Work, Energy and Power", "Work"),
    67: ("Work, Energy and Power", "Energy"),
    68: ("Work, Energy and Power", "Work–energy theorem"),
    69: ("Work, Energy and Power", "Conservative and non-conservative force"),
    70: ("Work, Energy and Power", "Potential energy I"),
    71: ("Work, Energy and Power", "Potential energy II — spring"),
    72: ("Work, Energy and Power", "Equilibrium"),
    73: ("Work, Energy and Power", "Work, energy and power — question practice"),
    74: ("Centre of Mass and System of Particles", "Centre of mass I"),
    75: ("Centre of Mass and System of Particles", "Centre of mass II — rigid bodies and rod"),
    76: ("Centre of Mass and System of Particles", "Collision — head-on I"),
    77: ("Centre of Mass and System of Particles", "Collision — head-on II"),
    78: ("Centre of Mass and System of Particles", "Collision — head-on III — coefficient of restitution"),
    79: ("Centre of Mass and System of Particles", "Collision — 2D / oblique"),
    80: ("Centre of Mass and System of Particles", "Centre of mass — question practice"),
    81: ("Rotational Motion", "Introduction; rigid body"),
    82: ("Rotational Motion", "Moment of inertia I"),
    83: ("Rotational Motion", "Moment of inertia II — standard objects"),
    84: ("Rotational Motion", "Theorems of moment of inertia"),
    85: ("Rotational Motion", "Radius of gyration, torque, Newton’s laws for rotation"),
    86: ("Rotational Motion", "Angular momentum"),
    87: ("Rotational Motion", "Rotational kinetic energy"),
    88: ("Rotational Motion", "Conservation of total mechanical energy (rotation)"),
    89: ("Rotational Motion", "Rolling motion; rolling on an incline"),
    90: ("Rotational Motion", "Comparison of translational and rotational motion I"),
    91: ("Rotational Motion", "Comparison of translational and rotational motion II"),
    92: ("Rotational Motion", "Rotational motion — question practice"),
    93: ("Gravitation", "Law of gravitation"),
    94: ("Gravitation", "Acceleration due to gravity"),
    95: ("Gravitation", "Gravitational potential energy and gravitational potential"),
    96: ("Gravitation", "Field and potential; escape velocity; satellite motion"),
    97: ("Gravitation", "Kepler’s laws of planetary motion"),
    98: ("Gravitation", "Gravitation — question practice"),
    99: ("Mechanical Properties of Solids", "Elasticity, stress, strain, Hooke’s law"),
    100: ("Mechanical Properties of Solids", "Modulus of elasticity; Young’s modulus"),
    101: ("Mechanical Properties of Solids", "Solids — question practice"),
    102: ("Mechanical Properties of Fluids", "Introduction of hydrostatics"),
    103: ("Mechanical Properties of Fluids", "Pressure"),
    104: ("Mechanical Properties of Fluids", "Pascal’s law"),
    105: ("Mechanical Properties of Fluids", "Buoyancy and buoyant force"),
    106: ("Mechanical Properties of Fluids", "Hydrodynamics; equation of continuity"),
    107: ("Mechanical Properties of Fluids", "Bernoulli’s theorem"),
    108: ("Mechanical Properties of Fluids", "Viscosity and surface tension"),
    109: ("Mechanical Properties of Fluids", "Fluids — question practice"),
    110: ("Thermal Properties of Matter", "Thermal expansion"),
    111: ("Thermal Properties of Matter", "Applications of thermal expansion"),
    112: ("Thermal Properties of Matter", "Specific heat capacity"),
    113: ("Thermal Properties of Matter", "Latent heat and calorimetry"),
    114: ("Thermal Properties of Matter", "Heat transfer — conduction, convection, radiation"),
    115: ("Thermal Properties of Matter", "Thermal properties — question practice"),
    116: ("Kinetic Theory", "Ideal gases; real-gas equation"),
    117: ("Kinetic Theory", "Kinetic theory — question practice"),
    118: ("Thermodynamics", "Introduction; thermal equilibrium; zeroth law; first law"),
    119: ("Thermodynamics", "Different types of processes"),
    120: ("Thermodynamics", "Second law; heat engine; Carnot; heat pump; refrigerator"),
    121: ("Thermodynamics", "Thermodynamics — question practice"),
    122: ("Oscillations", "Periodic motion and oscillatory motion"),
    123: ("Oscillations", "Basic terms of oscillatory motion"),
    124: ("Oscillations", "SHM and equations of SHM"),
    125: ("Oscillations", "Energy in SHM"),
    126: ("Oscillations", "Time period of spring–block; simple pendulum"),
    127: ("Oscillations", "Oscillations — question practice"),
    128: ("Waves", "Waves and equation of waves"),
    129: ("Waves", "Characteristics of waves"),
    130: ("Waves", "Progressive wave on string; sound waves; speed of sound"),
    131: ("Waves", "Superposition; reflection and refraction of waves"),
    132: ("Waves", "Stationary waves — strings, organ pipes, beats"),
    133: ("Waves", "Waves — question practice"),
}
assert list(PW_TOPICS) == list(range(1, 134))

# Each session: (pw_ids, why_if_shared_or_full)
# why is shown only for shared days in the pace table.
SESSIONS: list[tuple[list[int], str]] = [
    # Mathematical Tools — 9 days / 11 topics
    ([1], "Full. First trig lecture; ratios and identities need a slow start."),
    ([2], "Full. ASTC + radian/degree + height-distance is already a packed hour."),
    ([3, 4], "Shared 30+30. Both are algebraic tools; binomial is one extra formula set, not a new chapter."),
    ([5], "Full. A.P./G.P. is new and used later in decay, series, and some NEET numericals."),
    ([6], "Full. Straight-line graphs run through all of kinematics. Do not rush."),
    ([7], "Full. Log laws look short; students lose marks on change-of-base and ln vs log."),
    ([8, 9], "Shared 30+30. Coordinate geometry I is the straight line already practised in graphs; II is parabola shape for later motion."),
    ([10], "Full. First genuine calculus for physics. v = dx/dt and a = dv/dt must land cleanly."),
    ([11], "Practice. Full hour of mixed maths."),
    # Vectors — 5 days / 7 topics
    ([12], "Full. Scalar vs vector and types. Foundation for the next four days."),
    ([13, 14], "Shared 40+15. Subtraction is addition of −A. Do not give subtraction its own day."),
    ([15], "Full. Components are used in almost every later chapter."),
    ([16, 17], "Shared 30+30. Dot then cross. Same hour, two operations, two physics uses (W and τ)."),
    ([18], "Practice."),
    # Units — 8 days / 10 topics
    ([19], "Full. SI, prefixes, conversion. First NCERT chapter."),
    ([20], "Full. Writing dimensional formulae by hand."),
    ([21], "Full. First time checking an equation by dimensions — needs guided examples."),
    ([22, 23], "Shared 35+20. More analysis problems, then uses and limits of the method."),
    ([24], "Full. Significant figures: rules + rounding traps."),
    ([25, 26], "Shared 30+25. Same error formulae; day 2 of PW is more cases, not a new idea."),
    ([27], "Full. Vernier and screw gauge are slow, practical, and NEET-heavy."),
    ([28], "Practice + chapter gate."),
    # 1D — 9 days / 14 topics
    ([29], "Full. Frame, position, distance, displacement. Vocabulary lecture."),
    ([30], "Full. Average speed vs average velocity. Classic trap."),
    ([31], "Full. Instantaneous velocity from calculus and from x–t slope."),
    ([32, 33], "Shared 25+30. Acceleration, then SUVAT as the uniform-a special case. One kinematics block."),
    ([34, 35], "Shared 30+25. Same free-fall equations; PW II is extra cases (two bodies, from a height)."),
    ([36, 37], "Shared 30+30. x–t then v–t in one graph lecture. Both skills, two halves."),
    ([38, 39], "Shared 30+25. One calculus idea: a variable, integrate for v and x."),
    ([40, 41], "Shared practice. PW III and IV are both mixed-graph drills."),
    ([42], "Practice + chapter gate."),
    # Plane — 5 days / 7 topics
    ([43, 44], "Shared 40+15. Level-ground projectile, then trajectory equation as the last result."),
    ([45, 46], "Shared 35+20. From-a-height setup (new), then numericals on the same setup."),
    ([47], "Full. Relative velocity in 1D and 2D. Students need the full hour."),
    ([48], "Full. UCM and centripetal acceleration."),
    ([49], "Practice + chapter gate."),
    # Laws — 11 days / 16 topics
    ([50], "Full. Inertia and N1."),
    ([51, 52], "Shared 25+30. Momentum, then N2 as F = dp/dt. One story."),
    ([53], "Full. N3 and FBD. Never share this hour."),
    ([54, 55], "Shared 30+25. Application hour after the three laws are already taught."),
    ([56, 57], "Shared 40+15. Pulleys and connected bodies need 40 min. Spring/rocket as a short second block, not dropped."),
    ([58, 59], "Shared 40+15. Friction types, then the F–f graph as the picture of the same laws."),
    ([60, 61], "Shared 25+30. Repose, then rough-surface acceleration. Same μ, two standard problems."),
    ([62], "Full. Level-road circular dynamics."),
    ([63], "Full. Banking. Own hour — this is where students fail."),
    ([64], "Full. Vertical circle. Own hour."),
    ([65], "Practice + chapter gate."),
    # WEP — 5 days / 8 topics
    ([66], "Full. Work by constant and variable force, F–s area."),
    ([67, 68, 69], "Shared 20+20+15. Energy, work–energy theorem, then conservative vs not (when ME is conserved). One energy story."),
    ([70, 71], "Shared 30+25. Gravitational U, then spring U. Same idea, two formulae."),
    ([72], "Full. Stable / unstable / neutral equilibrium."),
    ([73], "Practice + chapter gate."),
    # CoM — 5 days / 7 topics
    ([74, 75], "Shared 30+25. Two-particle CoM, then rod / rigid body as the next example."),
    ([76, 77], "Shared 30+25. Elastic then inelastic head-on. Same diagrams."),
    ([78], "Full. Coefficient of restitution numericals. Needs its own practice hour."),
    ([79], "Full. 2D / oblique collision. Own hour."),
    ([80], "Practice + chapter gate."),
    # Rotation — 9 days / 12 topics
    ([81], "Full. Rigid body and rotational kinematics language."),
    ([82, 83], "Shared 25+30. Define I, then standard objects (ring, disc, rod, plate)."),
    ([84], "Full. Parallel- and perpendicular-axis theorems with use."),
    ([85], "Full. k, torque, τ = Iα."),
    ([86], "Full. Angular momentum and conservation."),
    ([87, 88], "Shared 25+30. Rotational KE, then energy conservation including rotation."),
    ([89], "Full. Rolling and incline. Own hour."),
    ([90, 91], "Shared 30+25. Translation vs rotation table. PW split a summary into two days."),
    ([92], "Practice + chapter gate."),
    # Gravitation — 5 days / 6 topics
    ([93], "Full. Newton’s law and superposition."),
    ([94], "Full. g, g(h), g(d). Several cases — own hour."),
    ([95], "Full. V and U."),
    ([96, 97], "Shared 35+20. Escape and satellite first; Kepler as the orbit laws for the same motion."),
    ([98], "Practice + chapter gate."),
    # Solids — 2 days / 3 topics
    ([99, 100], "Shared 25+30. Hooke, then Y/B/η with Poisson and elastic energy named inside the moduli block."),
    ([101], "Practice + chapter gate."),
    # Fluids — 5 days / 8 topics
    ([102, 103, 104], "Shared 15+20+20. Hydrostatics, P = ρgh, Pascal/hydraulics. One pressure family."),
    ([105], "Full. Buoyancy. Own hour."),
    ([106, 107], "Shared 20+35. Continuity first, Bernoulli as the main theorem."),
    ([108], "Full. Viscosity and surface tension are already two topics in one PW hour. Keep the full 60 min; 30+30, do not add a third topic."),
    ([109], "Practice + chapter gate."),
    # Thermal — 5 days / 6 topics
    ([110, 111], "Shared 35+20. Expansion formulae, then applications / anomalous expansion."),
    ([112], "Full. Specific heat, Cp, Cv."),
    ([113], "Full. Latent heat and calorimetry numericals."),
    ([114], "Full. Conduction, convection, radiation, Wien, Stefan."),
    ([115], "Practice + chapter gate."),
    # KTG — 2 days / 2 topics
    ([116], "Full. Already a dense PW hour (assumptions, P, v_rms, equipartition, mean free path). Do not add more."),
    ([117], "Practice + chapter gate."),
    # Thermo — 4 days / 4 topics
    ([118], "Full. Zeroth law, Q, W, U, first law."),
    ([119], "Full. Isothermal, adiabatic, PV graphs."),
    ([120], "Full. Second law, Carnot, refrigerator."),
    ([121], "Practice + chapter gate."),
    # Oscillations — 5 days / 6 topics
    ([122, 123], "Shared 25+30. Periodic motion, then A, ω, φ. Definition hour."),
    ([124], "Full. SHM equation, v and a."),
    ([125], "Full. Energy in SHM."),
    ([126], "Full. Spring–block and pendulum derivations. Already two derivations in one PW hour."),
    ([127], "Practice + chapter gate."),
    # Waves — 5 days / 6 topics
    ([128, 129], "Shared 30+25. Wave equation, then v = fλ and characteristics."),
    ([130], "Full. String wave speed and sound. Two physical speeds."),
    ([131], "Full. Superposition, reflection, refraction."),
    ([132], "Full. Stationary waves, pipes and beats. Already three ideas in one PW hour — keep 60 min."),
    ([133], "Practice + chapter gate. Class 11 theory complete."),
]

flat = [p for s, _ in SESSIONS for p in s]
assert flat == list(range(1, 134)), f"Topic mismatch: {flat}"
assert len(SESSIONS) == 99, len(SESSIONS)

LOCK = [
    ("—", "Class 11 lock", "Mechanics mixed NEET MCQ — Ch 1–7, every topic represented", "Mock", "Mechanics MCQ"),
    ("—", "Class 11 lock", "Mechanics CBSE written — derivations and numericals", "Mock", "Mechanics written"),
    ("—", "Class 11 lock", "Fluids + thermal + KTG + thermo mixed", "Practice", "Part II mixed"),
    ("—", "Class 11 lock", "Oscillations + waves mixed", "Practice", "SHM and waves"),
    ("—", "Class 11 lock", "Full Class 11 NEET mock", "Mock", "No extra DT — mock day"),
    ("—", "Class 11 lock", "Formula / graph / derivation lock  •  Class 11 complete", "Revision", "No DT — closing review"),
]


def teaching_days(start: date, end: date) -> list[date]:
    out = []
    d = start
    while d <= end:
        if d.weekday() < 6:
            out.append(d)
        d += timedelta(days=1)
    return out


days = teaching_days(START, END)
assert len(days) == 105, len(days)


def md_day(d: date) -> str:
    return d.strftime("%a %d %b %Y")


def week_key(d: date) -> date:
    return d - timedelta(days=d.weekday())


def pw_label(pws: list[int]) -> str:
    return "+".join(str(p) for p in pws)


def topics_cell(pws: list[int]) -> str:
    if len(pws) == 1:
        return PW_TOPICS[pws[0]][1]
    return "  ·  ".join(f"({i}) {PW_TOPICS[p][1]}" for i, p in enumerate(pws, 1))


def session_type(pws: list[int], why: str) -> str:
    if any("question practice" in PW_TOPICS[p][1].lower() for p in pws):
        return "Practice"
    return "Concept"


def pace(pws: list[int]) -> str:
    if len(pws) == 1:
        name = PW_TOPICS[pws[0]][1].lower()
        return "Practice" if "question practice" in name else "Full 60 min"
    if len(pws) == 2:
        return "Shared 60 min"
    return "Shared 60 min (3 titles)"


LEC = []
for pws, why in SESSIONS:
    ch = PW_TOPICS[pws[0]][0]
    LEC.append((pw_label(pws), ch, topics_cell(pws), session_type(pws, why),
                topics_cell(pws) if len(pws) == 1 else " + ".join(PW_TOPICS[p][1] for p in pws),
                pws, why, pace(pws)))
for row in LOCK:
    LEC.append((row[0], row[1], row[2], row[3], row[4], [], "", "Lock"))
assert len(LEC) == 105

n_teach = 99
shared = sum(1 for pws, _ in SESSIONS if len(pws) > 1)
full = n_teach - shared


def write_markdown() -> None:
    lines: list[str] = []
    A = lines.append

    A("# Class 11 Physics — Detailed Lecture Plan")
    A("")
    A("## CBSE 2026–27 + NEET Foundation")
    A("")
    A("**All 133 PW topics are kept.** A topic is never deleted.")
    A("Days are packed only when the **teaching time is enough** — not by a 4-into-3 rule.")
    A("")
    A("| Setting | Value |")
    A("|---|---|")
    A("| Start | **Tue 1 Sep 2026** |")
    A("| PW topics | **133 / 133** |")
    A("| Our teaching days | **99** |")
    A(f"| Full-hour days | {full} (hard or new idea) |")
    A(f"| Shared days | {shared} (same idea, or a short add-on) |")
    A("| Lock | 6 days, 25–31 Dec |")
    A(f"| New theory ends | **{md_day(days[98])}** (Waves practice) |")
    A(f"| Class 11 complete | **{md_day(days[-1])}** |")
    A("| Lecture | 60 min · Mon–Sat (Week 1 starts Tuesday) |")
    A("| After class | Daily test · both titles on a shared day |")
    A("| Chapter gate | **70%** |")
    A("")
    A("## How a day is planned")
    A("")
    A("| Kind of day | When we use it | Clock |")
    A("|---|---|---|")
    A("| **Full 60 min** | New idea, student trap, or already-dense PW lecture (FBD, banking, rolling, Bernoulli, SHM, pipes) | 8 min DT review · 32 min concept · 15 min numerical · 5 min close |")
    A("| **Shared 60 min** | PW split one idea across two days, or the second title is a 15–30 min add-on | 8 min DT · 22–30 min title (1) · 15–25 min title (2) · 10 min mixed numerical |")
    A("| **Practice** | End of each chapter | Full hour of CBSE + NEET items + 70% gate |")
    A("")
    A("If a shared day overruns, **finish title (1) properly** and move title (2)’s numericals to the next-day DT — do not skip the definition.")
    A("")
    A("## Why some PW days are shared (and why some are not)")
    A("")
    A("| Shared PW titles | Why they share one hour |")
    A("|---|---|")

    for pws, why in SESSIONS:
        if len(pws) > 1:
            labels = ", ".join(f"PW {p}" for p in pws)
            A(f"| {labels} | {why} |")

    A("")
    A("## Kept as a full hour on purpose")
    A("")
    A("- Trigonometry I and II, graphs, logs, differentiation")
    A("- Vector resolution (components)")
    A("- Vernier / screw gauge")
    A("- Frame / distance / displacement; instantaneous velocity")
    A("- Relative velocity; UCM")
    A("- FBD; connected bodies (main block); level road; **banking**; **vertical circle**")
    A("- Work; equilibrium")
    A("- Restitution numericals; **2D collision**")
    A("- Rigid-body intro; axis theorems; torque; angular momentum; **rolling**")
    A("- Newton’s law of gravitation; g-variation; potential")
    A("- Buoyancy; viscosity + surface tension (already two topics in PW’s hour)")
    A("- Specific heat; calorimetry; heat transfer")
    A("- Entire kinetic theory hour; all four thermodynamics days")
    A("- SHM equation; energy in SHM; spring + pendulum")
    A("- String/sound; superposition; stationary waves / pipes / beats")
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
        if da.month == db.month:
            span = f"{da.strftime('%d')}–{db.strftime('%d %b')}"
        else:
            span = f"{da.strftime('%d %b')} – {db.strftime('%d %b')}"
        pws = []
        for j in range(a, b + 1):
            pws.extend(LEC[j][5])
        pw_span = f"{min(pws)}–{max(pws)}" if pws else "—"
        A(f"| {ch} | L{a+1:03d}–L{b+1:03d} | {span} | {b-a+1} | {pw_span} |")

    A("")
    A("## Holiday rule")
    A("")
    A("Keep the lecture number. If the day is closed, teach **every listed topic** on the next working day before a new chapter.")
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
    A("**Topics covered** = every PW title that day. **Pace** says whether the hour is full or shared.")
    A("")

    weeks: OrderedDict[date, list[int]] = OrderedDict()
    for i, dte in enumerate(days):
        weeks.setdefault(week_key(dte), []).append(i)

    for w_i, (_monday, idxs) in enumerate(weeks.items(), start=1):
        a, b = days[idxs[0]], days[idxs[-1]]
        if a.month == b.month:
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
    A("1. Daily test is due before the next lecture. A shared day tests **every title listed**.")
    A("2. Open the next class with weak items only.")
    A("3. Tag errors: concept / formula / calculation / reading / time.")
    A("4. Reattempt wrong numericals within 48 hours.")
    A("5. Chapter complete: NCERT examples + exercises + formula sheet + practice lecture + **≥ 70%**.")
    A("6. Below 70%: remediation within 7 days. If a shared day was the cause, reteach title (1) first.")
    A("")
    A("# After 31 December")
    A("")
    A("All 133 PW Class 11 topics are done. January is for school exams or error-log repair. Class 12 electrostatics is not in this file.")
    A("")

    OUT.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Wrote {OUT}")
    print(f"Teach {n_teach}  shared {shared}  full {full}  total {len(LEC)}")
    print(f"Theory {days[98]}  Complete {days[-1]}")


if __name__ == "__main__":
    write_markdown()
