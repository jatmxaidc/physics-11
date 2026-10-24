# Excellence Class 11 Physics — continue on a new device

Copy this entire folder to the other computer and open it in Cursor. Then paste the prompt below as the first message. Do not start from an empty chat and try to reinvent the decks. The `.py` generators, `pw slides`, `lecture slides`, and `class notes` are the source of the current quality.

Install on the new device: Python 3, `python-pptx`, `pymupdf`, `reportlab`, `pywin32`, and desktop Microsoft PowerPoint (PDF export uses PowerPoint). Fonts: Aptos and Georgia if present; Arial is the fallback. Formula-sheet PDFs use DejaVu.

---

## Prompt to paste

You are building Excellence Class 11 Physics (2026–27) lecture decks and class-note PDFs for a real class. Match the quality already in this folder. Read the existing generators before you write a new one. Do not invent a second visual system.

### What this folder is

- `pw slides/` — source class notes (Physics Wallah / MR Sir). Mostly image pages, not selectable text. Render with `pymupdf` (`fitz` is deprecated), read the page images, and teach from those. Recalculate every number. A tick or a handwritten answer on the source slide can be wrong.
- `generate_lecture_*.py` — python-pptx builders. The current dark Excellence system is `generate_lecture_units_06.py` and, for motion, `generate_lecture_motion_05.py`. Copy helpers from Motion 05 (or Units 06), including subscript handling. Do not reuse `generate_lecture_05_motion_basics.py` or other early light-theme drafts.
- `lecture slides/` — finished `.pptx` and `.pdf`.
- `class notes/` — the PDF the students keep. After each new deck, export a PDF and copy it here.
- Formula sheets: `generate_basic_maths_formula_sheet.py`, `generate_vectors_formula_sheet.py`, `generate_units_formula_sheet.py`.

Shell is PowerShell. Chain commands with `;`, never `&&`. The console is cp1252, so a Unicode `print` can crash. Write a small `.py` file instead of a quoted one-liner, or set `sys.stdout.reconfigure(encoding="utf-8")`.

Do not commit, push, or change git config unless asked.

### Slide quality (this is the standard)

Widescreen 13.333 × 7.5 in, blank layout, near-black background, gold titles, white body, teal/blue/green/orange section bars. Cover is fixed:

- EXCELLENCE
- CLASS 11   -   2026-27
- White band: Physics | Motion 0N (or the real lecture name)
- Chapter line in Georgia, gold subtitle
- By - Excellence Programme
- The class date the user gives. Do not invent a date. Motion 01 and 02 were 20 October 2026. Motion 03, 04, and 05 were 22 October 2026.

Student-facing text:

- Physics only. Short teaching voice. One idea, then the steps.
- Theory is a paragraph, not a stack of slogan bullets.
- Worked lines use the existing `fml` / `ncert` / `result_card` / `trap_card` helpers.
- State the fact directly. Do not write contrastive lines such as “X, not Y” or “X rather than Y”.
- Every figure must show the object and the path, or the axes and the curve, so a student can see the situation. A text-only slide is not enough when the source has a graph, a path, or a vector.
- One formula per line. Never join two formulas with a middle dot.
- In source strings, write powers with a caret (`t^2`, `m/s^2`, `10^(-3)`). `tx()` turns those into Unicode superscripts. Write subscripts as `v_i`, `x_{avg}`. The letter f has no Unicode subscript; the helper marks it and sets OOXML `a:baseline val="-25000"`. `Font.subscript` does not work. Never leave a literal underscore like `r_f` or `x_f` on a slide.
- Recalculate. If your arithmetic disagrees with the source image, teach the correct value and say what the source wrote.
- Do not repeat a topic or a problem type already taught in an earlier lecture of the same chapter, and do not repeat it inside the new lecture. A later lecture adds only the new ideas and the new numericals.
- After the deck builds, export PDF with PowerPoint: `SaveAs(path, 32)`, then copy the PDF bytes into `class notes/`. Open a few figure slides (or render the PDF) and fix overflow, arrows that point the wrong way, labels on top of curves, and formulas that still show underscores. Delete temp render folders (`_motNN`, `_mNN_*.py`, page png dumps) when the PDF is copied. Do not leave audit image folders behind.

### Formula-sheet quality

ReportLab, A4, warm cream/teal cards, DejaVu, plain-language meaning next to the formula. `nice_table` returns `[table, Spacer]` with no `KeepTogether` (a wide table inside `KeepTogether` vanishes). Cards return lists. Use `story.extend`. Keep spacing tight. Do not add extra `PageBreak`s after the cover. Trig tables are split (0°–90° and 90°–180°) so they fit. One formula per line.

### Motion in a straight line — already done

Source: `pw slides/motion in straight line/`. Thirteen lectures exist. Files 01–07, 09, and 13 have ASCII names. 08, 10, 11, and 12 may still have non-ASCII names; list the directory in Python if the console shows question marks.

Built, and do not rebuild unless asked:

- Motion 01 — frame of reference, rest and motion, distance and displacement, with visible figures. `Lecture_Motion_01_Frame_Distance_Displacement`.
- Motion 02 — deeper distance and displacement: 1D/2D/3D paths, shortcuts, circular arcs. `Lecture_Motion_02_Distance_Displacement_Deep`.
- Motion 03 — only what was new after 01–02: harder path problems and the speed block (average speed, average velocity, time-mix and distance-mix). `Lecture_Motion_03_Advanced_and_Speed`.
- Motion 04 — instantaneous velocity, |average velocity| and average speed, v–t kinks versus a real U-turn, chord versus tangent, nested and impossible averages. `Lecture_Motion_04_Velocity`.
- Motion 05 — acceleration: a = dv/dt and a = v dv/dx, constant-a arithmetic, which formulas have constant a, including v = √x giving a = 1/2. `Lecture_Motion_05_Acceleration`.

If the user says “06” or “lecture 06” during the motion chapter, build Motion 06 from PW motion-in-a-straight-line 06 only. Skip anything already in 01–05. Same rule for later numbers. Do not start the next lecture unless asked.

`pw slides/motion in a plane/` is a different chapter. Do not mix it into straight-line lectures.

### How to build the next lecture

1. Render the matching PW PDF to images and read the pages.
2. List what is genuinely new. Drop recap, homework, and thank-you pages.
3. Copy the helper block from `generate_lecture_motion_05.py` up to the cover. Change the docstring and `OUTPUT`.
4. Write the slide body in the same helpers: cover, “What is new today”, idea slides, figures, worked examples, question then answer for real MCQs, lock lines, thank you.
5. Run it, export the PDF, copy it to `class notes/`, check the figure slides, delete temps.
6. Tell the user the slide count and the two paths (`lecture slides/` pptx+pdf and `class notes/` pdf).

When the user only writes a lecture number, that is the request. Continue the current chapter. Do not ask which chapter if the conversation is already inside motion.
