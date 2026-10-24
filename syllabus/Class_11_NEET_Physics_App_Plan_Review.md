# Class 11 NEET Physics Study App — Recommended Plan Changes

## Overall assessment

The proposed app direction is strong. A real Next.js application with authentication and cloud progress is the correct deliverable. The main change needed is to make the **daily learning-and-feedback cycle** the centre of the product rather than treating quizzes as one screen among many.

The most valuable first demo is not the largest amount of lesson content. It is one complete working loop:

**Lesson → Daily test → Attempt data → Teacher analysis → Next-day review → Retry → Mastery update**

If this loop works well for Mathematical Tools and Units and Measurements, it can be expanded safely to every later chapter.

---

# Critical changes before development

## 1. Replace the read-only Mentor role with a Teacher role

The current plan describes the mentor dashboard as read-only. That conflicts with the requirement that a teacher should review test data, solve priority questions, assign remediation and manage the next topic.

Recommended roles:

### Student

- Read lessons
- Complete daily tests
- Add highlights and private notes
- Retry incorrect questions
- View progress and error log

### Teacher

- View linked students
- View attempt and topic analytics
- Create or publish daily tests
- Review common mistakes
- Mark questions as reviewed
- Assign retries or remediation
- Adjust lesson dates and unlock lessons
- Add feedback to a student without editing private notes

### Parent — later or lightweight v1

- Read-only summary
- Attendance and test completion
- Weekly progress
- Strong and weak topics
- No access to private note text

If only two roles are possible in v1, use **Student** and **Teacher**. A parent report can initially be generated from the Teacher dashboard.

## 2. Make the daily test mandatory in the core flow

Each scheduled lecture should have one linked daily test:

- 5–10 questions
- Expected completion time: 10–15 minutes
- Due before the next lecture
- Approximately 70% from the current lesson
- Approximately 30% spaced retrieval from previous lessons
- Low-stakes and diagnostic
- One recovery opportunity each week for a legitimately missed test

Do not hard-lock the next lesson solely because a student scores below 70%. Require completion, create remediation below 70%, and allow a teacher override. A strict score lock can cause a struggling student to fall further behind.

Use the 70% gate at chapter level after sufficient practice, not as an inflexible daily gate.

## 3. Add a Next-Day Review Queue

Before starting a new topic, the teacher needs a short actionable view—not a large analytics dashboard.

The review queue should show:

- Questions with the highest class or student miss rate
- High-confidence wrong answers
- Correct answers that took unusually long
- Repeated misconceptions
- Students who did not attempt the test
- Questions due for a 48-hour retry

The teacher reviews the top priority items for 5–10 minutes and marks them as reviewed.

## 4. Import the existing dated lecture plan

Use `Class_11_Physics_Detailed_Lecture_Planner_Aug2026-Jan2027.csv` as the schedule source.

Each scheduled lecture should link to:

- Date and lecture number
- Chapter, topic and subtopic
- Lesson content
- Daily test
- Homework
- Next-day review status
- Student completion

The teacher must be able to reschedule a lecture because of holidays, absence or school examinations. Dates should be data, not hardcoded into the page components.

---

# Recommended first demo scope

## Keep in the first demo

- Student and Teacher authentication
- Secure teacher–student linking
- Dated lecture path
- Mathematical Tools lessons
- Units and Measurements lessons
- Daily 5–10 question tests
- Attempt submission and automatic marking for objective questions
- Question-level attempt storage
- Error log and 48-hour retry
- Teacher review queue
- Topic-wise analytics
- Student progress page
- Locked roadmap for later chapters
- Mobile-responsive interface

## Include only one mini tool initially

Start with the **unit converter** or **significant-figures checker**. One polished tool is enough to prove the interactive lesson model.

## Move to the next release

- Multiple highlight colours
- Notebook aggregation
- Automatic formula sheet
- Flashcards
- Streaks
- Additional mini tools
- Parent login

Highlights are useful, but they should not delay the daily test, analytics and teacher-review loop.

---

# Revised app screens

## Student Home

- Today’s scheduled lecture
- Daily test due
- Retry due within 48 hours
- Current chapter progress
- Next lesson
- Full-year locked roadmap
- Recent improvement rather than only a streak

## Lesson Reader

- Previous-test review summary
- CBSE/NEET tags
- Concept → example → try-it flow
- Stable content blocks
- One optional interactive tool
- Link to the daily test

## Daily Test

- 5–10 questions
- Current-topic and spaced-retrieval labels
- Save progress automatically
- Submit confirmation
- Score after submission
- Explanations after submission or teacher release
- Retry status for incorrect questions

## Error Log

- Wrong question
- Chapter, topic and misconception
- Student’s selected answer
- Correct answer and explanation
- Retry deadline
- Retry result
- Resolved or unresolved status

## Teacher Dashboard

- Students requiring attention
- Tests due or missing
- Daily completion rate
- Seven-day accuracy trend
- Weakest topics
- Unresolved errors
- Next-day review queue

## Teacher Student Detail

- Lesson completion timeline
- Daily-test history
- Topic mastery
- Error patterns
- Retry compliance
- Teacher feedback
- Schedule adjustments

---

# Data that must be captured

## Question metadata

Every question should have:

- Chapter
- Topic
- Subtopic
- CBSE, NEET or common tag
- Conceptual, numerical, graph, derivation or experimental-skill type
- Difficulty
- Correct answer
- Explanation
- Misconception code where applicable
- Content version

Without consistent question tags, topic-wise analytics will not be reliable.

## Attempt data

For every answer store:

- Student
- Quiz and question
- Selected or entered answer
- Correctness
- Attempt number
- Submission time
- Time spent, when available
- Confidence, optional
- Retry deadline
- Resolution status

## Suggested database areas

- Profiles and roles
- Teacher–student relationships
- Courses and enrolments
- Schedule and lecture instances
- Chapters, topics and lessons
- Quizzes and questions
- Quiz attempts and answer attempts
- Topic mastery
- Error log and retries
- Highlights and notes
- Teacher feedback

---

# Analytics recommendations

## Student view

Keep analytics simple:

- Daily-test completion
- Accuracy by chapter
- Topics to revise
- Retry queue
- Progress toward chapter completion
- Improvement over time

## Teacher view

Show actionable metrics:

- Completion rate
- Accuracy by topic
- Question miss rate
- Repeated misconception count
- Median response time where useful
- High-confidence wrong answers
- Unresolved errors
- Seven-day improvement

Avoid presenting one unexplained “mastery percentage.” If a mastery score is used, show how it is calculated and require a minimum number of attempts.

---

# Content and highlight design

## Stable content identifiers

Highlights can break when lesson text is edited. Every selectable lesson block should have a stable ID and content version.

Store:

- Lesson ID
- Block ID
- Selected quote
- Start and end offsets
- Content version

If lesson content changes, the quote helps recover or flag the highlight.

## Privacy rule

Student notes should remain private if that is the product promise. The teacher may see:

- Note count
- Highlight count
- Highlight category
- Recent activity

The teacher should not see private note text unless the student explicitly shares it.

---

# Security and reliability requirements

- Supabase Row Level Security on every student-data table
- A teacher can access only explicitly linked students
- Students cannot read another student’s attempts or notes
- Parent access is separately authorised
- No service-role key in browser code
- No secrets committed to git
- Input validation on quiz submissions
- Server-side authorisation for teacher actions
- Data export and deletion process
- Automated database backup or scheduled export
- Clear timezone handling using Asia/Kolkata
- Autosave for tests and notes
- Mobile-first and low-bandwidth design

Because the student may be a minor, collect only necessary personal data and obtain appropriate parent or guardian consent.

---

# Recommended build order

## Milestone 1 — Secure foundation

- Next.js app
- Supabase project
- Student and Teacher login
- Profiles, roles and teacher–student linking
- Row Level Security tests

## Milestone 2 — Complete learning loop

- Import lecture schedule
- Two demo content areas
- Daily test
- Attempt storage
- Automatic objective marking
- Error log and retry
- Teacher review queue

## Milestone 3 — Useful analytics

- Topic and question tagging
- Student progress
- Teacher topic analysis
- Weak-topic and missing-test alerts
- Weekly parent summary

## Milestone 4 — Learning enhancements

- Highlights and notes
- Formula sheet
- Flashcards
- Additional mini tools

## Milestone 5 — Deployment

- Production Supabase configuration
- Vercel deployment
- Environment variables
- Privacy and data-retention settings
- Mobile and low-bandwidth testing

---

# Important product decisions

Before implementation, decide:

1. Can a teacher edit and publish questions, or are questions stored only in the repository?
2. Are explanations shown immediately after submission or after the teacher reviews the test?
3. Can students access the next lesson after completing a low-scoring daily test?
4. Will a parent have a separate login in v1?
5. Will CBSE written answers be teacher-marked, self-marked with a rubric, or postponed?
6. How are cancelled classes and holidays rescheduled?
7. How long will student attempt data be retained?

## Recommended defaults

- Questions authored in repository content for the first demo
- Explanations shown after submission
- Completion unlocks the next lesson; score below 70% creates remediation
- Parent receives a report first; separate parent login later
- Objective questions auto-marked; one weekly CBSE response teacher-marked
- Teacher can reschedule lecture instances
- Attempts retained for the academic year with export and deletion options

---

# Final recommendation

Proceed with the app, but reduce the first demo to the smallest complete academic system:

**Student login → lesson → daily test → tracked attempt → teacher review queue → next-day correction → retry → progress update**

Build highlights, flashcards and multiple mini tools only after this loop is reliable. The daily feedback cycle and teacher analytics are the features that make this product meaningfully better for the planned programme.

