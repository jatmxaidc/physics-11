export type Role = "student" | "teacher" | "parent";

export type Material = {
  id: string;
  title: string;
  kind: "Class notes" | "Short notes" | "Slides" | "Worksheet";
  format: "PDF" | "PPTX" | "Note";
  size?: string;
  updated?: boolean;
};

export type PhysicsClass = {
  id: number;
  lecture: string;
  date: string;
  title: string;
  chapter: string;
  description: string;
  status: "mastered" | "practiced" | "in-progress" | "upcoming";
  score?: number;
  duration: string;
  tags: string[];
  materials: Material[];
};

export const physicsClasses: PhysicsClass[] = [
  {
    id: 1,
    lecture: "Lecture 01",
    date: "24 Aug",
    title: "Trigonometric ratios & identities",
    chapter: "Mathematical Tools",
    description: "Ratios, identities, similar triangles and height–distance applications.",
    status: "mastered",
    score: 84,
    duration: "2 hr",
    tags: ["Foundation", "NEET"],
    materials: [
      { id: "1-slides", title: "Lecture 01 classroom deck", kind: "Slides", format: "PDF", size: "8.4 MB" },
      { id: "1-notes", title: "Complete class notes", kind: "Class notes", format: "PDF", size: "3.1 MB" },
      { id: "1-short", title: "Trigonometry lock-lines", kind: "Short notes", format: "Note" },
    ],
  },
  {
    id: 2,
    lecture: "Lecture 02",
    date: "26 Aug",
    title: "ASTC, height & radians",
    chapter: "Mathematical Tools",
    description: "Quadrant signs, standard values, arc length and degree–radian conversion.",
    status: "mastered",
    score: 78,
    duration: "2 hr",
    tags: ["Foundation", "CBSE + NEET"],
    materials: [
      { id: "2-slides", title: "ASTC and radians deck", kind: "Slides", format: "PDF", size: "7.2 MB" },
      { id: "2-short", title: "Radians in one page", kind: "Short notes", format: "Note" },
      { id: "2-work", title: "Angle conversion practice", kind: "Worksheet", format: "PDF", size: "480 KB" },
    ],
  },
  {
    id: 3,
    lecture: "Lecture 03",
    date: "29 Aug",
    title: "Compound angles & small-angle tools",
    chapter: "Mathematical Tools",
    description: "Compound, double and half angles, allied angles, graphs and approximations.",
    status: "practiced",
    score: 72,
    duration: "2 hr",
    tags: ["Foundation", "NEET"],
    materials: [
      { id: "3-slides", title: "Compound angle classroom deck", kind: "Slides", format: "PDF", size: "9.8 MB" },
      { id: "3-notes", title: "Class notes with solved boards", kind: "Class notes", format: "PDF", size: "4.4 MB" },
      { id: "3-short", title: "Small-angle conditions", kind: "Short notes", format: "Note", updated: true },
    ],
  },
  {
    id: 4,
    lecture: "Lecture 04",
    date: "04 Sep",
    title: "Maximum–minimum & phase",
    chapter: "Mathematical Tools",
    description: "ASTC review, amplitude, maximum–minimum and introductory phasors.",
    status: "practiced",
    score: 68,
    duration: "2 hr",
    tags: ["Foundation", "NEET trap"],
    materials: [
      { id: "4-slides", title: "Max–min and phase slides", kind: "Slides", format: "PDF", size: "11.6 MB" },
      { id: "4-short", title: "Amplitude formula card", kind: "Short notes", format: "Note" },
      { id: "4-work", title: "Max–min remediation set", kind: "Worksheet", format: "PDF", size: "620 KB" },
    ],
  },
  {
    id: 5,
    lecture: "Lecture 05",
    date: "08 Sep",
    title: "Phase, A.P., G.P. & graph basics",
    chapter: "Mathematical Tools",
    description: "Finish max–min, phase and superposition, then progressions and graph foundations.",
    status: "in-progress",
    duration: "2 hr",
    tags: ["Current class", "PW Class 04"],
    materials: [
      { id: "5-slides", title: "Lecture 05 — 61 classroom slides", kind: "Slides", format: "PDF", size: "12.1 MB", updated: true },
      { id: "5-notes", title: "Complete class notes", kind: "Class notes", format: "PDF", size: "5.2 MB" },
      { id: "5-short", title: "Phase, A.P. and G.P. lock-lines", kind: "Short notes", format: "Note" },
      { id: "5-work", title: "PW Class 04 homework", kind: "Worksheet", format: "PDF", size: "390 KB" },
    ],
  },
  {
    id: 6,
    lecture: "Lecture 06",
    date: "Next class",
    title: "Straight-line graphs",
    chapter: "Mathematical Tools",
    description: "Coordinates, distance, slope, straight-line equations and Physics graph applications.",
    status: "upcoming",
    duration: "2 hr",
    tags: ["Scheduled", "PW Class 05"],
    materials: [
      { id: "6-slides", title: "Lecture 06 teacher draft", kind: "Slides", format: "PPTX", size: "10.8 MB" },
    ],
  },
];

export const dailyQuiz = {
  title: "Lecture 05 daily quiz",
  subtitle: "Phase, progressions and graphs",
  due: "Today, 9:00 PM",
  duration: "12 min",
  questions: [
    {
      id: 1,
      prompt: "The maximum value of 3 sin θ + 4 cos θ is",
      options: ["3", "4", "5", "7"],
      answer: 2,
      explanation: "For a sin θ + b cos θ, the amplitude is √(a²+b²) = √25 = 5.",
      topic: "Maximum–minimum",
    },
    {
      id: 2,
      prompt: "Before finding phase difference between a sine and cosine expression, you should",
      options: ["Add amplitudes", "Convert both to the same trigonometric function", "Set ωt = 0", "Differentiate both"],
      answer: 1,
      explanation: "Phase angles are compared only after both expressions use sine or both use cosine.",
      topic: "Phase",
    },
    {
      id: 3,
      prompt: "For an infinite G.P. to have a finite sum, which condition is required?",
      options: ["r > 1", "r = 1", "|r| < 1", "r must be positive"],
      answer: 2,
      explanation: "The infinite sum S∞ = a/(1−r) converges only when |r| < 1.",
      topic: "Geometric progression",
    },
    {
      id: 4,
      prompt: "The graph of y = constant is",
      options: ["Vertical", "Horizontal", "Circular", "Always through the origin"],
      answer: 1,
      explanation: "The y-value stays fixed while x can change, producing a horizontal line.",
      topic: "Graphs",
    },
    {
      id: 5,
      prompt: "In an A.P., a = 2 and d = 3. What is T₁₀?",
      options: ["27", "29", "30", "32"],
      answer: 1,
      explanation: "T₁₀ = a + 9d = 2 + 27 = 29.",
      topic: "Arithmetic progression",
    },
  ],
};

export const weeklyTests = [
  {
    id: 1,
    title: "Mathematical Tools — Weekly Test 01",
    date: "12 Sep · 6:00 PM",
    duration: "45 min",
    questions: 25,
    marks: 100,
    status: "scheduled",
    scope: "Lectures 01–05",
  },
  {
    id: 2,
    title: "Trigonometry Foundation Check",
    date: "05 Sep",
    duration: "30 min",
    questions: 20,
    marks: 80,
    status: "released",
    score: 61,
    scope: "Lectures 01–04",
  },
];

export const scoreTrend = [
  { label: "L01", score: 64 },
  { label: "L02", score: 72 },
  { label: "L03", score: 76 },
  { label: "L04", score: 68 },
  { label: "L05", score: 82 },
];

export const topicPerformance = [
  { topic: "Trigonometry", accuracy: 82, status: "Strong" },
  { topic: "Radians", accuracy: 76, status: "On track" },
  { topic: "Maximum–minimum", accuracy: 58, status: "Needs review" },
  { topic: "Phase", accuracy: 64, status: "Developing" },
  { topic: "A.P. & G.P.", accuracy: 79, status: "On track" },
];
