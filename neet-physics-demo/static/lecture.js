const slides = [
  {
    title: "Every measurement tells a story.",
    short: "Opening hook",
    time: "3 minutes",
    cue: `
      <p>Do not begin with definitions. Ask students:</p>
      <ul>
        <li>“How far is your home from here?”</li>
        <li>“Can the answer be only 5?”</li>
        <li>“What is missing?”</li>
      </ul>
      <p>Lead them to the idea that a physical measurement needs both a number and a unit.</p>`,
    html: `
      <div class="slide-kicker">Class 11 Physics · Demo lecture</div>
      <h1 class="slide-title">Every measurement tells a story.</h1>
      <p class="slide-subtitle">Physical quantities, SI units and the art of converting between them.</p>
      <div class="opening-visual"><strong>SI</strong></div>
      <div class="slide-footer"><span>Orbit Physics</span><span>Learn → Test → Analyse → Correct</span></div>`,
  },
  {
    title: "A $125 million unit mistake",
    short: "Why units matter",
    time: "5 minutes",
    cue: `
      <p>Tell the Mars Climate Orbiter story in simple language. One engineering team used pound-force seconds while another expected newton seconds.</p>
      <p><strong>Teaching point:</strong> numbers without a shared unit system can be dangerously ambiguous.</p>
      <p>Ask: “Where else could a unit mistake matter—medicine, construction, travel?”</p>`,
    html: `
      <div class="slide-kicker">Opening story · Real consequence</div>
      <div class="split-layout">
        <div>
          <span class="big-number">$125M</span>
          <h1 class="slide-title" style="font-size:clamp(28px,4vw,58px)">Lost because two teams used different units.</h1>
        </div>
        <div class="story-card danger">
          <h3>Mars Climate Orbiter · 1999</h3>
          <p>One team used imperial force units. Another expected SI units. The mismatch sent the spacecraft too close to Mars.</p>
        </div>
      </div>
      <div class="slide-footer"><span>Units are part of the answer—not decoration.</span><span>Discuss</span></div>`,
  },
  {
    title: "Can it be measured?",
    short: "Physical quantity",
    time: "5 minutes",
    cue: `
      <p>Invite fast responses. Ask students to vote with thumbs up/down:</p>
      <ul>
        <li>Length — measurable</li>
        <li>Happiness — not a physical quantity in this context</li>
        <li>Temperature — measurable</li>
        <li>Beauty — not objectively measurable</li>
      </ul>
      <p>Then reveal the definition.</p>`,
    html: `
      <div class="question-prompt">
        <div class="slide-kicker" style="justify-content:center">Quick class challenge</div>
        <h1 class="slide-title">Which of these are physical quantities?</h1>
        <div class="quantity-grid">
          <div class="quantity-card"><strong>Length</strong><span>Can be measured</span></div>
          <div class="quantity-card"><strong>Happiness</strong><span>Can it be objective?</span></div>
          <div class="quantity-card"><strong>Temperature</strong><span>Can be measured</span></div>
          <div class="quantity-card"><strong>Beauty</strong><span>Can everyone agree?</span></div>
          <div class="quantity-card"><strong>Mass</strong><span>Can be measured</span></div>
          <div class="quantity-card"><strong>Time</strong><span>Can be measured</span></div>
        </div>
        <span class="answer-reveal">Physical quantity = numerical value × unit</span>
      </div>`,
  },
  {
    title: "Base vs derived quantities",
    short: "Quantity families",
    time: "6 minutes",
    cue: `
      <p>Use a “building blocks” analogy.</p>
      <ul>
        <li>Base quantities are independent building blocks.</li>
        <li>Derived quantities are constructed using base quantities.</li>
      </ul>
      <p>Ask students to classify speed, area and force. Draw arrows between quantity and formula using the pen.</p>`,
    html: `
      <div class="slide-kicker">The quantity family</div>
      <h1 class="slide-title">Base quantities build the language. Derived quantities build the physics.</h1>
      <div class="split-layout" style="margin-top:22px">
        <div class="story-card success">
          <h3>Base quantities</h3>
          <p>Independent foundations: length, mass, time, current, temperature, amount of substance and luminous intensity.</p>
        </div>
        <div class="story-card">
          <h3>Derived quantities</h3>
          <p>Built from foundations: area = length², speed = length/time, force = mass × acceleration.</p>
        </div>
      </div>
      <div class="slide-footer"><span>Ask: Is velocity base or derived?</span><span>CBSE + NEET common</span></div>`,
  },
  {
    title: "The seven SI foundations",
    short: "SI base units",
    time: "8 minutes",
    cue: `
      <p>Do not force memorisation during the demo. Connect each unit to an instrument or everyday measurement.</p>
      <p>Spend most time on metre, kilogram and second. Briefly introduce the other four.</p>
      <p>Point out: kelvin has no degree symbol; unit symbols do not become plural.</p>`,
    html: `
      <div class="slide-kicker">International System of Units</div>
      <h1 class="slide-title">Seven agreed foundations. One shared scientific language.</h1>
      <div class="unit-grid">
        <div class="unit-card"><div class="unit-symbol">m</div><strong>Length</strong><span>metre</span></div>
        <div class="unit-card"><div class="unit-symbol">kg</div><strong>Mass</strong><span>kilogram</span></div>
        <div class="unit-card"><div class="unit-symbol">s</div><strong>Time</strong><span>second</span></div>
        <div class="unit-card"><div class="unit-symbol">A</div><strong>Current</strong><span>ampere</span></div>
        <div class="unit-card"><div class="unit-symbol">K</div><strong>Temperature</strong><span>kelvin</span></div>
        <div class="unit-card"><div class="unit-symbol">mol</div><strong>Amount</strong><span>mole</span></div>
        <div class="unit-card"><div class="unit-symbol">cd</div><strong>Intensity</strong><span>candela</span></div>
      </div>
      <div class="slide-footer"><span>SI symbols are universal.</span><span>7 base quantities</span></div>`,
  },
  {
    title: "Derived units carry meaning",
    short: "Derived units",
    time: "5 minutes",
    cue: `
      <p>Build each unit from a familiar equation.</p>
      <ul>
        <li>Speed = distance/time → m s⁻¹</li>
        <li>Acceleration = velocity/time → m s⁻²</li>
        <li>Force = mass × acceleration → kg m s⁻² = N</li>
      </ul>
      <p>Ask students to predict the unit before revealing each answer.</p>`,
    html: `
      <div class="slide-kicker">Build, don't memorise</div>
      <h1 class="slide-title">A unit can be derived from the equation.</h1>
      <div class="formula-stage">
        <div>
          <div class="formula-main">Force = mass × acceleration</div>
          <div class="formula-steps">
            <span>kg</span><strong>×</strong><span>m s⁻²</span><strong>=</strong><span>kg m s⁻²</span><strong>=</strong><span>newton (N)</span>
          </div>
        </div>
      </div>
      <div class="slide-footer"><span>Draw the cancellation steps on the board.</span><span>CBSE presentation skill</span></div>`,
  },
  {
    title: "The conversion method",
    short: "Unit conversion",
    time: "7 minutes",
    cue: `
      <p>Teach conversion as multiplication by forms of 1—not as a magic formula.</p>
      <p>Write units on every line and cancel them visibly using the whiteboard pen.</p>
      <p>Then give the speed shortcut: km/h to m/s multiply by 5/18; reverse multiply by 18/5.</p>`,
    html: `
      <div class="slide-kicker">Conversion without guesswork</div>
      <h1 class="slide-title">Multiply by a conversion factor equal to one.</h1>
      <div class="formula-stage">
        <div>
          <div class="formula-main">72 km h⁻¹ = ? m s⁻¹</div>
          <div class="formula-steps">
            <span>72 km/h</span><strong>×</strong><span>1000 m / 1 km</span><strong>×</strong><span>1 h / 3600 s</span><strong>=</strong><span>20 m/s</span>
          </div>
        </div>
      </div>
      <div class="slide-footer"><span>Units cancel like algebraic factors.</span><span>Shortcut: × 5/18</span></div>`,
  },
  {
    title: "Your turn: NEET speed check",
    short: "Student challenge",
    time: "6 minutes",
    cue: `
      <p>Give students 45 seconds without showing the answer. Ask them to hold up A, B, C or D.</p>
      <p>After voting, circle option C with the pen and ask one student to explain the conversion.</p>
      <p>Correct process: 90 × 5/18 = 25 m/s.</p>`,
    html: `
      <div class="question-prompt">
        <div class="slide-kicker" style="justify-content:center">NEET speed check · 45 seconds</div>
        <h1 class="slide-title">A car moves at 90 km h⁻¹. What is its speed in m s⁻¹?</h1>
        <div class="challenge-options">
          <div class="challenge-option">A · 15 m s⁻¹</div>
          <div class="challenge-option">B · 20 m s⁻¹</div>
          <div class="challenge-option correct">C · 25 m s⁻¹</div>
          <div class="challenge-option">D · 32.4 m s⁻¹</div>
        </div>
      </div>`,
  },
  {
    title: "Four mistakes to eliminate today",
    short: "Exam traps",
    time: "6 minutes",
    cue: `
      <p>Ask students to correct each red mistake aloud.</p>
      <p>Emphasise that CBSE marks can be lost for missing units even when the numerical value is correct.</p>
      <p>For NEET, estimation catches wrong conversion direction quickly.</p>`,
    html: `
      <div class="slide-kicker">CBSE presentation + NEET traps</div>
      <h1 class="slide-title">Four mistakes to eliminate today.</h1>
      <div class="mistake-list">
        <div class="mistake"><span class="mistake-index">1</span><span>Writing a number without a unit</span><em>20 m s⁻¹</em></div>
        <div class="mistake"><span class="mistake-index">2</span><span>Using “kgs” or “secs” as symbols</span><em>kg, s</em></div>
        <div class="mistake"><span class="mistake-index">3</span><span>Converting km/h in the wrong direction</span><em>× 5/18</em></div>
        <div class="mistake"><span class="mistake-index">4</span><span>Giving calculator digits without precision</span><em>Respect data</em></div>
      </div>
      <div class="slide-footer"><span>Correct method earns marks and prevents traps.</span><span>Exam ready</span></div>`,
  },
  {
    title: "Learning continues after class",
    short: "App feedback loop",
    time: "4 minutes",
    cue: `
      <p>This is the transition from teaching demo to product demo.</p>
      <p>Explain that the student now opens the daily five-question test. The teacher receives topic and question analytics.</p>
      <p>Close with: “Tomorrow's lesson begins with evidence, not assumptions.” Then open the Study app from the top-right button.</p>`,
    html: `
      <div class="slide-kicker">Orbit learning system</div>
      <h1 class="slide-title">The lecture ends. The learning loop does not.</h1>
      <div class="flow-row">
        <div class="flow-step"><strong>Learn</strong><span>60-minute concept lesson</span></div>
        <div class="flow-step"><strong>Attempt</strong><span>5–10 daily questions</span></div>
        <div class="flow-step"><strong>Analyse</strong><span>Topic-wise teacher data</span></div>
        <div class="flow-step"><strong>Correct</strong><span>Next-day review</span></div>
        <div class="flow-step"><strong>Continue</strong><span>Confident new topic</span></div>
      </div>
      <div class="slide-footer"><span>Next: open the student app and complete the daily test.</span><span>Demo complete</span></div>`,
  },
];

const slideLayer = document.getElementById("slide-layer");
const slideRail = document.getElementById("slide-rail");
const counter = document.getElementById("slide-counter");
const progress = document.getElementById("slide-progress");
const prevButton = document.getElementById("prev-slide");
const nextButton = document.getElementById("next-slide");
const cuePanel = document.getElementById("teacher-cues");
const cueTitle = document.getElementById("cue-title");
const cueTime = document.getElementById("cue-time");
const cueBody = document.getElementById("cue-body");
const canvas = document.getElementById("draw-canvas");
const stage = document.getElementById("stage");
const context = canvas.getContext("2d");

let currentSlide = 0;
let tool = "pointer";
let color = "#101820";
let drawing = false;
let lastPoint = null;
let slideDrawings = new Map();
let undoHistory = [];

function renderThumbnails() {
  slideRail.innerHTML = slides.map((slide, index) => `
    <button class="thumb ${index === currentSlide ? "active" : ""}" data-slide="${index}">
      <span class="thumb-number">${index + 1}</span>
      <strong>${slide.short}</strong>
      <span>${slide.title}</span>
    </button>
  `).join("");
  slideRail.querySelectorAll("[data-slide]").forEach((button) => {
    button.addEventListener("click", () => goToSlide(Number(button.dataset.slide)));
  });
  slideRail.querySelector(".thumb.active")?.scrollIntoView({ block: "nearest", behavior: "smooth" });
}

function saveDrawing() {
  if (canvas.width && canvas.height) slideDrawings.set(currentSlide, canvas.toDataURL());
}

function restoreDrawing() {
  context.clearRect(0, 0, canvas.width, canvas.height);
  undoHistory = [];
  const stored = slideDrawings.get(currentSlide);
  if (!stored) return;
  const image = new Image();
  image.onload = () => context.drawImage(image, 0, 0, canvas.width, canvas.height);
  image.src = stored;
}

function renderSlide() {
  const slide = slides[currentSlide];
  slideLayer.innerHTML = slide.html;
  counter.textContent = `${currentSlide + 1} / ${slides.length}`;
  progress.style.width = `${((currentSlide + 1) / slides.length) * 100}%`;
  prevButton.disabled = currentSlide === 0;
  nextButton.textContent = currentSlide === slides.length - 1 ? "Finish" : "Next →";
  cueTitle.textContent = slide.short;
  cueTime.textContent = slide.time;
  cueBody.innerHTML = slide.cue;
  renderThumbnails();
  restoreDrawing();
}

function goToSlide(index) {
  saveDrawing();
  currentSlide = Math.max(0, Math.min(slides.length - 1, index));
  renderSlide();
}

function resizeCanvas() {
  const previous = slideDrawings.get(currentSlide);
  const rect = stage.getBoundingClientRect();
  const ratio = window.devicePixelRatio || 1;
  canvas.width = Math.round(rect.width * ratio);
  canvas.height = Math.round(rect.height * ratio);
  canvas.style.width = `${rect.width}px`;
  canvas.style.height = `${rect.height}px`;
  context.setTransform(ratio, 0, 0, ratio, 0, 0);
  if (previous) {
    const image = new Image();
    image.onload = () => context.drawImage(image, 0, 0, rect.width, rect.height);
    image.src = previous;
  }
}

function setTool(nextTool) {
  tool = nextTool;
  canvas.classList.toggle("drawing", tool !== "pointer");
  document.querySelectorAll("[data-tool]").forEach((button) => button.classList.toggle("active", button.dataset.tool === tool));
}

function pointFromEvent(event) {
  const rect = canvas.getBoundingClientRect();
  const source = event.touches?.[0] || event;
  return { x: source.clientX - rect.left, y: source.clientY - rect.top };
}

function pushUndo() {
  undoHistory.push(canvas.toDataURL());
  if (undoHistory.length > 20) undoHistory.shift();
}

function startDrawing(event) {
  if (tool === "pointer") return;
  event.preventDefault();
  pushUndo();
  drawing = true;
  lastPoint = pointFromEvent(event);
}

function continueDrawing(event) {
  if (!drawing || tool === "pointer") return;
  event.preventDefault();
  const point = pointFromEvent(event);
  context.save();
  context.lineCap = "round";
  context.lineJoin = "round";
  context.beginPath();
  context.moveTo(lastPoint.x, lastPoint.y);
  context.lineTo(point.x, point.y);
  if (tool === "eraser") {
    context.globalCompositeOperation = "destination-out";
    context.lineWidth = 24;
    context.globalAlpha = 1;
  } else {
    context.globalCompositeOperation = "source-over";
    context.strokeStyle = tool === "highlighter" ? "#f5d84c" : color;
    context.lineWidth = tool === "highlighter" ? 18 : 3;
    context.globalAlpha = tool === "highlighter" ? .32 : 1;
  }
  context.stroke();
  context.restore();
  lastPoint = point;
}

function stopDrawing() {
  if (!drawing) return;
  drawing = false;
  lastPoint = null;
  saveDrawing();
}

function undo() {
  const previous = undoHistory.pop();
  if (!previous) return;
  context.clearRect(0, 0, canvas.width, canvas.height);
  const image = new Image();
  image.onload = () => {
    const rect = canvas.getBoundingClientRect();
    context.drawImage(image, 0, 0, rect.width, rect.height);
    saveDrawing();
  };
  image.src = previous;
}

function clearBoard() {
  if (!confirm("Clear all drawing on this slide?")) return;
  pushUndo();
  context.clearRect(0, 0, canvas.width, canvas.height);
  slideDrawings.delete(currentSlide);
}

prevButton.addEventListener("click", () => goToSlide(currentSlide - 1));
nextButton.addEventListener("click", () => {
  if (currentSlide === slides.length - 1) {
    window.location.href = "/";
  } else {
    goToSlide(currentSlide + 1);
  }
});

document.getElementById("notes-toggle").addEventListener("click", () => cuePanel.classList.toggle("open"));
document.getElementById("close-cues").addEventListener("click", () => cuePanel.classList.remove("open"));
document.getElementById("fullscreen-button").addEventListener("click", async () => {
  if (!document.fullscreenElement) {
    await document.documentElement.requestFullscreen();
  } else {
    await document.exitFullscreen();
  }
});

document.querySelectorAll("[data-tool]").forEach((button) => button.addEventListener("click", () => setTool(button.dataset.tool)));
document.querySelectorAll("[data-color]").forEach((button) => {
  button.addEventListener("click", () => {
    color = button.dataset.color;
    document.querySelectorAll("[data-color]").forEach((item) => item.classList.remove("active"));
    button.classList.add("active");
    setTool("pen");
  });
});
document.getElementById("undo-button").addEventListener("click", undo);
document.getElementById("clear-button").addEventListener("click", clearBoard);

canvas.addEventListener("mousedown", startDrawing);
canvas.addEventListener("mousemove", continueDrawing);
window.addEventListener("mouseup", stopDrawing);
canvas.addEventListener("touchstart", startDrawing, { passive: false });
canvas.addEventListener("touchmove", continueDrawing, { passive: false });
window.addEventListener("touchend", stopDrawing);

window.addEventListener("keydown", (event) => {
  if (event.key === "ArrowRight" || event.key === "PageDown") goToSlide(currentSlide + 1);
  if (event.key === "ArrowLeft" || event.key === "PageUp") goToSlide(currentSlide - 1);
  if (event.key.toLowerCase() === "p") setTool("pen");
  if (event.key.toLowerCase() === "e") setTool("eraser");
  if (event.key === "Escape") cuePanel.classList.remove("open");
});

let resizeTimer;
window.addEventListener("resize", () => {
  saveDrawing();
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(resizeCanvas, 100);
});

resizeCanvas();
renderSlide();
