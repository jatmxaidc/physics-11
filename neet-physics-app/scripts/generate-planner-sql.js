const fs = require("fs");

const source =
  "C:/Users/DC7586/Documents/physics 11/syllabus/Class_11_Physics_Detailed_Lecture_Planner_Aug2026-Jan2027.csv";
const dest =
  "C:/Users/DC7586/Documents/physics 11/neet-physics-app/supabase/seed_lecture_planner.sql";

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cur = "";
  let quoted = false;
  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    if (quoted) {
      if (char === '"' && text[i + 1] === '"') {
        cur += '"';
        i += 1;
      } else if (char === '"') {
        quoted = false;
      } else {
        cur += char;
      }
    } else if (char === '"') {
      quoted = true;
    } else if (char === ",") {
      row.push(cur);
      cur = "";
    } else if (char === "\n") {
      row.push(cur.replace(/\r$/, ""));
      rows.push(row);
      row = [];
      cur = "";
    } else {
      cur += char;
    }
  }
  if (cur.length || row.length) {
    row.push(cur.replace(/\r$/, ""));
    rows.push(row);
  }
  return rows.filter((item) => item.some((cell) => String(cell).trim()));
}

function sqlText(value) {
  if (value == null || String(value).trim() === "") return "null";
  return `'${String(value).trim().replace(/'/g, "''")}'`;
}

const rows = parseCsv(fs.readFileSync(source, "utf8").replace(/^\uFEFF/, "")).slice(1);
const values = rows
  .map((row) => {
    const [lecture, date, day, phase, chapter, topic, delivery, homework] = row;
    const dateSql = date ? `'${date}'::date` : "null";
    return `(${[lecture, dateSql, sqlText(day), sqlText(phase), sqlText(chapter), sqlText(topic), sqlText(delivery), sqlText(homework)].join(", ")})`;
  })
  .join(",\n");

const sql = `-- Generated lecture planner seed
insert into public.lecture_planner (
  lecture_number, planned_date, day_name, phase, chapter_title, topic, delivery_focus, assessment_or_homework
) values
${values}
on conflict (lecture_number) do update set
  planned_date = excluded.planned_date,
  day_name = excluded.day_name,
  phase = excluded.phase,
  chapter_title = excluded.chapter_title,
  topic = excluded.topic,
  delivery_focus = excluded.delivery_focus,
  assessment_or_homework = excluded.assessment_or_homework;
`;

fs.writeFileSync(dest, sql);
console.log(`wrote ${rows.length} planner rows`);
