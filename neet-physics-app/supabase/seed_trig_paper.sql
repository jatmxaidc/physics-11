-- Draft CBSE trigonometry paper from Lectures 01–04. Not auto-scheduled.
-- Re-run safe: replaces only this titled draft if it has no student attempts.

do $$
declare
  v_teacher uuid;
  v_chapter uuid;
  v_class uuid;
  v_assessment uuid;
  v_question uuid;
  v_seq integer := 0;
  rec record;
begin
  select id into v_teacher from public.profiles where role = 'teacher' order by created_at limit 1;
  select id into v_chapter from public.chapters where sequence = 1 limit 1;
  select id into v_class
  from public.classes
  where lecture_number = 4
  order by created_at
  limit 1;

  if v_teacher is null or v_chapter is null then
    raise exception 'Teacher profile or Basic Maths chapter is missing';
  end if;

  select a.id into v_assessment
  from public.assessments a
  where a.title = 'CBSE Trigonometry Paper · Lectures 01–04'
  limit 1;

  if v_assessment is not null then
    if exists (select 1 from public.attempts where assessment_id = v_assessment) then
      raise notice 'Trigonometry paper already has attempts; leaving it unchanged';
      return;
    end if;
    create temporary table if not exists tmp_trig_questions (id uuid);
    delete from tmp_trig_questions;
    insert into tmp_trig_questions (id)
    select question_id from public.assessment_questions where assessment_id = v_assessment;
    delete from public.assessment_questions where assessment_id = v_assessment;
    delete from public.questions where id in (select id from tmp_trig_questions);
    delete from public.assessments where id = v_assessment;
  end if;

  insert into public.assessments (
    chapter_id, class_id, title, assessment_type, instructions,
    duration_minutes, total_marks, negative_marking, attempt_limit,
    opens_at, due_at, status, created_by
  ) values (
    v_chapter,
    v_class,
    'CBSE Trigonometry Paper · Lectures 01–04',
    'weekly_test',
    'CBSE-level objective paper from Lectures 01–04. 20 questions × 2 marks. Use radian measure unless a degree is written. The teacher must set the opening and due times before students can attempt this paper.',
    35,
    40,
    0,
    1,
    null,
    null,
    'draft',
    v_teacher
  ) returning id into v_assessment;

  for rec in
    select * from (values
      (1, 'L01 ratios', 'In a right triangle, the side opposite an acute angle θ is 8 cm and the hypotenuse is 17 cm. The value of sin θ is',
        '["8/15","8/17","15/17","17/8"]'::jsonb, 1, 'Sine is opposite divided by hypotenuse: 8/17.'),
      (2, 'L01 ratios', 'If cos θ = 12/13 and θ is acute, then tan θ equals',
        '["5/12","12/5","5/13","13/12"]'::jsonb, 0, 'The missing side is 5 by Pythagoras, so tan θ = opposite/adjacent = 5/12.'),
      (3, 'L01 identities', 'Which identity is always true for every admissible θ?',
        '["sin²θ − cos²θ = 1","sec²θ − tan²θ = 1","1 + cot²θ = sin²θ","tan²θ − sec²θ = 1"]'::jsonb, 1, 'The Pythagorean identity 1 + tan²θ = sec²θ rearranges to sec²θ − tan²θ = 1.'),
      (4, 'L01 identities', 'If sin θ = 3/5, then 1 − cos²θ equals',
        '["9/25","16/25","3/5","4/5"]'::jsonb, 0, '1 − cos²θ = sin²θ = 9/25.'),
      (5, 'L02 ASTC', 'In the second quadrant, the trigonometric ratio that remains positive is',
        '["cosine","tangent","sine","secant"]'::jsonb, 2, 'ASTC: only sine (and its reciprocal cosecant) are positive in quadrant II.'),
      (6, 'L02 ASTC', 'The angle 240° lies in which quadrant, and the sign of tan 240° is',
        '["Quadrant II, negative","Quadrant III, positive","Quadrant III, negative","Quadrant IV, positive"]'::jsonb, 1, '240° is in quadrant III, where tangent is positive.'),
      (7, 'L02 complementary', 'The complementary relation that is correct is',
        '["sin(90° − θ) = cos θ","sin(90° − θ) = sin θ","cos(90° − θ) = −sin θ","tan(90° − θ) = tan θ"]'::jsonb, 0, 'Sine of a complementary angle equals cosine of the original angle.'),
      (8, 'L02 conversion', 'The radian measure of 30° is',
        '["π/2","π/3","π/4","π/6"]'::jsonb, 3, '180° = π rad, so 30° = π/6 rad.'),
      (9, 'L02 conversion', 'π/3 radian is equal to',
        '["30°","45°","60°","90°"]'::jsonb, 2, 'π rad = 180°, therefore π/3 = 60°.'),
      (10, 'L03 radians', 'One complete revolution is',
        '["π radian","2π radian","π/2 radian","4π radian"]'::jsonb, 1, 'A full circle is 360° = 2π radian.'),
      (11, 'L03 compound', 'The expansion of sin(A + B) is',
        '["sin A cos B + cos A sin B","sin A cos B − cos A sin B","cos A cos B − sin A sin B","cos A cos B + sin A sin B"]'::jsonb, 0, 'Standard compound-angle formula for sine of a sum.'),
      (12, 'L03 compound', 'cos(A − B) expands to',
        '["cos A cos B + sin A sin B","cos A cos B − sin A sin B","sin A cos B + cos A sin B","sin A sin B − cos A cos B"]'::jsonb, 0, 'Cosine of a difference is the product of cosines plus the product of sines.'),
      (13, 'L03 double', 'sin 2θ is equal to',
        '["2 sin θ","sin²θ + cos²θ","2 sin θ cos θ","cos²θ − sin²θ"]'::jsonb, 2, 'The double-angle formula is sin 2θ = 2 sin θ cos θ.'),
      (14, 'L03 small angle', 'For a small angle θ measured in radian, the first-order Physics approximation is',
        '["sin θ ≈ 1","sin θ ≈ θ","sin θ ≈ θ²","sin θ ≈ 1/θ"]'::jsonb, 1, 'In radian measure, sin θ and tan θ both approach θ when θ is small.'),
      (15, 'L03 small angle', 'The small-angle approximation tan θ ≈ θ is valid only when',
        '["θ is in degree and large","θ is in radian and small","θ is exactly 90°","θ is any real number"]'::jsonb, 1, 'The Taylor expansion is in radian. The linear term is accurate only for small θ.'),
      (16, 'L04 max-min', 'The maximum value of 5 sin θ is',
        '["1","5","10","0"]'::jsonb, 1, 'sin θ has range [−1, 1], so 5 sin θ has maximum 5.'),
      (17, 'L04 max-min', 'The maximum value of 3 sin θ + 4 cos θ is',
        '["3","4","5","7"]'::jsonb, 2, 'The amplitude of a sin θ + b cos θ is √(a² + b²) = √(9 + 16) = 5.'),
      (18, 'L04 max-min', 'The minimum value of 2 − cos θ is',
        '["1","2","3","0"]'::jsonb, 0, 'cos θ has maximum 1, so 2 − cos θ has minimum 2 − 1 = 1.'),
      (19, 'L04 phase', 'In y = A sin(ωt + φ), the constant φ is called the',
        '["amplitude","angular frequency","phase constant","time period"]'::jsonb, 2, 'φ shifts the sine wave at t = 0 and is the phase constant.'),
      (20, 'L04 phase', 'Two oscillations of equal amplitude and frequency differ in phase by π. Their superposition has amplitude',
        '["2A","A","0","A/2"]'::jsonb, 2, 'A phase difference of π means opposite displacement at every instant, so the resultant amplitude is zero.')
    ) as q(seq, topic, prompt, options, correct, explanation)
  loop
    v_seq := rec.seq;
    insert into public.questions (
      chapter_id, class_id, prompt, question_type, options, marks, difficulty, exam_tags, topic, created_by
    ) values (
      v_chapter, v_class, rec.prompt, 'mcq', rec.options, 2, 'standard', array['CBSE'], rec.topic, v_teacher
    ) returning id into v_question;

    insert into public.question_keys (question_id, correct_answer, explanation)
    values (v_question, to_jsonb(rec.correct), rec.explanation);

    insert into public.assessment_questions (assessment_id, question_id, sequence, marks)
    values (v_assessment, v_question, v_seq, 2);
  end loop;
end;
$$;
