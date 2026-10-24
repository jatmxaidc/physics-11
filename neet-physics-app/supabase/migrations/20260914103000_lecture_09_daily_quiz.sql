do $$
declare
  v_teacher_id uuid;
  v_chapter_id uuid;
  v_class_id uuid;
  v_assessment_id uuid := gen_random_uuid();
  v_question_id uuid;
  v_item jsonb;
  v_sequence integer := 0;
begin
  select p.id
  into v_teacher_id
  from public.profiles p
  join auth.users u on u.id = p.id
  where p.role = 'teacher'
  order by
    case when lower(u.email) = 'jatmanis1.mail@gmail.com' then 0 else 1 end,
    p.created_at
  limit 1;

  if v_teacher_id is null then
    raise exception 'A teacher profile is required to create the Lecture 09 quiz';
  end if;

  select id
  into v_chapter_id
  from public.chapters
  where title = 'Basic Maths and Calculus'
  limit 1;

  if v_chapter_id is null then
    raise exception 'Basic Maths and Calculus chapter was not found';
  end if;

  select id
  into v_class_id
  from public.classes
  where chapter_id = v_chapter_id and lecture_number = 9
  limit 1;

  if exists (
    select 1
    from public.assessments
    where title = 'Lecture 09 daily quiz'
      and assessment_type = 'daily_quiz'
      and due_at = timestamptz '2026-09-14 15:30:00+00'
  ) then
    raise notice 'Lecture 09 daily quiz already exists; no duplicate was created';
    return;
  end if;

  insert into public.assessments (
    id, class_id, chapter_id, title, assessment_type, instructions,
    duration_minutes, total_marks, negative_marking, attempt_limit,
    randomize_questions, randomize_options, explanations_after_submission,
    opens_at, due_at, status, created_by
  ) values (
    v_assessment_id,
    v_class_id,
    v_chapter_id,
    'Lecture 09 daily quiz',
    'daily_quiz',
    'Answer all 11 questions. Use the circle, ellipse, exponential and differentiation ideas taught in Lecture 09.',
    20,
    11,
    0,
    1,
    false,
    false,
    true,
    now(),
    timestamptz '2026-09-14 15:30:00+00',
    'published',
    v_teacher_id
  );

  for v_item in
    select value
    from jsonb_array_elements(
      jsonb_build_array(
        jsonb_build_object(
          'prompt', 'For x² + y² = 49, what are the centre and radius?',
          'topic', 'Circle',
          'options', jsonb_build_array(
            'Centre (0, 0), radius 7',
            'Centre (0, 0), radius 49',
            'Centre (7, 7), radius 7',
            'Centre (−7, 0), radius 7'
          ),
          'correctIndex', 0,
          'explanation', 'Compare x² + y² = 49 with x² + y² = R². Thus R² = 49, so R = 7, and the centre is (0, 0).'
        ),
        jsonb_build_object(
          'prompt', 'For (x − 4)² + (y − 3)² = 25, what are the centre and radius?',
          'topic', 'Circle',
          'options', jsonb_build_array(
            'Centre (−4, −3), radius 5',
            'Centre (4, 3), radius 5',
            'Centre (4, 3), radius 25',
            'Centre (3, 4), radius 5'
          ),
          'correctIndex', 1,
          'explanation', 'Compare with (x − h)² + (y − k)² = R². Here h = 4, k = 3 and R = √25 = 5.'
        ),
        jsonb_build_object(
          'prompt', 'For (x + 3)² + (y − 4)² = 36, what are the centre and radius?',
          'topic', 'Circle',
          'options', jsonb_build_array(
            'Centre (3, 4), radius 6',
            'Centre (−3, −4), radius 6',
            'Centre (−3, 4), radius 6',
            'Centre (−3, 4), radius 36'
          ),
          'correctIndex', 2,
          'explanation', 'Write x + 3 as x − (−3). Therefore the centre is (−3, 4), and R = √36 = 6.'
        ),
        jsonb_build_object(
          'prompt', 'Find the centre and radius of x² + y² − 6x + 4y − 12 = 0.',
          'topic', 'Circle',
          'options', jsonb_build_array(
            'Centre (−3, 2), radius 5',
            'Centre (3, −2), radius 5',
            'Centre (3, 2), radius 25',
            'Centre (−3, −2), radius 5'
          ),
          'correctIndex', 1,
          'explanation', 'In x² + y² + 2gx + 2fy + c = 0, g = −3 and f = 2. The centre is (−g, −f) = (3, −2), and R = √(g² + f² − c) = 5.'
        ),
        jsonb_build_object(
          'prompt', 'For x²/4 + y²/9 = 1, which set gives all four intercepts?',
          'topic', 'Ellipse',
          'options', jsonb_build_array(
            '(±4, 0) and (0, ±9)',
            '(±3, 0) and (0, ±2)',
            '(±2, 0) and (0, ±3)',
            '(±2, 0) and (0, ±9)'
          ),
          'correctIndex', 2,
          'explanation', 'Write the equation as x²/2² + y²/3² = 1. Hence the x-intercepts are (±2, 0) and the y-intercepts are (0, ±3).'
        ),
        jsonb_build_object(
          'prompt', 'For (x − 2)²/16 + (y + 1)²/4 = 1, identify the centre and semi-axes.',
          'topic', 'Ellipse',
          'options', jsonb_build_array(
            'Centre (2, −1); horizontal semi-axis 4; vertical semi-axis 2',
            'Centre (−2, 1); horizontal semi-axis 4; vertical semi-axis 2',
            'Centre (2, −1); horizontal semi-axis 16; vertical semi-axis 4',
            'Centre (2, 1); horizontal semi-axis 2; vertical semi-axis 4'
          ),
          'correctIndex', 0,
          'explanation', 'The centre is (2, −1). The semi-axes are √16 = 4 horizontally and √4 = 2 vertically.'
        ),
        jsonb_build_object(
          'prompt', 'For y = eˣ, what are the values at x = 0, x = 1 and x = −1, respectively?',
          'topic', 'Exponential functions',
          'options', jsonb_build_array(
            '0, 1, −1',
            '1, e, 1/e',
            '1, 1/e, e',
            'e, 1, 0'
          ),
          'correctIndex', 1,
          'explanation', 'e⁰ = 1, e¹ = e and e⁻¹ = 1/e. Therefore the values are 1, e and 1/e.'
        ),
        jsonb_build_object(
          'prompt', 'On the labelled Lecture 09 curve, what are the slope signs at A, B, C and D, respectively?',
          'topic', 'Slope and differentiation',
          'options', jsonb_build_array(
            'Positive, positive, zero, negative',
            'Negative, positive, zero, positive',
            'Positive, zero, negative, negative',
            'Zero, positive, positive, negative'
          ),
          'correctIndex', 0,
          'explanation', 'The curve rises at A and B, has a horizontal tangent at C, and falls at D. The signs are positive, positive, zero and negative.'
        ),
        jsonb_build_object(
          'prompt', 'For y = x² + 3x − 4, find dy/dx and the slope at x = 2.',
          'topic', 'Differentiation',
          'options', jsonb_build_array(
            'dy/dx = 2x + 3; slope = 7',
            'dy/dx = x + 3; slope = 5',
            'dy/dx = 2x − 4; slope = 0',
            'dy/dx = 2x + 7; slope = 11'
          ),
          'correctIndex', 0,
          'explanation', 'Differentiate term by term: dy/dx = 2x + 3. At x = 2, the slope is 2(2) + 3 = 7.'
        ),
        jsonb_build_object(
          'prompt', 'Find both roots of x² + 7x + 12 = 0 by factorisation.',
          'topic', 'Quadratic equations',
          'options', jsonb_build_array(
            'x = 3 or x = 4',
            'x = −3 or x = −4',
            'x = −2 or x = −6',
            'x = 2 or x = 6'
          ),
          'correctIndex', 1,
          'explanation', 'x² + 7x + 12 = (x + 3)(x + 4). Therefore x = −3 or x = −4.'
        ),
        jsonb_build_object(
          'prompt', 'Find both roots of 2x² + 3x − 2 = 0 by factorisation.',
          'topic', 'Quadratic equations',
          'options', jsonb_build_array(
            'x = 2 or x = −1/2',
            'x = −2 or x = −1/2',
            'x = 1/2 or x = −2',
            'x = 1 or x = −2'
          ),
          'correctIndex', 2,
          'explanation', '2x² + 3x − 2 = (2x − 1)(x + 2). Therefore x = 1/2 or x = −2.'
        )
      )
    )
  loop
    v_sequence := v_sequence + 1;
    v_question_id := gen_random_uuid();

    insert into public.questions (
      id, chapter_id, class_id, prompt, question_type, options, marks,
      difficulty, exam_tags, topic, created_by
    ) values (
      v_question_id,
      v_chapter_id,
      v_class_id,
      v_item ->> 'prompt',
      'mcq',
      v_item -> 'options',
      1,
      'standard',
      array['CBSE', 'NEET'],
      v_item ->> 'topic',
      v_teacher_id
    );

    insert into public.question_keys (question_id, correct_answer, explanation)
    values (
      v_question_id,
      to_jsonb((v_item ->> 'correctIndex')::integer),
      v_item ->> 'explanation'
    );

    insert into public.assessment_questions (
      assessment_id, question_id, sequence, marks
    ) values (
      v_assessment_id, v_question_id, v_sequence, 1
    );
  end loop;
end
$$;
