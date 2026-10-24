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
    raise exception 'A teacher profile is required to create the Lecture 10 quiz';
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
  where chapter_id = v_chapter_id and lecture_number = 10
  limit 1;

  if exists (
    select 1
    from public.assessments
    where lower(title) = 'lecture 10 daily quiz'
      and assessment_type = 'daily_quiz'
  ) then
    raise notice 'Lecture 10 daily quiz already exists; no duplicate was created';
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
    'Lecture 10 daily quiz',
    'daily_quiz',
    'Answer all 12 questions from Lecture 10. Each question carries one mark. You may resubmit before the deadline.',
    25,
    12,
    0,
    1,
    false,
    false,
    true,
    now(),
    timestamptz '2026-09-17 18:29:00+00',
    'published',
    v_teacher_id
  );

  for v_item in
    select value
    from jsonb_array_elements(
      jsonb_build_array(
        jsonb_build_object(
          'prompt', 'If y = 2x + 4, find the change in y and the average rate of change when x changes from 1 to 3.',
          'topic', 'Rate of change',
          'options', jsonb_build_array(
            'Change in y = 2; average rate = 1',
            'Change in y = 4; average rate = 2',
            'Change in y = 6; average rate = 3',
            'Change in y = 8; average rate = 4'
          ),
          'correctIndex', 1,
          'explanation', 'y(1) = 6 and y(3) = 10. Therefore Δy = 10 − 6 = 4, while Δx = 2. The average rate is Δy/Δx = 4/2 = 2.'
        ),
        jsonb_build_object(
          'prompt', 'Find both roots of x² − 7x + 12 = 0 by factorisation.',
          'topic', 'Quadratic equations',
          'options', jsonb_build_array(
            'x = 3 or x = 4',
            'x = −3 or x = −4',
            'x = 2 or x = 6',
            'x = −2 or x = −6'
          ),
          'correctIndex', 0,
          'explanation', 'x² − 7x + 12 = (x − 3)(x − 4). Hence x = 3 or x = 4.'
        ),
        jsonb_build_object(
          'prompt', 'Find both roots of x² + 7x − 18 = 0 by factorisation.',
          'topic', 'Quadratic equations',
          'options', jsonb_build_array(
            'x = 9 or x = −2',
            'x = 2 or x = −9',
            'x = 3 or x = −6',
            'x = −3 or x = 6'
          ),
          'correctIndex', 1,
          'explanation', 'x² + 7x − 18 = (x + 9)(x − 2). Hence x = 2 or x = −9.'
        ),
        jsonb_build_object(
          'prompt', 'Find both roots of x² + 8x + 15 = 0 by factorisation.',
          'topic', 'Quadratic equations',
          'options', jsonb_build_array(
            'x = 3 or x = 5',
            'x = −3 or x = −5',
            'x = −1 or x = −15',
            'x = 1 or x = 15'
          ),
          'correctIndex', 1,
          'explanation', 'x² + 8x + 15 = (x + 3)(x + 5). Hence x = −3 or x = −5.'
        ),
        jsonb_build_object(
          'prompt', 'A ball is thrown vertically downward at 10 m s⁻¹ from a height of 75 m. Taking g = 10 m s⁻², find its time of flight.',
          'topic', 'Quadratic equations in physics',
          'options', jsonb_build_array(
            '2 s',
            '3 s',
            '5 s',
            '7.5 s'
          ),
          'correctIndex', 1,
          'explanation', 'Use s = ut + ½gt²: 75 = 10t + 5t². Thus t² + 2t − 15 = 0 = (t + 5)(t − 3). Physical time is t = 3 s.'
        ),
        jsonb_build_object(
          'prompt', 'For y = 4x³ − 5x² + 2x − 7, find dy/dx.',
          'topic', 'Differentiation',
          'options', jsonb_build_array(
            '12x² − 10x + 2',
            '12x² − 5x + 2',
            '4x² − 10x + 2',
            '12x³ − 10x² + 2'
          ),
          'correctIndex', 0,
          'explanation', 'Differentiate term by term: d(4x³)/dx = 12x², d(−5x²)/dx = −10x, d(2x)/dx = 2 and the constant gives zero.'
        ),
        jsonb_build_object(
          'prompt', 'Differentiate y = tan x · logₑx.',
          'topic', 'Product rule',
          'options', jsonb_build_array(
            'sec²x/x',
            'tan x/x + logₑx · sec²x',
            'tan x · sec²x/x',
            'logₑx/x + tan x · sec²x'
          ),
          'correctIndex', 1,
          'explanation', 'Using (AB)′ = AB′ + BA′, dy/dx = tan x(1/x) + logₑx(sec²x).'
        ),
        jsonb_build_object(
          'prompt', 'If y = A/x³ − B/x⁴, find x for which dy/dx = 0. Assume A and B are non-zero constants.',
          'topic', 'Stationary points',
          'options', jsonb_build_array(
            'x = 3A/(4B)',
            'x = 4A/(3B)',
            'x = 4B/(3A)',
            'x = 3B/(4A)'
          ),
          'correctIndex', 2,
          'explanation', 'Write y = Ax⁻³ − Bx⁻⁴. Then dy/dx = −3Ax⁻⁴ + 4Bx⁻⁵. Setting it to zero gives −3Ax + 4B = 0, so x = 4B/(3A).'
        ),
        jsonb_build_object(
          'prompt', 'Differentiate y = eˣ/x², where x ≠ 0.',
          'topic', 'Quotient rule',
          'options', jsonb_build_array(
            'eˣ(x − 2)/x³',
            'eˣ(x + 2)/x³',
            'eˣ(x − 2)/x²',
            'eˣ/x³'
          ),
          'correctIndex', 0,
          'explanation', 'By the quotient rule, dy/dx = [x²eˣ − eˣ(2x)]/x⁴ = eˣ(x − 2)/x³.'
        ),
        jsonb_build_object(
          'prompt', 'Differentiate y = (x² − 1)/(x² + 1).',
          'topic', 'Quotient rule',
          'options', jsonb_build_array(
            '2x/(x² + 1)',
            '4x/(x² + 1)²',
            '4x²/(x² + 1)²',
            '−4x/(x² + 1)²'
          ),
          'correctIndex', 1,
          'explanation', 'The quotient rule gives [(x² + 1)(2x) − (x² − 1)(2x)]/(x² + 1)². The numerator simplifies to 4x.'
        ),
        jsonb_build_object(
          'prompt', 'Differentiate y = cos x + 2/x³, where x ≠ 0.',
          'topic', 'Mixed differentiation',
          'options', jsonb_build_array(
            'sin x + 6/x⁴',
            '−sin x + 6/x⁴',
            '−sin x − 6/x⁴',
            'cos x − 6/x³'
          ),
          'correctIndex', 2,
          'explanation', 'Write 2/x³ as 2x⁻³. Then d(cos x)/dx = −sin x and d(2x⁻³)/dx = −6x⁻⁴.'
        ),
        jsonb_build_object(
          'prompt', 'Differentiate y = (√x + 1/√x)² for x > 0.',
          'topic', 'Simplification before differentiation',
          'options', jsonb_build_array(
            '1 + 1/x²',
            '1 − 1/x²',
            '2 − 2/x²',
            '1 − 2/x'
          ),
          'correctIndex', 1,
          'explanation', 'First expand: y = x + 2 + 1/x. Therefore dy/dx = 1 − 1/x².'
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
