do $$
declare
  v_item jsonb;
  v_question_id uuid;
begin
  for v_item in
    select value
    from jsonb_array_elements(
      jsonb_build_array(
        jsonb_build_object(
          'sequence', 4,
          'prompt', 'Find the centre and radius of x² + y² − 8x + 6y − 11 = 0.',
          'options', jsonb_build_array(
            'Centre (−4, 3), radius 6',
            'Centre (4, −3), radius 6',
            'Centre (4, 3), radius 36',
            'Centre (−4, −3), radius 6'
          ),
          'correctIndex', 1,
          'explanation', 'Here 2g = −8 and 2f = 6, so g = −4 and f = 3. The centre is (−g, −f) = (4, −3), and R = √(g² + f² − c) = √(16 + 9 + 11) = 6.'
        ),
        jsonb_build_object(
          'sequence', 5,
          'prompt', 'For x²/9 + y²/25 = 1, which set gives all four intercepts?',
          'options', jsonb_build_array(
            '(±9, 0) and (0, ±25)',
            '(±5, 0) and (0, ±3)',
            '(±3, 0) and (0, ±5)',
            '(±3, 0) and (0, ±25)'
          ),
          'correctIndex', 2,
          'explanation', 'Write the equation as x²/3² + y²/5² = 1. Hence the x-intercepts are (±3, 0) and the y-intercepts are (0, ±5).'
        ),
        jsonb_build_object(
          'sequence', 6,
          'prompt', 'For (x − 3)²/25 + (y + 2)²/9 = 1, identify the centre and semi-axes.',
          'options', jsonb_build_array(
            'Centre (3, −2); horizontal semi-axis 5; vertical semi-axis 3',
            'Centre (−3, 2); horizontal semi-axis 5; vertical semi-axis 3',
            'Centre (3, −2); horizontal semi-axis 25; vertical semi-axis 9',
            'Centre (3, 2); horizontal semi-axis 3; vertical semi-axis 5'
          ),
          'correctIndex', 0,
          'explanation', 'The centre is (3, −2). The semi-axes are √25 = 5 horizontally and √9 = 3 vertically.'
        ),
        jsonb_build_object(
          'sequence', 7,
          'prompt', 'For y = eˣ, what are the values at x = 0, x = 2 and x = −2, respectively?',
          'options', jsonb_build_array(
            '0, 2e, −2e',
            '1, e², 1/e²',
            '1, 1/e², e²',
            'e, 2, 1/2'
          ),
          'correctIndex', 1,
          'explanation', 'e⁰ = 1, e² remains e² and e⁻² = 1/e². Therefore the values are 1, e² and 1/e².'
        ),
        jsonb_build_object(
          'sequence', 10,
          'prompt', 'Find both roots of x² + 9x + 20 = 0 by factorisation.',
          'options', jsonb_build_array(
            'x = 4 or x = 5',
            'x = −4 or x = −5',
            'x = −2 or x = −10',
            'x = 2 or x = 10'
          ),
          'correctIndex', 1,
          'explanation', 'x² + 9x + 20 = (x + 4)(x + 5). Therefore x = −4 or x = −5.'
        )
      )
    )
  loop
    select aq.question_id
    into v_question_id
    from public.assessment_questions aq
    join public.assessments a on a.id = aq.assessment_id
    where lower(a.title) = 'lecture 09 daily quiz'
      and aq.sequence = (v_item ->> 'sequence')::integer
    limit 1;

    if v_question_id is null then
      raise exception 'Lecture 09 question % was not found', v_item ->> 'sequence';
    end if;

    update public.questions
    set
      prompt = v_item ->> 'prompt',
      options = v_item -> 'options'
    where id = v_question_id;

    update public.question_keys
    set
      correct_answer = to_jsonb((v_item ->> 'correctIndex')::integer),
      explanation = v_item ->> 'explanation'
    where question_id = v_question_id;

    v_question_id := null;
  end loop;
end
$$;
