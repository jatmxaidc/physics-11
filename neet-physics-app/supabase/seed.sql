-- Safe curriculum seed. Run after schema.sql.
insert into public.chapters (title, sequence, description, is_locked)
values
  ('Basic Maths and Calculus', 1, 'Physics-focused algebra, trigonometry, calculus and graphs.', false),
  ('Units and Measurements', 2, 'SI units, errors, significant figures and dimensions.', true),
  ('Motion in a Straight Line', 3, 'One-dimensional kinematics and graphs.', true),
  ('Motion in a Plane', 4, 'Vectors, projectiles and circular motion.', true),
  ('Laws of Motion', 5, 'Newton’s laws, momentum and friction.', true),
  ('Work, Energy and Power', 6, 'Work-energy theorem, conservation and collisions.', true),
  ('System of Particles and Rotational Motion', 7, 'Centre of mass, torque and rotation.', true),
  ('Gravitation', 8, 'Universal gravitation, satellites and planetary motion.', true),
  ('Mechanical Properties of Solids', 9, 'Elasticity, stress and strain.', true),
  ('Mechanical Properties of Fluids', 10, 'Pressure, viscosity, flow and surface tension.', true),
  ('Thermal Properties of Matter', 11, 'Heat, expansion and heat transfer.', true),
  ('Thermodynamics', 12, 'Thermal laws and thermodynamic processes.', true),
  ('Kinetic Theory', 13, 'Molecular model of gases.', true),
  ('Oscillations', 14, 'SHM, energy and simple pendulum.', true),
  ('Waves', 15, 'Wave motion, standing waves, harmonics and beats.', true)
on conflict (sequence) do update set
  title = excluded.title,
  description = excluded.description;

-- Teacher and parent promotion must be done by a project administrator,
-- never from a public signup form:
--
-- update public.profiles set role = 'teacher' where id = '<teacher-auth-user-id>';
-- update public.profiles set role = 'parent' where id = '<parent-auth-user-id>';
-- insert into public.student_links (student_id, adult_id, relationship)
-- values ('<student-id>', '<teacher-id>', 'teacher'),
--        ('<student-id>', '<parent-id>', 'parent');
