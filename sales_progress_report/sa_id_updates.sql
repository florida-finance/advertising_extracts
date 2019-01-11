UPDATE public.sales_employees
  SET rep_employee_id='113329', rep_name='Covatta,Stephanie T.', manager_employee_id='082647', digital_strategist_employee_id='210159'
  WHERE sa_id = 'SA00269';

INSERT INTO public.sales_employees(
  sa_id, rep_name, rep_employee_id, manager_employee_id, digital_strategist_employee_id)
  VALUES ('SA03830', 'Covatta,Stephanie T.', '113329', '082647', '210159');

-- SELECT *
-- FROM employee_goals
-- WHERE sa_id = 'SA03830';
