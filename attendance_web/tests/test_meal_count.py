from app.services.salary_meal_export import _get_meal_count_for_row
class S: pass
class D: pass
shift = S(); shift.meal_count = 0
detail = D(); detail.shift_code = 'OFF'; detail.shift_name = 'Nghi ca OFF'; detail.notes = ''
print('meal_count for OFF shift template 0 ->', _get_meal_count_for_row(detail, shift))
# NU behavior
detail2 = D(); detail2.shift_code = 'NU'; detail2.shift_name = 'ca toi'; detail2.notes = ''
print('NU night ->', _get_meal_count_for_row(detail2, None))
detail3 = D(); detail3.shift_code = 'NU'; detail3.shift_name = 'ca sang'; detail3.notes = ''
print('NU morning ->', _get_meal_count_for_row(detail3, None))
