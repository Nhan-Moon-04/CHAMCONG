import sys
sys.path.insert(0, 'attendance_web')
from app.services.salary_meal_export import _get_meal_count_for_row, _get_meal_allowance_for_row

D = type('D', (), {})
detail_vp = D(); detail_vp.shift_code = 'VP'; detail_vp.shift_name = 'Van Phong'
detail_vp.notes = ''

meal_count = _get_meal_count_for_row(detail_vp, None)
meal_allowance = _get_meal_allowance_for_row(detail_vp, None)

print(f"VP shift override:")
print(f"  meal_count: {meal_count} (expect 2)")
print(f"  meal_allowance: {meal_allowance} (expect 40000)")
print(f"  total: {meal_count} x {meal_allowance} = {meal_count * meal_allowance} (expect 80000)")
