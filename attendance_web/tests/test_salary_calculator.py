#!/usr/bin/env python
# Test salary calculator locally
import sys
sys.path.insert(0, 'attendance_web')

from app.services.salary_calculator import get_salary_month_details

# Simulate minimal test (requires DB context)
print("Salary calculator module loaded successfully")
print("Functions available:")
print("  - get_salary_month_details(employee_id, month_key)")
print("\nUsage example:")
print("  result = get_salary_month_details(1, '2026-04')")
print("  print(result['summary']['total_wage'])")
