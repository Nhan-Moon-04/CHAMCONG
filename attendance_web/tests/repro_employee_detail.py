import sys
import traceback
sys.path.insert(0, 'attendance_web')

from app import create_app

app = create_app()
app.testing = True
client = app.test_client()

with client.session_transaction() as session_data:
    session_data['is_authenticated'] = True
    session_data['user_id'] = 1
    session_data['username'] = 'admin'
    session_data['is_admin'] = True

try:
    response = client.get('/employees/6?month=2026-04')
    print('status', response.status_code)
    print(response.data.decode('utf-8', 'ignore')[:2000])
except Exception:
    traceback.print_exc()
