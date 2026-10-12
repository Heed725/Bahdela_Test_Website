import os
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

from streamlit.testing.v1 import AppTest
from roster_store import RosterStore


class FakeTerminal:
    def __init__(self, url):
        self.base_url = url
        self.users = {}

    def create_user(self, record):
        self.users[record['employee_no']] = {'employeeNo': record['employee_no'], 'name': record['name']}

    def get_user(self, employee_id):
        return self.users.get(employee_id)

    def search_users(self):
        return list(self.users.values())

    def delete_user(self, employee_id):
        del self.users[employee_id]


class InterfaceTests(unittest.TestCase):
    def test_public_reports_and_employee_add_delete(self):
        with tempfile.TemporaryDirectory() as folder, patch.dict(os.environ, {'_APP_LAUNCHED': '1', 'ROSTER_DB_PATH': folder+'/roster.db'}):
            terminal = FakeTerminal('http://fake.test')
            with patch('employee_management.client_for', return_value=terminal):
                at = AppTest.from_file(str(Path(__file__).parents[1] / 'app.py'), default_timeout=30).run()
                self.assertFalse(at.exception)
                self.assertFalse(at.tabs)
                self.assertIn('View / Generate Reports', [b.label for b in at.button])
                self.assertIn('login_form', [f.key for f in at.get('form')])
                at.session_state['management_account'] = 'HR'
                at.session_state['management_last_seen'] = time.time()
                at.session_state['management_open'] = True
                at.run()
                self.assertFalse(at.exception)
                self.assertEqual(len(at.tabs), 4)
                at.selectbox(key='manage_add_site').select('Buguruni').run()
                at.selectbox(key='manage_add_department_Buguruni').select('Shop').run()
                at.selectbox(key='manage_add_job_Buguruni_Shop').select('SHOP')
                next(t for t in at.text_input if t.label == 'Employee ID *').set_value('98765')
                next(t for t in at.text_input if t.label == 'Full name *').set_value('UI Test Person')
                next(c for c in at.checkbox if c.label.startswith('I confirm this person')).check()
                next(b for b in at.button if b.label == 'Add employee').click().run()
                self.assertFalse(at.exception)
                self.assertIn('98765', terminal.users)
                terminal.users['99999'] = {'employeeNo': '99999', 'name': 'Unlisted Terminal Person'}
                self.assertEqual(RosterStore(folder+'/roster.db').history()[0]['action'], 'add')
                at.selectbox(key='manage_delete_site').select('Buguruni').run()
                at.button(key='manage_delete_Buguruni_refresh').click().run()
                self.assertEqual(at.selectbox(key='manage_delete_Buguruni_choice').options, ['98765 · UI Test Person · Shop · SHOP'])
                at.selectbox(key='manage_transfer_source').select('Buguruni').run()
                at.button(key='manage_transfer_Buguruni_refresh').click().run()
                self.assertEqual(at.selectbox(key='manage_transfer_Buguruni_choice').options, ['98765 · UI Test Person · Shop · SHOP'])
                at.selectbox(key='manage_delete_Buguruni_choice').select_index(0).run()
                at.text_input(key='manage_delete_confirm_Buguruni').set_value('98765').run()
                next(b for b in at.button if b.label == 'Delete employee').click().run()
                self.assertFalse(at.exception)
                self.assertNotIn('98765', terminal.users)
                self.assertEqual(RosterStore(folder+'/roster.db').history()[0]['action'], 'delete')
                next(b for b in at.button if b.label == 'Log out').click().run()
                self.assertFalse(at.exception)
                self.assertFalse(at.tabs)
                self.assertIn('View / Generate Reports', [b.label for b in at.button])

    def test_create_site_department_and_add_employee_with_job(self):
        with tempfile.TemporaryDirectory() as folder, patch.dict(os.environ, {'_APP_LAUNCHED': '1', 'ROSTER_DB_PATH': folder+'/roster.db'}):
            terminal = FakeTerminal('http://fake.test')
            with patch('employee_management.client_for', return_value=terminal):
                at = AppTest.from_file(str(Path(__file__).parents[1] / 'app.py'), default_timeout=30).run()
                at.session_state['management_account'] = 'HR'
                at.session_state['management_last_seen'] = time.time()
                at.session_state['management_open'] = True
                at.run()
                at.selectbox(key='manage_add_site').select('Buguruni').run()
                at.text_input(key='manage_add_department_Buguruni_new_name').set_value('Accounts')
                at.text_input(key='manage_add_department_Buguruni_new_job').set_value('Cashier')
                at.button(key='manage_add_department_Buguruni_create').click().run()
                self.assertFalse(at.exception)
                self.assertIn('Accounts', at.selectbox(key='manage_add_department_Buguruni').options)
                at.selectbox(key='manage_add_department_Buguruni').select('Accounts').run()
                self.assertEqual(at.selectbox(key='manage_add_job_Buguruni_Accounts').options, ['Cashier', 'Add a new job title…'])
                at.selectbox(key='manage_add_job_Buguruni_Accounts').select('Add a new job title…').run()
                at.text_input(key='manage_add_job_Buguruni_Accounts_new').set_value('Senior Cashier')
                next(t for t in at.text_input if t.label == 'Employee ID *').set_value('202')
                next(t for t in at.text_input if t.label == 'Full name *').set_value('New Department Person')
                next(c for c in at.checkbox if c.label.startswith('I confirm this person')).check()
                next(b for b in at.button if b.label == 'Add employee').click().run()
                self.assertFalse(at.exception)
                self.assertIn('202', terminal.users)
                sites = {n: {'employees': {}, 'dept_order': [], 'dept_colors': {}} for n in ['Buguruni', 'Puma Ocean Road']}
                RosterStore(folder+'/roster.db').apply(sites)
                self.assertEqual(sites['Buguruni']['employees']['Accounts']['member_jobs']['New Department Person'], 'Senior Cashier')
                self.assertNotIn('Accounts', sites['Puma Ocean Road']['employees'])
                at.selectbox(key='manage_add_site').select('Puma Ocean Road').run()
                self.assertNotIn('Accounts', at.selectbox(key='manage_add_department_Puma Ocean Road').options)
                self.assertFalse(at.exception)


if __name__ == '__main__':
    unittest.main()
