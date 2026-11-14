import unittest
import json
import app

class FinSetAppTestCase(unittest.TestCase):
    def setUp(self):
        app.app.config['TESTING'] = True
        self.client = app.app.test_client()
        app.init_db()

    def test_health(self):
        res = self.client.get('/api/health')
        self.assertEqual(res.status_code, 200)
        data = res.get_json()
        self.assertEqual(data['status'], 'healthy')

    def test_auth_login_and_register(self):
        # 1. Login with demo user
        login_res = self.client.post('/api/auth/login',
                                     data=json.dumps({'email': 'adalinect@gmail.com', 'password': 'password123'}),
                                     content_type='application/json')
        self.assertEqual(login_res.status_code, 200)
        self.assertEqual(login_res.get_json()['user']['email'], 'adalinect@gmail.com')

        # 2. Register test user
        reg_res = self.client.post('/api/auth/register',
                                   data=json.dumps({
                                       'name': 'Test Financial Analyst',
                                       'email': 'analyst@finset.com',
                                       'password': 'securepassword'
                                   }),
                                   content_type='application/json')
        self.assertIn(reg_res.status_code, [201, 400])

    def test_monthly_excel_report(self):
        # 1. Preview
        prev_res = self.client.get('/api/report/preview?month=2026-11')
        self.assertEqual(prev_res.status_code, 200)
        data = prev_res.get_json()
        self.assertIn('total_income', data)
        self.assertIn('total_expense', data)
        self.assertIn('category_summary', data)

        # 2. Download Excel .xlsx
        excel_res = self.client.get('/api/report/excel?month=2026-11')
        self.assertEqual(excel_res.status_code, 200)
        self.assertEqual(excel_res.content_type, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        self.assertGreater(len(excel_res.data), 1000)

    def test_bills_management(self):
        # 1. List bills
        res = self.client.get('/api/bills')
        self.assertEqual(res.status_code, 200)
        bills = res.get_json()
        self.assertGreater(len(bills), 0)

        # 2. Bill summary
        sum_res = self.client.get('/api/bills/summary')
        self.assertEqual(sum_res.status_code, 200)
        s_data = sum_res.get_json()
        self.assertIn('total_bills', s_data)
        self.assertIn('total_paid', s_data)
        self.assertIn('total_pending', s_data)

        # 3. Create a bill
        new_bill_res = self.client.post('/api/bills',
                                        data=json.dumps({
                                            'name': 'Cloud Infrastructure Hosting',
                                            'amount': 45.00,
                                            'due_day': 20,
                                            'category': 'Utilities',
                                            'auto_pay': 1
                                        }),
                                        content_type='application/json')
        self.assertEqual(new_bill_res.status_code, 201)
        created = new_bill_res.get_json()

        # 4. Pay the bill
        pay_res = self.client.post(f"/api/bills/{created['id']}/pay",
                                   data=json.dumps({'record_expense': True}),
                                   content_type='application/json')
        self.assertEqual(pay_res.status_code, 200)
        self.assertEqual(pay_res.get_json()['bill']['status'], 'paid')

        # 5. Delete the bill
        del_res = self.client.delete(f"/api/bills/{created['id']}")
        self.assertEqual(del_res.status_code, 200)

if __name__ == '__main__':
    unittest.main()
