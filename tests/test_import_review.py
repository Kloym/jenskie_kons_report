import hashlib
import json
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import zhk_dashboard as app
from zhk_import_review import ImportReview

HEADERS = ['Дата','Код услуги','Наименование услуги','Кол-во','Тариф','Врач','Подразделение','Пациент']
FIRST, LAST = date(2026,11,1), date(2026,11,17)


class ReviewTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.folder=Path(self.temp.name)
    def tearDown(self):self.temp.cleanup()
    def row(self,**changes):
        values=dict(date=FIRST,code=1983,name='Приём',quantity=2,tariff=100,doctor='Врач',department='ЖК №10')
        values.update(changes)
        return [values[x] for x in ['date','code','name','quantity','tariff','doctor','department']]+['ПАЦИЕНТ_НЕ_ПЕРЕНОСИТЬ']
    def book(self,rows):
        path=self.folder/'source.xlsx';w=Workbook();s=w.active;s.title='Лист1';s.append(HEADERS)
        for row in rows:s.append(row)
        w.save(path);w.close();return path
    def collect(self,path,resolve=None,**kwargs):
        return app.collect(path,FIRST,LAST,resolve_issues=resolve,**kwargs)

    def test_all_failures_in_one_batch_including_ambiguous_tariff(self):
        path=self.book([self.row(),self.row(tariff=120),self.row(code=None,quantity=None),
                        self.row(date='bad',tariff=-1),self.row(tariff=None),[None]*8])
        with self.assertRaises(app.ImportIssuesError) as caught:self.collect(path)
        r=caught.exception.review
        self.assertEqual(set(r.issues),{4,5,6})
        self.assertEqual(set(r.issues[4]),{'code','quantity'})
        self.assertEqual(set(r.issues[5]),{'date','tariff'})
        self.assertIn('tariff',r.issues[6])
        self.assertNotIn('ПАЦИЕНТ_НЕ_ПЕРЕНОСИТЬ',str(r.rows))

    def test_repair_and_exclusion_feed_both_dashboard_and_annual_actuals(self):
        path=self.book([self.row(),self.row(code=None,quantity=None),self.row(tariff=-2)])
        before=hashlib.sha256(path.read_bytes()).hexdigest();calls=[]
        def resolve(r):
            calls.append(set(r.issues));r.edit(3,{**r.values(3),'code':'001984','quantity':'3','tariff':'50,25'})
            r.exclude([4]);return True
        with patch.object(app,'_open_source',wraps=app._open_source) as reader:
            data=self.collect(path,resolve,actuals_from_services=True)
        self.assertEqual(reader.call_count,1)
        self.assertEqual(calls,[{3,4}])
        self.assertEqual(data['meta']['totalQuantity'],5)
        self.assertEqual(data['meta']['totalCents'],35075)
        self.assertEqual(sum(r['quantity'] for r in data['serviceActuals']['rows']),5)
        self.assertEqual(data['meta']['stats']['edited_rows'],1)
        self.assertEqual(data['meta']['stats']['excluded_by_user_rows'],1)
        self.assertEqual(before,hashlib.sha256(path.read_bytes()).hexdigest())
        self.assertNotIn('ПАЦИЕНТ_',json.dumps(data,ensure_ascii=False))
        self.assertEqual(data['meta']['importReview']['decisions'][0],{'row':3,'action':'edited','fields':['code','quantity','tariff']})

    def test_cancel_and_unfixed_acceptance_do_not_return_partial_report(self):
        path=self.book([self.row(code=None)])
        with self.assertRaises(app.ImportCancelled):self.collect(path,lambda r:False)
        with self.assertRaises(app.ImportIssuesError):self.collect(path,lambda r:True)

    def test_blank_rows_automatic_but_partial_rows_require_decision(self):
        path=self.book([[None]*8,self.row(),[FIRST,None,None,None,None,None,None,None]])
        def resolve(r):
            self.assertEqual(set(r.issues),{4});r.exclude([4]);return True
        data=self.collect(path,resolve)
        self.assertEqual(data['meta']['stats']['blank_rows'],1)
        self.assertEqual(data['meta']['totalQuantity'],2)

    def test_revalidate_price_donors_after_edit_and_exclusion(self):
        path=self.book([self.row(tariff=100),self.row(tariff=200),self.row(tariff=None)])
        with self.assertRaises(app.ImportIssuesError) as caught:self.collect(path)
        r=caught.exception.review;r.exclude([3]);self.assertTrue(r.evaluate())
        self.assertEqual(r.prepared[-1].price,100)
        r.restore([3]);self.assertFalse(r.evaluate());self.assertIn('tariff',r.issues[4])
        r.edit(4,{**r.values(4),'tariff':'150'});self.assertTrue(r.evaluate())

    def test_changes_can_be_undone_and_zero_is_not_missing(self):
        path=self.book([self.row(code=None),self.row(tariff=0)])
        with self.assertRaises(app.ImportIssuesError) as caught:self.collect(path)
        r=caught.exception.review;r.edit(2,{**r.values(2),'code':'1983','quantity':'0','tariff':''})
        self.assertTrue(r.evaluate());self.assertEqual(r.prepared[0].price,0)
        r.exclude([2]);r.restore([2]);self.assertFalse(r.evaluate())
        self.assertNotIn(2,r.overrides);self.assertNotIn(2,r.excluded)

    def test_actuals_before_visible_period_are_validated_and_corrected(self):
        path=self.book([self.row(quantity='bad'),self.row(date=LAST,quantity=1)])
        def resolve(r):
            self.assertEqual(set(r.issues),{2});r.edit(2,{**r.values(2),'quantity':'4'});return True
        data=app.collect(path,LAST,LAST,actuals_from_services=True,resolve_issues=resolve)
        self.assertEqual(data['meta']['totalQuantity'],1)
        self.assertEqual(sum(r['quantity'] for r in data['serviceActuals']['rows']),5)

    def test_date_and_department_correction_changes_filtering(self):
        path=self.book([self.row(date='bad',department=None),self.row()])
        def resolve(r):
            r.edit(2,{**r.values(2),'date':'17.11.2026','department':'ЦЖЗ на Лобненской'});return True
        data=self.collect(path,resolve)
        self.assertEqual(data['meta']['totalQuantity'],4)
        self.assertIn('ЦЖЗ на Лобненской',data['departments'])
        self.assertEqual(data['meta']['sourceEnd'],'2026-11-17')

    def test_excluding_everything_does_not_create_zero_report(self):
        path=self.book([self.row(code=None)])
        def resolve(r):r.exclude(r.rows);return True
        with self.assertRaisesRegex(app.DashboardError,'нет услуг'):self.collect(path,resolve)

    def test_review_metadata_visible_in_generated_report(self):
        path=self.book([self.row(code=None),self.row()])
        def resolve(r):r.exclude([2]);return True
        data=self.collect(path,resolve)
        target=app.write_dashboard(data,self.folder/'report.html',include_plan=False)
        html=target.read_text(encoding='utf-8')
        self.assertIn('Исключено пользователем',html)
        self.assertIn('"action":"excluded"',html)
        self.assertNotIn('ПАЦИЕНТ_',html)


if __name__=='__main__':unittest.main()
