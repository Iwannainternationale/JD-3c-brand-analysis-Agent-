import csv
import importlib.util
import json
from pathlib import Path
import re
import subprocess
import sys
import tempfile
import unittest
from decimal import Decimal

from brand_agent import analyze, InputError, report, risk_list

ROOT = Path(__file__).resolve().parents[1]
EX = ROOT/'examples'


class RulesTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.p = Path(self.tmp.name)
        self.gmv = self.p/'g.csv'
        self.keys = self.p/'k.csv'
        self.share = self.p/'w.csv'
        self.save(self.keys, ['brand'], [['A'],['B'],['C'],['D'],['E']])
        self.save(self.share, ['week','brand','share'], [
            [w, b, v] for w, values in zip(['2025-W51', '2025年第52周', '2026-W01'],
                                          [[.25,.1,.15,.2], [.21,None,.15,.18], [.19,0,.1,.13]])
            for b,v in zip(['A','B','C','F'],values)
        ])
        self.save(self.gmv, ['brand','category','current_gmv','prior_year_gmv'], [
            ('A','one',85,100), ('A','one',5,None), ('B','one',0,10), ('C','one',29,30),
            ('D','one',5,5), ('E','one',3,None), ('F','one',10,15), ('G','one',2,-9),
            ('H','one',90,100),('H','two',30,None),('I','one',5,8),
        ])

    def save(self, path, header, rows):
        with path.open('w', encoding='utf-8', newline='') as file:
            w = csv.writer(file)
            w.writerow(header)
            w.writerows(rows)

    def run_analysis(self, **kwargs):
        return analyze(self.gmv,self.keys,self.share,period='2025-12',source='样例',**kwargs)

    def test_rules(self):
        a = self.run_analysis(expected_latest='2026-W01')
        rows = {r['brand']:r for r in a['rows']}
        self.assertEqual(a['weeks'], ['2025-W51','2025-W52','2026-W01'])
        self.assertEqual(a['profile']['repeated_brand_category_rows'], 1)
        self.assertEqual(a['profile']['blank_prior_rows'], 3)
        self.assertEqual((rows['A']['label'], rows['A']['yoy']), ('风险',Decimal('-.1')))
        self.assertEqual((rows['B']['label'], rows['B']['yoy'], rows['B']['shares'][-1]), ('预警',Decimal(-1),Decimal(0)))
        self.assertIn('缺数', rows['B']['reason'])
        self.assertIn('未持续', rows['C']['reason'])
        self.assertEqual((rows['D']['kind'],rows['E']['kind'],rows['G']['kind']), ('持平','无基数','不可比'))
        self.assertEqual(rows['F']['label'], '次要')
        self.assertEqual(rows['H']['yoy'], Decimal('.2'))
        self.assertEqual(rows['I']['label'], '次要')
        self.assertEqual(len(risk_list(a)['brands']), 5)

    def test_invalid_windows(self):
        with self.assertRaisesRegex(InputError,'最新周不符'):
            self.run_analysis(expected_latest='2026-W02')
        self.save(self.share, ['week','brand','share'], [[w,'A','.2'] for w in ['2025-W51','2025-W52','2026-W02']])
        with self.assertRaisesRegex(InputError,'不是连续三周'):
            self.run_analysis()

    def test_invalid_values_and_missing_prior_column(self):
        self.save(self.gmv, ['brand','category','current_gmv'], [['A','one',23]])
        with self.assertRaisesRegex(InputError,'prior_year_gmv'):
            self.run_analysis()
        self.save(self.gmv, ['brand','category','current_gmv','prior_year_gmv'], [['A','one',23,'NA']])
        with self.assertRaisesRegex(InputError,'数值格式'):
            self.run_analysis()
        self.save(self.gmv, ['brand','category','current_gmv','prior_year_gmv'], [['A','one','',23]])
        with self.assertRaisesRegex(InputError,'本期 GMV 为空'):
            self.run_analysis()

    def test_week_precision_cannot_round_away_the_trend(self):
        samples = [['2025-W51','A','.2000000000001'],['2025-W52','A','.2'],['2026-W01','A','.199999999999']]
        self.save(self.share, ['week','brand','share'], samples)
        with self.assertRaisesRegex(InputError,'小数最多保留'):
            self.run_analysis()
        samples[0][2] = '.200000000001'
        self.save(self.share, ['week','brand','share'], samples)
        result = self.run_analysis()
        self.assertEqual(risk_list(result)['brands'][0]['label'], '风险')
        self.assertIn('20.0000000001%', report(result))

    def test_normalized_weekly_duplicate_fails(self):
        with self.share.open('a', encoding='utf-8') as f:
            f.write('2026-W01,Ａ,0.2\n')
        with self.assertRaisesRegex(InputError,'同品牌同周'):
            self.run_analysis()

    def test_alias_collision(self):
        alias = self.p/'alias.json'
        alias.write_text('{"A":"One","Ａ":"Two"}',encoding='utf-8')
        with self.assertRaisesRegex(InputError,'重复别名'):
            self.run_analysis(alias_path=alias)

    def test_deep_cards_selected_by_gmv_not_label(self):
        self.save(self.gmv, ['brand','category','current_gmv','prior_year_gmv'],
                  [('A','one',1,2)]+[(f'large{i:03}','one',100+i,200+i) for i in range(105)])
        a = self.run_analysis()
        md = report(a)
        self.assertEqual(len([l for l in md.splitlines() if l.startswith('### ')]), 30)
        self.assertNotIn('### A（风险）', md)
        self.assertIn('未展示 6 个',md)
        self.assertEqual(len(risk_list(a)['brands']),106)

    @unittest.skipUnless(importlib.util.find_spec('openpyxl'), 'optional XLSX dependency not installed')
    def test_xlsx_contract(self):
        from openpyxl import Workbook
        p = self.p/'source.xlsx'
        wb = Workbook()
        wb.active.append(['not input'])
        sheet = wb.create_sheet('data')
        sheet.append(['brand','category','current_gmv','prior_year_gmv'])
        sheet.append(['A','sample',0,100]);sheet.append(['E','sample',30,None]);wb.save(p);wb.close()
        self.gmv = p
        self.assertEqual([r['kind'] for r in self.run_analysis()['rows']],['无基数','负增长'])


class ReleaseTest(unittest.TestCase):
    def test_cli_matches_examples_and_dashboard(self):
        with tempfile.TemporaryDirectory() as tmp:
            proc = subprocess.run([sys.executable,str(ROOT/'brand_agent.py'),'--gmv',str(EX/'gmv.csv'),'--key-brands',str(EX/'key-brands.csv'),
                                    '--weekly-share',str(EX/'weekly-share.csv'),'--aliases',str(EX/'aliases.json'),'--period','2025-12（模拟）',
                                    '--share-source','本地模拟周市占','--expected-latest-week','2026-W01','--demo','--out',tmp],cwd=ROOT,capture_output=True)
            self.assertEqual(proc.returncode,0,proc.stderr.decode('utf-8','replace'))
            for name in ('report.md','brand-risk-list.json'):
                self.assertEqual((Path(tmp)/name).read_text(encoding='utf-8'), (EX/'expected'/name).read_text(encoding='utf-8'))
            data = json.loads((Path(tmp)/'brand-risk-list.json').read_text(encoding='utf-8'))
            self.assertEqual(len(data['brands']), 12)
            self.assertEqual(sum(b['label']=='风险' for b in data['brands']),5)
            html = (ROOT/'dashboard'/'index.html').read_text(encoding='utf-8')
            inline = re.search(r'<script type="application/json" id="sampleData">(.*?)</script>',html).group(1)
            self.assertEqual(json.loads(inline),data)
            monthly = json.loads((EX/'monthly-demo.json').read_text(encoding='utf-8'))
            self.assertEqual(len(monthly['months']),22)
            self.assertEqual(len({r[1] for r in monthly['rows']}),16)


if __name__ == '__main__':
    unittest.main()
