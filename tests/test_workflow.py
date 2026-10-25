"""Public, synthetic end-to-end checks. No original research datasets required."""
import copy
import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from decimal import Decimal
from zipfile import ZipFile

ROOT=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('sample',ROOT/'examples/make_sample.py')
sample=importlib.util.module_from_spec(spec)
spec.loader.exec_module(sample)

class Workflow(unittest.TestCase):
    def run_typed_rows(self, rows, configure, boolean_cells=()):
        with tempfile.TemporaryDirectory() as td:
            d = Path(td); source = d / 'typed.xlsx'
            sample.workbook(source, rows)
            if boolean_cells:
                with ZipFile(source) as archive:
                    members = {name: archive.read(name) for name in archive.namelist()}
                name = 'xl/worksheets/sheet1.xml'
                xml = members[name].decode()
                for ref in boolean_cells:
                    xml = xml.replace(f'<c r="{ref}">', f'<c r="{ref}" t="b">')
                members[name] = xml.encode()
                with ZipFile(source, 'w') as archive:
                    for name, content in members.items():
                        archive.writestr(name, content)
            plan = sample.plan(str(source), hashlib.sha256(source.read_bytes()).hexdigest())
            configure(plan)
            pp = d / 'plan.json'; pp.write_text(json.dumps(plan))
            out = d / 'result'
            cmd = [sys.executable, str(ROOT / 'src/table_analysis.py')]
            run = subprocess.run(cmd + ['run', str(pp), '--out', str(out)], capture_output=True, text=True)
            self.assertEqual(run.returncode, 0, run.stderr)
            result = json.loads((out / 'result.json').read_text())
            verify = subprocess.run(cmd + ['verify', str(out / 'result.json'), '--source', str(source)], capture_output=True, text=True)
            self.assertEqual(verify.returncode, 0, verify.stderr)
            return result

    def test_distinct_numbers_ignore_decimal_spelling_but_preserve_types(self):
        def configure(plan):
            plan['group_by'] = []
            plan['metrics'] = [{'id':'distinct','label':'不同值','op':'count_distinct',
                                'expr':{'field':'Channel'},'missing':'skip','unit':'count','scope':'group'}]
        rows = [['S1', Decimal('1'), 0], ['S2', Decimal('1.0'), 0],
                ['S3', Decimal('1E+0'), 0], ['S4', '1', 0], ['S5', 1, 0]]
        result = self.run_typed_rows(rows, configure, boolean_cells=['B6'])
        value = result['results'][0]['values']['distinct']
        self.assertEqual(value['value'], '3')
        self.assertEqual(value['evidence']['valid_count'], 5)

    def test_grouping_keeps_boolean_numeric_and_text_values_separate(self):
        def configure(plan):
            plan['metrics'] = [plan['metrics'][0]]
        rows = [['S1', Decimal('1'), 0], ['S2', Decimal('1.0'), 0],
                ['S3', 1, 0], ['S4', '1', 0], ['S5', 0, 0], ['S6', 0, 0]]
        result = self.run_typed_rows(rows, configure, boolean_cells=['B4', 'B7'])
        self.assertEqual(result['population']['total_groups'], 5)
        self.assertEqual([g['source_rows'] for g in result['results']], [[2,3],[4],[5],[6],[7]])
        self.assertEqual([g['values']['orders']['value'] for g in result['results']], ['2','1','1','1','1'])

    def test_workflow_and_failure_boundaries(self):
        with tempfile.TemporaryDirectory() as td:
            d=Path(td); source=d/'orders.xlsx'; sample.workbook(source)
            def call(*args, code=0):
                p=subprocess.run([sys.executable,str(ROOT/'src/table_analysis.py'),*map(str,args)],capture_output=True,text=True)
                self.assertEqual(p.returncode,code,p.stderr)
                return json.loads(p.stdout if code==0 else p.stderr)
            def save(plan,name='plan.json'):
                path=d/name;path.write_text(json.dumps(plan,ensure_ascii=False));return path
            sha=hashlib.sha256(source.read_bytes()).hexdigest()
            plan=sample.plan(str(source),sha); pp=save(plan)
            self.assertEqual(call('inspect',source)['tables'][0]['name'],'OrdersTable')
            out=d/'result';call('run',pp,'--out',out)
            result=json.loads((out/'result.json').read_text())
            groups={g['group']['Channel']:g['values'] for g in result['results']}
            self.assertEqual(groups['Web']['total']['value'],'400')
            self.assertEqual(groups['Store']['average']['value'],'100')
            self.assertEqual(groups['Store']['orders']['value'],'3')
            evidence=groups['Store']['average']['evidence']
            self.assertEqual(evidence['valid_count'],2)
            self.assertEqual(evidence['missing_rows'],[5])
            self.assertEqual([c['cells'] for c in evidence['contributions']],[['C4'],['C6']])
            self.assertTrue(call('verify',out/'result.json','--source',source)['recomputed'])
            self.assertEqual(call('run',pp,'--out',out,code=2)['code'],'OUTPUT_EXISTS')
            bad=copy.deepcopy(plan);bad['interpretation']['clarifications']=[{'question':'范围？','answer':''}]
            self.assertEqual(call('run',save(bad,'unresolved.json'),'--out',d/'bad',code=2)['code'],'NEEDS_CLARIFICATION')
            bad=copy.deepcopy(plan);bad['metrics'][2]['missing']='error'
            self.assertEqual(call('run',save(bad,'missing.json'),'--out',d/'bad',code=2)['code'],'MISSING_VALUE')
            bad=copy.deepcopy(plan);bad['question']='![remote](https://example.invalid/x)\n<script>alert(1)</script>'
            bad['metrics'][0]['label']=bad['question']
            call('run',save(bad,'text.json'),'--out',d/'text')
            md=(d/'text/answer.md').read_text()
            self.assertNotIn('![',md);self.assertNotIn('<script>',md);self.assertNotIn('https://',md)
            changed=copy.deepcopy(sample.ROWS);changed[0][2]=999
            sample.workbook(source,changed)
            self.assertEqual(call('verify',out/'result.json','--source',source,code=2)['code'],'SOURCE_CHANGED')
            self.assertEqual(call('run',pp,'--out',d/'stale',code=2)['code'],'SOURCE_CHANGED')
            self.assertFalse((d/'stale').exists())

if __name__=='__main__': unittest.main()

