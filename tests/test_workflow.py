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

ROOT=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('sample',ROOT/'examples/make_sample.py')
sample=importlib.util.module_from_spec(spec)
spec.loader.exec_module(sample)

class Workflow(unittest.TestCase):
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
