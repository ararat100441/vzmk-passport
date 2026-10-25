"""Simplified RC3 Back/Generate contract (all synthetic, no real registry).

Back saves the current window *without certifying* any inspection. Generate
prepares numeric values and lets the existing strict engine check evidence.
"""
import inspect
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import patch

from production_app import gui,core
from production_app.facts import validate as validate_facts
from production_app.storage import read
from production_app.simple_facts import save_screen_draft,prepare_for_check
from production_app.storage import Stop,save


MARKS=[
    {'mark':'Бп2-1','profile':'[16П','akz':'0','uzk_quantity':'8'},
    {'mark':'Бп2-5','profile':'[16П','akz':'','uzk_quantity':'0'},
]
HIDDEN={
    'Бп2-1':{'paint':'Грунт-1'},
    'Бп2-5':{},
}
VALID_AKZ={
    'coats':1,
    'air':{'min':'23','max':'27'},
    'surface':{'min':'20','max':'22'},
    'humidity':{'min':'50','max':'60'},
    'layer_intervals':[{'layer':1,'min':'90','max':'130'}],
}


class BackDraft(unittest.TestCase):
    def save(self,previous=None,marks=None,climates=None):
        return save_screen_draft(
            previous or {},'Гараж Казахстан','11/2026',False,
            marks if marks is not None else MARKS,
            HIDDEN,climates if climates is not None else {'0':VALID_AKZ},
        )

    def test_partial_climate_and_absent_proofs_save_without_raising(self):
        draft=self.save(climates={'0':{
            'coats':1,'air':{'min':'22','max':''},
            'surface':{'min':'','max':''},
            'humidity':{'min':'','max':''},
            'layer_intervals':[{'layer':1,'min':'','max':''}],
        }})
        self.assertEqual(draft['akz_drafts']['0']['air']['min'],'22')
        self.assertEqual(draft['marks'][0]['uzk_quantity'],'8')
        self.assertFalse(draft['confirmed'])
        self.assertFalse(draft['authorized'])
        self.assertFalse(draft['forms_approved'])
        self.assertEqual(draft['climates'],[])
        with self.assertRaises(Stop):
            prepare_for_check(draft)

    def test_back_does_not_write_files(self):
        with TemporaryDirectory() as folder:
            root=Path(folder)
            marker=root/'untouched.xlsx'
            marker.write_bytes(b'KEEP')
            self.save(climates={'0':VALID_AKZ})
            self.assertEqual(marker.read_bytes(),b'KEEP')
            self.assertEqual(list(root.iterdir()),[marker])

    def test_existing_verified_fields_preserved_but_consent_reset(self):
        old={
            'marks':[
                {'mark':'Бп2-1','uzk_result':'ГОДЕН',
                 'vik_result':'ГОДЕН','basis':'CURRENT-2026'},
                {'mark':'Бп2-5'},
            ],
            'evidence':[{'path':'current-evidence.pdf','sha256':'123'}],
            'doc_fields':{'QUALITY_ORDER':'CURRENT'},
            'confirmed':True,'authorized':True,'forms_approved':True,
        }
        draft=self.save(old)
        self.assertEqual(draft['marks'][0]['uzk_result'],'ГОДЕН')
        self.assertEqual(draft['marks'][0]['vik_result'],'ГОДЕН')
        self.assertEqual(draft['doc_fields']['QUALITY_ORDER'],'CURRENT')
        self.assertEqual(len(draft['evidence']),1)
        self.assertFalse(draft['confirmed'])
        self.assertFalse(draft['authorized'])
        self.assertFalse(draft['forms_approved'])

    def test_back_does_not_synthesize_control_verdicts(self):
        draft=self.save()
        self.assertNotIn('uzk_result',draft['marks'][0])
        self.assertNotIn('vik_result',draft['marks'][0])
        self.assertNotIn('basis',draft['marks'][0])
        self.assertNotIn('evidence',draft)

    def test_mark_deduplication(self):
        with self.assertRaisesRegex(Stop,'дублируется'):
            self.save(marks=[MARKS[0],MARKS[0]])

    def test_valid_generate_uses_current_measured_ranges(self):
        draft=self.save()
        prepared=prepare_for_check(draft)
        self.assertEqual(prepared['marks'][0]['akz'],0)
        self.assertIsNone(prepared['marks'][1]['akz'])
        self.assertEqual(len(prepared['climates']),1)
        actual=prepared['climates'][0]
        self.assertEqual(actual['variant'],0)
        self.assertEqual(actual['air'],'25')
        self.assertEqual(actual['surface'],'21')
        self.assertEqual(actual['humidity'],'55')
        self.assertEqual(actual['dew'],'15.3')
        self.assertEqual(actual['thickness'],'110')
        self.assertEqual(actual['material'],'Грунт-1')
        self.assertEqual(actual['basis'],'')
        self.assertFalse(prepared['confirmed'])
        self.assertFalse(prepared['authorized'])
        self.assertFalse(prepared['forms_approved'])

    def test_invalid_akz_and_missing_variant_are_rejected_at_generate(self):
        marks=[dict(MARKS[0],akz='8'),MARKS[1]]
        with self.assertRaisesRegex(Stop,'некорректный № АКЗ'):
            prepare_for_check(self.save(marks=marks))
        with self.assertRaisesRegex(Stop,'Не введён климат'):
            prepare_for_check(self.save(climates={}))

    def test_old_current_akz_proof_is_preserved_without_faking_new_one(self):
        proof={'path':'/synthetic/akz-current.pdf','sha256':'ABC',
               'role':'AKZ','variant':0}
        result=prepare_for_check(self.save({
            'akz_proofs':{'0':[proof]},
        }))
        self.assertEqual(result['climates'][0]['basis'],'akz-current.pdf')

    def test_draft_numeric_edits_do_not_change_existing_source(self):
        originals=[dict(x) for x in MARKS]
        self.save()
        self.assertEqual(MARKS,originals)

    def test_confirmation_requires_explicit_final_approval_for_simple_ui(self):
        with TemporaryDirectory() as td:
            run=Path(td)
            save(run/'verified.json',{
                'plan_seal':'XYZ','fingerprints':{},
            })
            plan={'run':str(run),'seal':'XYZ',
                  'facts':{'simple_ui':True,'responsible':''}}
            with patch.object(core,'load_plan',return_value=plan):
                with self.assertRaisesRegex(Stop,'Не подтверждены'):
                    core.confirm(run/'plan.json','',True)
                self.assertFalse((run/'confirmation.json').exists())
                result=core.confirm(run/'plan.json','',True,approved=True)
            proof=read(run/'confirmation.json')
            self.assertTrue(proof['reviewed'])
            self.assertTrue(proof['approved'])
            self.assertTrue(Path(result['confirmation']).is_file())

    def test_old_flow_still_accepts_existing_approval_contract(self):
        with TemporaryDirectory() as td:
            run=Path(td)
            save(run/'verified.json',{'plan_seal':'XYZ','fingerprints':{}})
            plan={'run':str(run),'seal':'XYZ',
                  'facts':{'responsible':'LEAD'}}
            with patch.object(core,'load_plan',return_value=plan):
                core.confirm(run/'plan.json','LEAD',True)
            self.assertFalse(read(run/'confirmation.json')['approved'])

    def test_deferred_approval_does_not_defer_missing_uzk_facts(self):
        datum={
            'project':'P','number':'1/2026',
            'kind':'TRAINING','confirmed':False,
            'authorized':False,'forms_approved':False,
            'responsible':'','no_responsible_input':True,
            'marks':[{
                'mark':'M1','profile':'P','akz':None,
                'uzk_quantity':'1','uzk_result':'',
            }],
        }
        vom=SimpleNamespace(rows=[('M1','Part',1,1,1,1,1)])
        with self.assertRaisesRegex(Stop,'результата УЗК'):
            validate_facts(datum,vom,'P','1/2026',
                           training=True,
                           defer_registry_dates=True,
                           defer_approvals=True)

    def test_original_and_simple_workflows_distinguished(self):
        old={'project':'Existing','number':'1/2026','climates':[1]}
        self.assertEqual(prepare_for_check(old),old)


class GUIContracts(unittest.TestCase):
    def test_only_climate_and_fixed_back_footer_visible(self):
        code=inspect.getsource(gui.App.fact_dialog)
        self.assertIn("text='← Назад'",code)
        self.assertIn("footer.pack(side='bottom'",code)
        self.assertIn("save_screen_draft(",code)
        self.assertNotIn("Проверить и принять введённые данные",code)
        for removed in (
            "Добавить фактическое соединение",
            "Дополнительные реквизиты утверждённых форм",
            "Приложить первичные протоколы",
            "Подтверждаю фактические результаты",
            "Заключения текущего контроля",
            "Обновить назначения из справочника",
        ):
            self.assertNotIn(removed,code)

    def test_main_generate_prepares_draft_before_validating(self):
        code=inspect.getsource(gui.App.check)
        self.assertIn('prepare_for_check(self.facts)',code)
        self.assertIn("self.run('check'",code)

    def test_simple_draft_never_gets_implicit_signed_approval(self):
        code=inspect.getsource(core.plan)
        self.assertIn('defer_approvals=bool(facts.get(',code)
        approve=inspect.getsource(core.confirm)
        self.assertIn('approved is not True',approve)
        commit=inspect.getsource(core.commit)
        self.assertIn("confirmation.get('approved') is not True",commit)
        self.assertIn('factual.validate(approved_facts',commit)

    def test_final_dialog_explains_duty_to_review_and_approve(self):
        code=inspect.getsource(gui.App.review_preview)
        self.assertIn('достоверны',code)
        self.assertIn('утверждённые формы действуют',code)


if __name__=='__main__':
    unittest.main(verbosity=2)
