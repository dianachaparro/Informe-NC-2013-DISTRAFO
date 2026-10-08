from pathlib import Path
import tempfile
import unittest
from openpyxl import Workbook
from PIL import Image
from proyectos import new_project, import_checklist, save_project, load_project, report_case, validate_review


class ProjectTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.list = self.root / 'lista.xlsx'
        w = Workbook()
        w.active.append(['No.', 'ACTIVIDADES', 'REFERENCIA'])
        w.active.append([1, 'Verificar diseño', 'RETIE 25.2(a)'])
        w.save(self.list)

    def test_new_project_does_not_inherit_findings_or_drive(self):
        a, b = new_project(), new_project()
        import_checklist(a, 'Red subterránea', self.list)
        a['drive_folder'] = 'carpeta-a'
        self.assertEqual(b['requirements'], [])
        self.assertEqual(b['drive_folder'], '')

    def test_insufficient_evidence_does_not_become_nc(self):
        p = new_project()
        p.update(project='Obra', inspection='A-1', address='Dirección', template='formato.xlsx', scopes=['Red subterránea'])
        import_checklist(p, 'Red subterránea', self.list)
        p['requirements'][0]['result'] = 'Evidencia insuficiente'
        self.assertEqual(report_case(p)['findings'], [])
        p['requirements'][0].update(result='No cumple', observation='No se evidencia diseño.')
        self.assertEqual(len(report_case(p)['findings']), 1)

    def test_copy_evidence_and_reopen_without_original_photo(self):
        p = new_project()
        import_checklist(p, 'Red subterránea', self.list)
        image = self.root / 'original.jpg'
        Image.new('RGB', (40, 30)).save(image)
        p['requirements'][0]['photos'] = [str(image)]
        destination = self.root / 'proyecto.json'
        save_project(p, destination)
        image.unlink()
        reopened = load_project(destination)
        self.assertTrue(Path(reopened['requirements'][0]['photos'][0]).is_file())
        save_project(reopened, destination)
        self.assertTrue(destination.with_suffix('.json.bak').is_file())

    def test_cannot_replace_oauth_credentials_with_project(self):
        credentials = self.root / 'credentials.json'
        content = '{"installed": {"client_id": "example"}}'
        credentials.write_text(content)
        with self.assertRaises(ValueError):
            save_project(new_project(), credentials)
        self.assertEqual(credentials.read_text(), content)

    def test_no_cumple_requires_the_requested_wording(self):
        p = new_project()
        import_checklist(p, 'Red subterránea', self.list)
        r = p['requirements'][0]
        r.update(result='No cumple', observation='Se ve incorrecto')
        with self.assertRaises(ValueError):
            validate_review(r)

    def test_duplicate_list_does_not_add_duplicate_rows(self):
        p = new_project()
        import_checklist(p, 'Red subterránea', self.list)
        with self.assertRaises(ValueError):
            import_checklist(p, 'Red subterránea', self.list)
        self.assertEqual(len(p['requirements']), 1)


if __name__ == '__main__':
    unittest.main()
