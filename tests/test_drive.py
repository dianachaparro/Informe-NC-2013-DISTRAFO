from pathlib import Path
import tempfile
import unittest
from unittest.mock import MagicMock

from subir_drive import folder_id, upload_reports, validate_inputs, FOLDER_MIME


class DriveTests(unittest.TestCase):
    def test_folder_link_and_invalid_host(self):
        self.assertEqual(folder_id('https://drive.google.com/drive/u/1/folders/abc-123?usp=sharing'), 'abc-123')
        with self.assertRaises(ValueError):
            folder_id('https://example.com/folders/abc')
        with self.assertRaises(ValueError):
            folder_id("abc' or trashed=false")

    def test_read_only_account_does_not_create_files(self):
        service = MagicMock()
        service.files.return_value.get.return_value.execute.return_value = {
            'mimeType': FOLDER_MIME, 'capabilities': {'canAddChildren': False}}
        with self.assertRaisesRegex(ValueError, 'permiso'):
            upload_reports(service, 'parent', [Path('informe.xlsx')], MagicMock())
        service.files.return_value.create.assert_not_called()

    def test_existing_folder_reused_without_replacing_report(self):
        service = MagicMock()
        files = service.files.return_value
        files.get.return_value.execute.return_value = {
            'mimeType': FOLDER_MIME, 'capabilities': {'canAddChildren': True}}
        files.list.return_value.execute.return_value = {'files': [{'id': 'child'}]}
        files.create.return_value.execute.return_value = {'id': 'new', 'webViewLink': 'https://example.com/new'}
        media = MagicMock()
        results = upload_reports(service, 'parent', [Path('informe.xlsx')], media)
        self.assertEqual(results[0]['id'], 'new')
        self.assertEqual(files.create.call_args.kwargs['body']['parents'], ['child'])
        files.update.assert_not_called()

    def test_missing_report_fails_before_login(self):
        with tempfile.TemporaryDirectory() as temp:
            client = Path(temp) / 'credentials.json'
            client.write_text('{"installed": {"client_id": "test"}}')
            with self.assertRaisesRegex(ValueError, 'Excel válido'):
                validate_inputs(client, [Path(temp) / 'no-existe.xlsx'])

    def test_ambiguous_destination_does_not_create_files(self):
        service = MagicMock()
        files = service.files.return_value
        files.get.return_value.execute.return_value = {
            'mimeType': FOLDER_MIME, 'capabilities': {'canAddChildren': True}}
        files.list.return_value.execute.return_value = {'files': [{'id': 'a'}, {'id': 'b'}]}
        with self.assertRaisesRegex(ValueError, 'varias'):
            upload_reports(service, 'parent', [Path('informe.xlsx')], MagicMock())
        files.create.assert_not_called()


if __name__ == '__main__':
    unittest.main()
