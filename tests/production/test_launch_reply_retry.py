import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from stage8_closure import storage
from stage9_app import ownership


class LaunchResponse(unittest.TestCase):
    def test_reopens_same_response_after_transient_windows_lock(self):
        with tempfile.TemporaryDirectory() as folder:
            reply = Path(folder) / 'excel.response.json'
            reply.write_text('{"status":"PASS"}', encoding='utf8')
            read = Path.read_text
            attempted = []

            def temporary_lock(path, *args, **kwargs):
                attempted.append(path)
                if len(attempted) == 1:
                    raise PermissionError('temporary sharing violation')
                return read(path, *args, **kwargs)

            with patch.object(Path, 'read_text', temporary_lock), patch.object(ownership.time, 'sleep'):
                self.assertEqual(ownership._read_launch_response(reply), {'status': 'PASS'})
            self.assertEqual(attempted, [reply, reply])

    def test_missing_response_at_deadline_reports_timeout(self):
        with tempfile.TemporaryDirectory() as folder:
            with self.assertRaisesRegex(RuntimeError, 'OFFICE_LAUNCH_RESPONSE_TIMEOUT'):
                ownership._read_launch_response(Path(folder) / 'absent', timeout=0)

    def test_invalid_json_is_not_retried(self):
        with tempfile.TemporaryDirectory() as folder:
            reply = Path(folder) / 'excel.response.json'
            reply.write_text('{', encoding='utf8')
            with patch.object(ownership.time, 'sleep') as sleep, self.assertRaises(storage.Stop):
                ownership._read_launch_response(reply)
            sleep.assert_not_called()
