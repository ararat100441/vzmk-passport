import os,sys,unittest
from unittest.mock import Mock,patch,call
from production_app import external_open


class ExternalViewer(unittest.TestCase):
    def fake_kernel(self):
        kernel=Mock()
        def get(size,buffer):buffer.value=r'C:\bundle\_internal';return len(buffer.value)
        kernel.GetDllDirectoryW.side_effect=get
        kernel.SetDllDirectoryW.return_value=True
        return kernel

    def test_external_viewer_does_not_inherit_private_loader_and_path(self):
        import ctypes
        kernel=self.fake_kernel()
        original=r'C:\Windows;C:\bundle\_internal;C:\bundle\_internal\pywin32_system32'
        with patch.object(sys,'frozen',True,create=True),\
             patch.object(sys,'_MEIPASS',r'C:\bundle\_internal',create=True),\
             patch.object(ctypes,'WinDLL',return_value=kernel),\
             patch.dict(os.environ,{'PATH':original}):
            with patch.object(os,'startfile',side_effect=lambda p:self.assertEqual(os.environ['PATH'],r'C:\Windows')) as start:
                external_open.open_external('draft.pdf')
                start.assert_called_once_with('draft.pdf')
            self.assertEqual(os.environ['PATH'],original)
        self.assertEqual(kernel.SetDllDirectoryW.call_args_list,[call(None),call(r'C:\bundle\_internal')])

    def test_loader_restored_even_when_viewer_does_not_start(self):
        import ctypes
        kernel=self.fake_kernel()
        with patch.object(sys,'frozen',True,create=True),\
             patch.object(sys,'_MEIPASS',r'C:\bundle\_internal',create=True),\
             patch.object(ctypes,'WinDLL',return_value=kernel),\
             patch.object(os,'startfile',side_effect=OSError('No viewer')):
            before=os.environ.get('PATH')
            with self.assertRaisesRegex(OSError,'No viewer'):external_open.open_external('draft.pdf')
            self.assertEqual(os.environ.get('PATH'),before)
        self.assertEqual(kernel.SetDllDirectoryW.call_args_list,[call(None),call(r'C:\bundle\_internal')])

    def test_source_launch_does_not_change_native_loader(self):
        with patch.object(sys,'frozen',False,create=True),patch.object(os,'startfile') as start:
            external_open.open_external('draft.pdf')
        start.assert_called_once_with('draft.pdf')


if __name__=='__main__':unittest.main()
