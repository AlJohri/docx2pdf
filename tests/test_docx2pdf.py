import subprocess
import sys
import unittest
from os.path import join
from tempfile import TemporaryDirectory

from docx2pdf import __version__, convert


def test_version():
    assert __version__ == '0.1.8'


@unittest.skipIf(sys.platform != 'darwin', 'MacOS-only tests')
class TestMacOSFailureModes(unittest.TestCase):
    """MacOS failure modes."""

    def test_handle_microsoft_word_not_found_on_macos(self):
        # Given that 'Microsoft Word' cannot be launched...
        proc = subprocess.Popen(
            ['open', '-Ra', 'Microsoft Word'], stderr=subprocess.PIPE
        )
        proc.wait()
        can_launch_msword = (proc.returncode == 0)
        if can_launch_msword:
            self.skipTest('Microsoft Word is present, so can\'t run '
                          'this test on this system.')

        # ...when convert() is called, then a RuntimeError is raised.
        with TemporaryDirectory() as td:
            input_path = join(td, 'some_file.docx')
            with self.assertRaisesRegex(
                RuntimeError, 'Application can\'t be found.'
            ):
                convert(input_path)
