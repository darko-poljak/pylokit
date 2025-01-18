import unittest
import os
import shutil
import multiprocessing as mp

from pylokit import (
    Office,
    LoKitInitializeError,
    LoKitImportError,
    LoKitExportError
)

TEST_DIR = os.path.dirname(__file__)


class LokitTest(unittest.TestCase):
    def setUp(self):
        self.lo_path = os.getenv('LO_PATH', '/usr/lib/libreoffice/program/')
        self.test_doc = self._to_abspath("tests/foo.doc")
        self.test_out_rtf = self._to_abspath("tests/out.rtf")
        self.test_out_docx = self._to_abspath("tests/out.docx")
        self.test_out_pdf = self._to_abspath("tests/out.pdf")
        self.test_lockfile = self._to_abspath(".~lock.foo.doc#")
        self.profile_dir = self._to_abspath("tests/lokit-profile")
        self.profile_url = "file://" + self.profile_dir

    def _to_abspath(self, path):
        return os.path.abspath(os.path.join(TEST_DIR, path))

    def _subprocess_func(self, func):
        def _func():
            func()
            os._exit(0)

        return _func

    def _run_in_subprocess(self, func):
        p = mp.Process(target=self._subprocess_func(func))
        p.start()
        p.join()
        exitcode = p.exitcode
        p.close()
        self.assertEqual(exitcode, 0)

    def test_init_no_lo_path(self):
        self.assertRaises(LoKitInitializeError, Office, "wronglopath")

    def test_init(self):
        def func():
            with Office(self.lo_path) as lo:
                self.assertIsNotNone(lo)

        self._run_in_subprocess(func)

    def test_init_with_profile_url(self):
        def func():
            with Office(self.lo_path, self.profile_url) as lo:
                self.assertIsNotNone(lo)
            shutil.rmtree(self.profile_dir)

        self._run_in_subprocess(func)

    def test_no_input_file(self):
        def func():
            with Office(self.lo_path) as lo:
                self.assertRaises(LoKitImportError, lo.documentLoad, "foo")

        self._run_in_subprocess(func)

    def test_no_output_file(self):
        def func():
            with Office(self.lo_path) as lo:
                with lo.documentLoad(self.test_doc) as doc:
                    self.assertRaises(LoKitExportError, doc.saveAs, "")

        self._run_in_subprocess(func)

    def test_wrong_filter(self):
        def func():
            with Office(self.lo_path) as lo:
                with lo.documentLoad(self.test_doc) as doc:
                    self.assertRaises(LoKitExportError, doc.saveAs,
                                      self.test_out_rtf, fmt="foobar")

        self._run_in_subprocess(func)

    def test_wrong_options(self):

        def func():
            # 4.4 used to fail, need to investigate what's going on
            # 2025-01-18: it seems that unknown options are now ignored.
            with Office(self.lo_path) as lo:
                with lo.documentLoad(self.test_doc) as doc:
                    doc.saveAs(self.test_out_rtf, fmt="docx", options="foobar")
                    os.unlink(self.test_out_rtf)

        self._run_in_subprocess(func)

    def test_filter_and_options(self):
        def func():
            with Office(self.lo_path) as lo:
                with lo.documentLoad(self.test_doc) as doc:
                    doc.saveAs(self.test_out_docx, fmt="docx",
                               options="SkipImages")
                    os.unlink(self.test_out_docx)

        self._run_in_subprocess(func)

    def test_multiple_calls(self):
        def func():
            with Office(self.lo_path) as lo:
                with lo.documentLoad(self.test_doc) as doc:
                    doc.saveAs(self.test_out_docx, fmt="docx",
                               options="SkipImages")
                    doc.saveAs(self.test_out_pdf)
                    os.unlink(self.test_out_docx)
                    os.unlink(self.test_out_pdf)
                with lo.documentLoad(self.test_doc) as doc:
                    doc.saveAs(self.test_out_pdf)
                    os.unlink(self.test_out_pdf)

        self._run_in_subprocess(func)

    def test_convert_with_profile_url(self):
        def func():
            with Office(self.lo_path, self.profile_url) as lo:
                with lo.documentLoad(self.test_doc) as doc:
                    doc.saveAs(self.test_out_docx, fmt="docx",
                               options="SkipImages")
                    os.unlink(self.test_out_docx)
            shutil.rmtree(self.profile_dir)

        self._run_in_subprocess(func)

    def tearDown(self):
        assert os.path.exists(self.test_lockfile) is False


if __name__ == '__main__':
    unittest.main()
