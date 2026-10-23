"""The repair command must reject substituted downloads before installing code."""

import io
import runpy
import subprocess
import sys
import tempfile
import unittest
from unittest import mock
from pathlib import Path


class PythonTrustRepairTest(unittest.TestCase):
    def test_repair_rejects_oversized_download_before_hashing(self):
        script = Path(__file__).with_name("upgrade-python-trust.py")
        download = runpy.run_path(str(script))["download"]
        limit = 8 * 1024 * 1024
        response = io.BytesIO(b"x" * (limit + 1))
        with (
            tempfile.TemporaryDirectory() as directory,
            mock.patch("urllib.request.urlopen", return_value=response),
            mock.patch.object(response, "read", wraps=response.read) as read,
            mock.patch("hashlib.sha256") as hash_wheel,
        ):
            with self.assertRaisesRegex(ValueError, "exceeds.*refusing to install"):
                download("pip", Path(directory))
            read.assert_called_once_with(limit + 1)
            hash_wheel.assert_not_called()
            self.assertEqual(list(Path(directory).iterdir()), [])

    def test_repair_rejects_tampered_download(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        environment = Path(temporary.name) / "venv"
        subprocess.run(
            [sys.executable, "-m", "venv", "--without-pip", str(environment)],
            check=True,
        )
        python = environment / (
            "Scripts/python.exe" if sys.platform == "win32" else "bin/python"
        )
        script = Path(__file__).with_name("upgrade-python-trust.py")
        result = subprocess.run(
            [
                str(python),
                "-I",
                "-c",
                "import io, runpy, sys, urllib.request; "
                "urllib.request.urlopen = lambda *args, **kwargs: io.BytesIO(b'tampered wheel'); "
                "sys.argv = [sys.argv[1]]; runpy.run_path(sys.argv[0], run_name='__main__')",
                str(script),
            ],
            capture_output=True,
            text=True,
        )
        assert result.returncode == 1
        assert "SHA-256 mismatch" in result.stderr
        audit = subprocess.run(
            [str(python), "-I", str(script), "--check"], capture_output=True, text=True
        )
        assert audit.returncode == 1
        assert "pip: missing" in audit.stdout
        assert "certifi: missing" in audit.stdout


if __name__ == "__main__":
    unittest.main()
