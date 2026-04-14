import logging
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from app.run_logging import RUN_LOG_PATH_ENV_VAR, setup_run_file_logging
from app.runtime_paths import ensure_workspace_dirs, resolve_runtime_paths


class RunLoggingTests(unittest.TestCase):
    def tearDown(self) -> None:
        self._close_logger_handlers()

    @staticmethod
    def _close_logger_handlers() -> None:
        logger = logging.getLogger("capacity_optimizer")
        for handler in list(logger.handlers):
            logger.removeHandler(handler)
            try:
                handler.close()
            except Exception:
                pass

    def test_setup_run_file_logging_creates_workspace_log_file(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            with patch.dict("os.environ", {"CAPACITY_OPTIMIZER_WORKSPACE": tmpdir}, clear=False):
                paths = ensure_workspace_dirs(resolve_runtime_paths())
                context = setup_run_file_logging(paths, run_label="unit_test_log")
                logger = logging.getLogger("capacity_optimizer")
                logger.debug("debug-line")
                self._close_logger_handlers()

            self.assertTrue(context.log_path.exists())
            self.assertTrue(str(context.log_path).startswith(str(Path(tmpdir))))
            content = context.log_path.read_text(encoding="utf-8")
            self.assertIn("Logger initialized.", content)
            self.assertIn("debug-line", content)

    def test_setup_run_file_logging_respects_env_override_path(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            forced_path = Path(tmpdir) / "custom_logs" / "forced.log"
            with patch.dict(
                "os.environ",
                {
                    "CAPACITY_OPTIMIZER_WORKSPACE": tmpdir,
                    RUN_LOG_PATH_ENV_VAR: str(forced_path),
                },
                clear=False,
            ):
                paths = ensure_workspace_dirs(resolve_runtime_paths())
                context = setup_run_file_logging(paths, run_label="ignored")
                logger = logging.getLogger("capacity_optimizer")
                logger.info("forced-log-line")
                self._close_logger_handlers()

            self.assertEqual(context.log_path, forced_path.resolve())
            self.assertTrue(forced_path.exists())
            self.assertIn("forced-log-line", forced_path.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
