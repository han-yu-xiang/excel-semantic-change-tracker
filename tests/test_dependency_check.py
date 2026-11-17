import sys
import unittest
from subprocess import CompletedProcess
from unittest.mock import Mock

from dependency_check import REQUIRED_MODULES, check_dependencies


class DependencyCheckTests(unittest.TestCase):
    def test_missing_import_is_reported_before_launch(self):
        run_command = Mock()

        ready, message = check_dependencies(
            module_finder=lambda name: None if name == "scipy" else object(),
            run_command=run_command,
        )

        self.assertFalse(ready)
        self.assertIn("scipy", message)
        run_command.assert_not_called()

    def test_inconsistent_environment_is_reported_before_launch(self):
        run_command = Mock(return_value=CompletedProcess([], 1, "", "broken dependency"))

        ready, message = check_dependencies(
            module_finder=lambda _name: object(),
            run_command=run_command,
        )

        self.assertFalse(ready)
        self.assertIn("broken dependency", message)
        run_command.assert_called_once_with(
            [sys.executable, "-m", "pip", "check"], capture_output=True, text=True, check=False
        )

    def test_complete_environment_passes_the_import_and_pip_checks(self):
        run_command = Mock(return_value=CompletedProcess([], 0, "No broken requirements found.\n", ""))

        ready, message = check_dependencies(
            module_finder=lambda _name: object(),
            run_command=run_command,
        )

        self.assertTrue(ready)
        self.assertEqual("", message)
        run_command.assert_called_once()
        self.assertEqual(4, len(REQUIRED_MODULES))


if __name__ == "__main__":
    unittest.main()

