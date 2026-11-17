"""Check that the application imports and its Python environment is consistent."""

from __future__ import annotations

from collections.abc import Callable
import importlib.util
import subprocess
import sys


REQUIRED_MODULES = ("openpyxl", "numpy", "scipy", "sentence_transformers")


def check_dependencies(
    module_finder: Callable[[str], object | None] | None = None,
    run_command: Callable[..., subprocess.CompletedProcess] | None = None,
) -> tuple[bool, str]:
    """Return whether imports and installed package requirements are healthy."""
    module_finder = module_finder or importlib.util.find_spec
    run_command = run_command or subprocess.run
    missing = [name for name in REQUIRED_MODULES if module_finder(name) is None]
    if missing:
        return False, "缺少 Python 模块：" + ", ".join(missing)

    try:
        result = run_command(
            [sys.executable, "-m", "pip", "check"],
            capture_output=True,
            text=True,
            check=False,
        )
    except OSError as exc:
        return False, f"无法运行 pip 检查：{exc}"
    if result.returncode:
        details = (result.stderr or result.stdout).strip()
        return False, details or "pip check 检测到依赖不完整。"
    return True, ""


def main() -> int:
    ready, message = check_dependencies()
    if not ready:
        print(message, file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

