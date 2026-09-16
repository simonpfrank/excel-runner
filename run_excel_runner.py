"""Compatibility launcher for environments that cannot use `python -m excel_runner`."""

import os
import sys
from pathlib import Path


def _runner_home() -> Path:
    arguments = sys.argv
    runner_home = os.environ.get("EXCEL_RUNNER_HOME")
    if "--runner-home" in arguments:
        index = arguments.index("--runner-home")
        if index + 1 == len(arguments):
            raise SystemExit("--runner-home requires the extracted deployment directory.")
        runner_home = arguments[index + 1]
        del arguments[index : index + 2]

    directory = Path(runner_home).resolve() if runner_home else Path(__file__).resolve().parent
    if not (directory / "excel_runner" / "__init__.py").is_file():
        raise SystemExit(
            f'Excel Runner package not found in "{directory}". '
            "Set --runner-home or EXCEL_RUNNER_HOME to the extracted deployment directory."
        )
    return directory


_RUNNER_HOME = str(_runner_home())
if _RUNNER_HOME not in sys.path:
    sys.path.insert(0, _RUNNER_HOME)

from excel_runner.cli import main  # noqa: E402 - package location is configured above.


if __name__ == "__main__":
    raise SystemExit(main())