"""CLI entrypoint: run a workflow YAML file and print its result as JSON.

Exists so a workflow can be triggered from outside Python — the driving use case is a UiPath
xaml workflow invoking this as a process/command-line step, but any external caller that can
run a command and read stdout/exit code works the same way. Only wraps `run_workflow()`
(Spec sec 6.1); no behavior of its own beyond argument parsing and result formatting.
"""

import argparse
import json
import sys
from dataclasses import asdict
from typing import Any

from excel_runner.core import ExcelRunnerError
from excel_runner.runner import run_workflow


def _parse_env_override(raw: str) -> tuple[str, str]:
    """Parse one `--env KEY=VALUE` argument.

    Args:
        raw: The raw "KEY=VALUE" string.

    Returns:
        The (key, value) pair.

    Raises:
        argparse.ArgumentTypeError: If `raw` doesn't contain "=".
    """
    if "=" not in raw:
        raise argparse.ArgumentTypeError(
            f"--env value {raw!r} must be in KEY=VALUE form"
        )
    key, _, value = raw.partition("=")
    return key, value


def main(argv: list[str] | None = None) -> int:
    """Run a workflow YAML file and print its result as JSON to stdout.

    Args:
        argv: Command-line arguments, excluding the program name. Defaults to `sys.argv[1:]`.

    Returns:
        0 if the workflow ran and every step succeeded/was skipped, 1 if any step errored or
        the workflow itself failed to load/validate/execute.
    """
    parser = argparse.ArgumentParser(
        prog="excel-runner", description="Run a declarative Excel workflow YAML file."
    )
    parser.add_argument("workflow", help="Path to the workflow YAML file.")
    parser.add_argument(
        "--env",
        action="append",
        default=[],
        type=_parse_env_override,
        metavar="KEY=VALUE",
        help=(
            "Override/add an env: value. Repeat the flag once per value, e.g. "
            "--env a=1 --env b=2 (not comma- or semicolon-separated)."
        ),
    )
    args = parser.parse_args(argv)
    env_overrides = dict(args.env)

    try:
        result = run_workflow(args.workflow, env_overrides or None)
    except ExcelRunnerError as exc:
        print(json.dumps({"status": "error", "error": asdict(exc.detail)}, default=str))
        return 1

    output: dict[str, Any] = asdict(result)
    print(json.dumps(output, default=str))
    return 0 if result.status == "success" else 1


if __name__ == "__main__":
    sys.exit(main())
