"""CLI entrypoint: run a workflow YAML file.

Exists so a workflow can be triggered from outside Python — the driving use case is an external
orchestration/automation workflow invoking this as a process/command-line step, but any external
caller that can run a command and check the exit code works the same way. Only wraps
`run_workflow()` (Spec sec 6.1); no behavior of its own beyond argument parsing. Results live at
the run's fixed `working_dir` path (`excel_runner_runs/<yaml_stem>/audit.jsonl`, PRD sec
6.3.4) — an external caller can read that file directly, so nothing is printed to stdout beyond
whatever the console logger (Spec sec 6.2.1) is configured to show.
"""

import argparse
import logging
import sys

from excel_runner.core import ExcelRunnerError
from excel_runner.runner import run_workflow

logger = logging.getLogger(__name__)


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
    """Run a workflow YAML file.

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
    parser.add_argument(
        "--working-dir",
        default=None,
        help=(
            "Base directory for this run's working_dir "
            "(excel_runner_runs/<yaml_stem>/ is always appended). Defaults to cwd."
        ),
    )
    parser.add_argument("--logging-level", help="DEBUG,INFO,WARNING,ERROR", default="INFO")
    args = parser.parse_args(argv)
    env_overrides = dict(args.env)

    # Only sets the severity threshold — no handler/formatter configuration here at all, that's
    # entirely the responsibility of whatever wraps this (Spec sec 6.2.1).
    logging.getLogger("excel_runner").setLevel(getattr(logging, args.logging_level))

    try:
        result = run_workflow(args.workflow, env_overrides or None, working_dir=args.working_dir)
    except ExcelRunnerError as exc:
        logger.error(exc.detail.message)
        return 1

    return 0 if result.status == "success" else 1


if __name__ == "__main__":
    sys.exit(main())
