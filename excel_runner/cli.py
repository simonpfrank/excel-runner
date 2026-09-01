"""CLI entrypoint: run a workflow YAML file.

Exists so a workflow can be triggered from outside Python — the driving use case is an external
orchestration/automation workflow invoking this as a process/command-line step, but any external
caller that can run a command and check the exit code works the same way. Only wraps
`run_workflow()` (Spec sec 6.1); no behavior of its own beyond argument parsing plus console
logging setup. Results live at the run's fixed `working_dir` path
(`excel_runner_runs/<yaml_stem>/audit.jsonl`, PRD sec 6.3.4) — an external caller can read that
file directly.

This is the one place in the project that attaches logging handlers (AGENTS.md's logging
section) — every library module (`runner.py`, `engine.py`, `backends.py`, ...) only ever calls
`logging.getLogger(__name__)` and never configures a handler itself.
"""

import argparse
import logging
import sys

from excel_runner.core import ExcelRunnerError
from excel_runner.runner import run_workflow

logger = logging.getLogger(__name__)

_LOG_FORMAT = "%(asctime)s %(levelname)-8s %(module)s %(funcName)s:%(lineno)d %(message)s"
_LOG_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"


class _BelowWarningFilter(logging.Filter):
    """Lets only DEBUG/INFO records through a handler — WARNING and above go to stderr
    instead, via a separate handler with its own level floor."""

    def filter(self, record: logging.LogRecord) -> bool:
        return record.levelno < logging.WARNING


def configure_logging(level_name: str) -> None:
    """Attach the CLI's stdout/stderr console handlers to the `excel_runner` logger.

    DEBUG/INFO go to stdout; WARNING/ERROR/CRITICAL go to stderr — so a caller piping only one
    stream still sees a coherent picture (AGENTS.md's logging section). Clears any handlers
    already on the logger first, so repeated calls (e.g. `main()` invoked more than once in the
    same process, as tests do) don't accumulate duplicate handlers.

    Args:
        level_name: One of "DEBUG", "INFO", "WARNING", "ERROR" — the `excel_runner` logger's
            new severity threshold.
    """
    package_logger = logging.getLogger("excel_runner")
    package_logger.handlers.clear()
    formatter = logging.Formatter(_LOG_FORMAT, datefmt=_LOG_DATE_FORMAT)

    stdout_handler = logging.StreamHandler(sys.stdout)
    stdout_handler.setFormatter(formatter)
    stdout_handler.addFilter(_BelowWarningFilter())
    package_logger.addHandler(stdout_handler)

    stderr_handler = logging.StreamHandler(sys.stderr)
    stderr_handler.setFormatter(formatter)
    stderr_handler.setLevel(logging.WARNING)
    package_logger.addHandler(stderr_handler)

    package_logger.setLevel(getattr(logging, level_name))


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

    configure_logging(args.logging_level)

    try:
        result = run_workflow(args.workflow, env_overrides or None, working_dir=args.working_dir)
    except ExcelRunnerError as exc:
        logger.error(exc.detail.message)
        return 1

    return 0 if result.status == "success" else 1


if __name__ == "__main__":
    sys.exit(main())
