"""Entrypoint for `python -m excel_runner`."""

import sys

from excel_runner.cli import main

if __name__ == "__main__":
    sys.exit(main())
