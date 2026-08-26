"""Command-line entry point for the bundled Claude skill.

Usage::

    python -m pptx2.skill install [TARGET]   # copy skill into target dir
    python -m pptx2.skill path               # print on-disk skill source
    python -m pptx2.skill                    # same as `path`
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pptx2.skill import DEFAULT_INSTALL_DIR, install_skill, skill_root


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="python -m pptx2.skill",
        description=(
            "Manage the bundled python-pptx2 Claude skill.  "
            "Ships SKILL.md and reference docs to your Claude skills directory."
        ),
    )
    sub = parser.add_subparsers(dest="cmd")

    install = sub.add_parser(
        "install", help="copy the bundled skill into a Claude skills directory"
    )
    install.add_argument(
        "target",
        nargs="?",
        type=Path,
        default=None,
        help=f"destination directory (default: {DEFAULT_INSTALL_DIR})",
    )
    install.add_argument(
        "--no-overwrite",
        action="store_true",
        help="error out if the destination already exists",
    )

    sub.add_parser("path", help="print the package-internal skill directory")

    args = parser.parse_args(argv)

    if args.cmd == "install":
        try:
            dest = install_skill(args.target, overwrite=not args.no_overwrite)
        except FileExistsError as exc:
            print(str(exc), file=sys.stderr)
            return 1
        print(f"installed python-pptx2 skill -> {dest}")
        return 0

    # default + `path`
    print(skill_root())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
