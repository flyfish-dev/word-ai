from __future__ import annotations

import os
import sys

from word_ai_mcp import __version__
from word_ai_mcp.resources import executable_path


HELP = f"""Word AI {__version__}

Single-file launcher for the Word AI MCP server, Office.js bridge helpers, and
agent skill installation.

Usage:
  word-ai mcp [--root DIR] [--allow-root DIR] [--read-only]
  word-ai http [--root DIR] [--host 127.0.0.1] [--port 8765]
  word-ai doctor [quickstart options]
  word-ai codex-config [quickstart options]
  word-ai install-skill [--agents auto|all|codex,...] [--project] [--dry-run]
  word-ai quickstart <doctor|codex-config|install-skills> [...]
  word-ai --version

If no subcommand is provided, Word AI starts the stdio MCP server.
"""


def _cwd_root_args(argv: list[str]) -> list[str]:
    if "--root" in argv:
        return argv
    return ["--root", os.getcwd(), *argv]


def run_mcp(argv: list[str]) -> int:
    from word_ai_mcp.server import main

    return main(argv)


def run_http(argv: list[str]) -> int:
    from word_ai_mcp.server_http import main

    return main(argv)


def run_quickstart(argv: list[str]) -> int:
    from word_ai_mcp.quickstart import main

    return main(argv)


def run_codex_config(argv: list[str]) -> int:
    return run_quickstart(
        [
            *_cwd_root_args([]),
            "codex-config",
            "--standalone-command",
            str(executable_path()),
            *argv,
        ]
    )


def run_install_skill(argv: list[str]) -> int:
    return run_quickstart([*_cwd_root_args([]), "install-skills", *argv])


def run_doctor(argv: list[str]) -> int:
    return run_quickstart([*_cwd_root_args([]), "doctor", *argv])


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if not args:
        return run_mcp([])

    first = args[0].lower()
    rest = args[1:]
    if first in {"-h", "--help", "help"}:
        print(HELP)
        return 0
    if first in {"-v", "--version", "version"}:
        print(__version__)
        return 0
    if first in {"mcp", "server", "stdio"}:
        return run_mcp(rest)
    if first in {"http", "bridge", "server-http"}:
        return run_http(rest)
    if first in {"doctor"}:
        return run_doctor(rest)
    if first in {"codex-config", "config"}:
        return run_codex_config(rest)
    if first in {"install-skill", "install-skills", "skill"}:
        return run_install_skill(rest)
    if first == "quickstart":
        return run_quickstart(rest)

    # Preserve MCP host ergonomics: unknown first token is treated as an option
    # for the stdio server, so `word-ai --root ...` works.
    if first.startswith("-"):
        return run_mcp(args)

    print(f"Unknown command: {args[0]}", file=sys.stderr)
    print(HELP, file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
