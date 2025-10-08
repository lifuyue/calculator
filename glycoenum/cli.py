"""Legacy CLI shim that forwards to the GUI application."""

from __future__ import annotations

from typing import Iterable

from glycoenum.gui import main as launch_gui


def main(_argv: Iterable[str] | None = None) -> None:
    """Invoke the GUI application from historical CLI entry points."""
    launch_gui()


if __name__ == "__main__":  # pragma: no cover
    main()
