"""Glycan sequence enumeration toolkit."""

from importlib.metadata import version, PackageNotFoundError

try:  # pragma: no cover - fallback when package metadata unavailable
    __version__ = version("glycoenum")
except PackageNotFoundError:  # pragma: no cover
    __version__ = "0.0.0"

__all__ = [
    "__version__",
]
