"""Module entry point for `python -m ateco_extractor`."""

from .cli import main


if __name__ == "__main__":  # pragma: no cover - thin wrapper
    raise SystemExit(main())
