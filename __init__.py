# Support running as a package or as a script by trying both relative and
# absolute imports for the ``utils`` module.
try:
    from .utils.updater import PREVIOUS_VERSION, CURRENT_VERSION, __version__  # type: ignore
except ImportError:  # pragma: no cover - fallback when executed as script
    from utils.updater import PREVIOUS_VERSION, CURRENT_VERSION, __version__

__all__ = ["__version__", "PREVIOUS_VERSION", "CURRENT_VERSION"]
