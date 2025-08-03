# Support running as a package or as a script by trying both relative and
# absolute imports for the ``utils`` module.
try:
    from .utils import PREVIOUS_VERSION, CURRENT_VERSION  # type: ignore
except ImportError:  # pragma: no cover - fallback when executed as script
    from utils import PREVIOUS_VERSION, CURRENT_VERSION

__version__ = CURRENT_VERSION

__all__ = ["__version__", "PREVIOUS_VERSION", "CURRENT_VERSION"]
