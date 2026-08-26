from __future__ import annotations

import warnings

import _pytest.unraisableexception as pytest_unraisable
import pytest


def pytest_configure(config: pytest.Config) -> None:
    # Pillow may emit unraisable cleanup warnings for multipage TIFF writers during
    # interpreter/session teardown on Python 3.14. This is noisy in pytest and not
    # actionable at the product level once explicit image close paths are already in place.
    warnings.filterwarnings(
        "ignore",
        message=r"Exception ignored while finalizing file <PIL\.TiffImagePlugin\.AppendingTiffWriter object .*: None",
        category=pytest.PytestUnraisableExceptionWarning,
    )

    def _filtered_collect_unraisable(active_config: pytest.Config) -> None:
        pop_unraisable = active_config.stash[pytest_unraisable.unraisable_exceptions].pop
        errors: list[pytest.PytestUnraisableExceptionWarning | RuntimeError] = []
        meta = None
        hook_error = None
        try:
            while True:
                try:
                    meta = pop_unraisable()
                except IndexError:
                    break

                if isinstance(meta, BaseException):
                    hook_error = RuntimeError("Failed to process unraisable exception")
                    hook_error.__cause__ = meta
                    errors.append(hook_error)
                    continue

                if "PIL.TiffImagePlugin.AppendingTiffWriter object" in meta.msg:
                    continue

                try:
                    warnings.warn(pytest.PytestUnraisableExceptionWarning(meta.msg), stacklevel=2)
                except pytest.PytestUnraisableExceptionWarning as exc:
                    if meta.exc_value is not None:
                        exc.args = (meta.cause_msg,)
                        exc.__cause__ = meta.exc_value
                    errors.append(exc)

            if len(errors) == 1:
                raise errors[0]
            if errors:
                raise ExceptionGroup("multiple unraisable exception warnings", errors)
        finally:
            del errors, meta, hook_error

    pytest_unraisable.collect_unraisable = _filtered_collect_unraisable


__all__ = ["pytest_configure"]
