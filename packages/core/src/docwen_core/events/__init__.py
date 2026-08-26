"""Events — task events and application events."""

from docwen_core.events.app_events import (
    APP_STARTED,
    BATCH_CANCELLED,
    BATCH_COMPLETED,
    BATCH_STARTED,
    CONFIG_CHANGED,
    FILES_ADDED,
    FILES_DROPPED,
    IPC_FILE_RECEIVED,
    TASK_ENQUEUED,
    TASK_FINISHED,
    WINDOW_ACTIVATED,
    AppEvent,
)
from docwen_core.events.task_events import (
    ARTIFACT_READY,
    DIAGNOSTIC,
    TASK_CANCELLED,
    TASK_COMPLETED,
    TASK_FAILED,
    TASK_PROGRESS,
    TASK_STARTED,
    make_artifact_ready,
    make_diagnostic,
    make_task_cancelled,
    make_task_completed,
    make_task_failed,
    make_task_progress,
    make_task_started,
)

__all__ = [
    # App event constants
    "APP_STARTED",
    # Task event constants
    "ARTIFACT_READY",
    "BATCH_CANCELLED",
    "BATCH_COMPLETED",
    "BATCH_STARTED",
    "CONFIG_CHANGED",
    "DIAGNOSTIC",
    "FILES_ADDED",
    "FILES_DROPPED",
    "IPC_FILE_RECEIVED",
    "TASK_CANCELLED",
    "TASK_COMPLETED",
    "TASK_ENQUEUED",
    "TASK_FAILED",
    "TASK_FINISHED",
    "TASK_PROGRESS",
    "TASK_STARTED",
    "WINDOW_ACTIVATED",
    "AppEvent",
    # Task event factories
    "make_artifact_ready",
    "make_diagnostic",
    "make_task_cancelled",
    "make_task_completed",
    "make_task_failed",
    "make_task_progress",
    "make_task_started",
]
