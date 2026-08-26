"""Qt bridge layer for the DocWen GUI.

Bridges external (non-Qt) event sources into Qt's main-thread event
loop.  All components here are thread-safe: events arriving from
background threads are queued and dispatched on the main thread via
Qt signals with ``Qt.QueuedConnection``.
"""
