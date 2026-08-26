"""DocWen CLI command implementations.

The parser and dispatcher import concrete command modules directly.  Keeping
this package initializer empty prevents deleted commands from surviving as an
accidental public API.
"""
