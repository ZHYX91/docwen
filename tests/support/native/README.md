# Linux ORT opt-out probe

Run `ORT_DISABLE_TELEMETRY=1 .venv/bin/python tests/support/native/verify_ort_optout.py`.
Requires Linux, a C compiler, repository OCR models, and installed project dependencies.
The helper compiles its fixture into a temporary directory; no telemetry-enabled
control is ever executed. It verifies interception with synthetic native calls
that are blocked before any network activity, then imports real ORT and runs the
repository's detection ONNX model on a blank image through a CPU session.

The preload fixture records and blocks the listed native libc network APIs and
forces `fopen` / `fopen64` of the exact DRM vendor path to fail with `ENOENT`.
It reports whether ORT actually opens that path. Hosts without a matching DRM
entry may have zero vendor-file reads, so check the reported count when
evaluating device-discovery behavior.
This is **not** a packet capture or proof against direct syscalls, static code,
private libc symbols, or other uninstrumented APIs. Keep those limitations in any
reported result. It supplements the ordinary unit tests, not the acceptance suite.

The source diagnostic uses the environment opt-out alone, with no call to
`disable_telemetry_events()`, so that API cannot mask an environment regression.

To additionally run the packaged Machine corpus:

```sh
ORT_DISABLE_TELEMETRY=1 .venv/bin/python tests/support/native/verify_ort_optout.py \
  --binary-dir /path/to/DocWenCLI_package
```

Use `--binary-name` if the executable is not named `DocWenCLI`. The packaged
process alone receives the preload library; its load event and absence of
instrumented network calls are checked independently of the source probe.
The probe reports the exact source and frozen stderr bytes as base64 separately
from its network observations. The packaged corpus checks framed stdout and
Bundle semantics, and rejects misplaced protocol frames or more than 256 KiB
on stderr. Bounded runtime diagnostics on stderr can coexist with a successful
conversion.
