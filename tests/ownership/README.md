# Ownership harness

Synthetic-boundary regression tests for 4D native ownership (strings, pictures,
variables). The tests compile the real plugin sources (`message-file-drop/4DPlugin.cpp`, `4D Plugin API/Classes/*.cpp`) together with the
unchanged bundled SDK (`4DPluginAPI.c`) under ASan/UBSan. They replace `Call4D` with a stub engine
(`engine_harness.mm`) that tracks every string, picture and handle the engine
allocates. A test fails on a leak, a double or foreign dispose, or a sanitizer report.

```sh
tests/ownership/run.sh . fixed                 # every test must pass

T=$(mktemp -d); git archive 4c0d991 | tar -x -C "$T"
tests/ownership/run.sh "$T" baseline           # named defects must fail, controls pass
```

`listenerLoopExecuteMethod` is reached only on Windows (inside `#if VERSIONWIN`).
The harness calls it directly on macOS. Windows and real 4D runtime behaviour are
not exercised.
