# Reference Materials

This folder contains the original v4 build script that produced the workbook
used as the Sprint 2 golden file. It is preserved here for traceability —
not for execution as part of the application.

## Contents

- `build_v4.py` — the original script that generated the reference workbook
  at `tests/golden_files/v4_reference.xlsx`. Reconstructed from the v4
  specification PDF at the start of Sprint 2.

## Why this is in the repo

The Streamlit application's faithful refactor (Sprint 2) targets exact
output parity with the workbook this script produces. Keeping the script
alongside the golden file documents how the baseline was created.

## Status

This script is reference material only. Do not import from it. The
production application code lives in the repo root and (later) in the
`builder/` package.
