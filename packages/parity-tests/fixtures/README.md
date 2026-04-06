# Lane 1 fixture bootstrap

This directory is the first parity corpus for the `OfficeCLI` -> `officekit` migration.
It intentionally focuses on **source visibility + reproducibility** instead of premature target
implementation.

## Contents
- `manifest.json` — machine-readable inventory of copied artifacts, source-only references, and canonical parity scenarios.
- `source/readme/OfficeCLI.README.md` — captured README quick starts and capability claims.
- `source/ci/officecli-build.yml` — current release + smoke-test workflow.
- `source/examples/**` — copied example READMEs, generator scripts, and a few lightweight sample PPTX files.

## Why some artifacts are reference-only
Large binaries such as the 3D PowerPoint output and the full PPT template corpus are cataloged in
`manifest.json` but not duplicated yet. That keeps this bootstrap commit reviewable while still giving
lane 3 / final parity verification an explicit list of heavy assets to pull when they start executing
PPT-specific long-tail tests.

## Canonical parity scenarios captured here
1. README quick-start PowerPoint flow (`create` -> `add` -> `view` -> `get`)
2. README live-preview/watch flow
3. CI smoke test (`create test_smoke.docx` -> `add paragraph` -> `get paragraph`)
4. Word formulas example script
5. Excel charts demo script
6. PowerPoint animations example script

## Next consumers
- Lane 2: use the scenario list to prioritize early CLI/core vertical slices.
- Lane 3: map Word/Excel/PPT fixtures to concrete package integration tests.
- Lane 4: use README + CI captures to preserve installation, skills, preview, and docs promises.
- Final verification: build differential scripts around the scenarios listed in `manifest.json`.
