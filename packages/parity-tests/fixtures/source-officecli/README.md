# Harvested OfficeCLI parity fixtures

This directory contains the lane-1 harvested parity baseline for the officekit migration.

## Included here
- Source example READMEs and executable example scripts copied from `../OfficeCLI/examples/**`
- Small representative binary outputs for Word, Excel, and PowerPoint differential checks
- The upstream CI workflow used as the source smoke-test reference

## Intentionally referenced but not vendored yet
- Large assets such as `examples/ppt/models/sun.glb` and the generated `3d-sun.pptx`
- The full template style tree under `examples/ppt/templates/styles/**`

Those larger artifacts are still tracked in `../manifest.json` so later lanes can pull them in selectively if preview or 3D parity work needs them.
