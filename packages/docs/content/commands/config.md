# config

Top-level configuration surface for officekit.

## Responsibilities

- persist user config under `~/.officekit/config.json`
- record update-check timestamps
- keep migration lineage metadata explicit

## Notes

This package preserves the OfficeCLI behavior of lightweight local config, but the Node/Bun migration uses a JSON-first config helper surface that can be shared across CLI commands.
