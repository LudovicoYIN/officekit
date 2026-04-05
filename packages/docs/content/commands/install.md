# install

Install helpers for the Bun + Node.js `officekit` distribution.

## Responsibilities

- pick the correct release asset for the local platform
- decide the install target directory
- emit PATH instructions for the active shell
- persist config after installation

## Compatibility note

The migration keeps the install/update/config *behavior family* from OfficeCLI, but not the exact scripts or command syntax.
