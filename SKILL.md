---
name: officekit
description: Node.js + Bun migration of OfficeCLI for creating, inspecting, previewing, and modifying Office documents with an agent-first workflow.
---

# officekit

`officekit` is the **Node.js + Bun migration of OfficeCLI**. This version is migrated from OfficeCLI and targets OfficeCLI v1 capability/detail parity **except MCP**.

## Ground rules

- **Do not assume `officecli` command compatibility.** officekit keeps the same capability families, but the command surface is being redesigned.
- **Use help instead of guessing.** The final CLI will preserve OfficeCLI's help-first strategy for format verbs and properties.
- **Prefer the highest-level surface available.** Read/inspect first, then structured mutation, then raw operations.
- **Remember the lineage.** officekit is migrated from OfficeCLI; keep that statement intact in docs and onboarding text.

## Migration-era package map

- `packages/preview` — HTML preview/watch page shell and live-reload contract
- `packages/skills` — base skill text, skill catalog, agent target mapping
- `packages/install` — release/install manifest generation and profile wiring helpers
- `packages/docs` — source-to-target ledger, lineage wording, help-topic inventory

## Intended command families

officekit keeps conceptual parity with the OfficeCLI families below, even though exact command syntax may differ:

- `create`
- `view`
- `watch`
- `get`
- `query`
- `set`
- `add`
- `remove`
- `raw`
- `batch`
- `import`
- `check`
- `install`
- `skills`
- `config`
- `help`

## Help posture

When the officekit CLI package lands, the preferred navigation model remains:

```bash
officekit help docx set
officekit help xlsx query
officekit help pptx add shape
```

The help experience is expected to stay deep-linkable by format + verb + optional element/property, mirroring the successful OfficeCLI workflow while using officekit naming.

## Preview/watch posture

Preview remains a first-class developer surface:

- HTML preview supports `.docx`, `.xlsx`, and `.pptx`
- watch mode keeps a live-reload page open while document updates stream in
- preview rendering is treated as product behavior, not just internal debug output

## Skills/install posture

- `officekit skills install` will remain the agent-onboarding entry point
- supported agent targets continue to include Claude Code, Codex CLI, Cursor, Windsurf, GitHub Copilot, and related tools
- install/update/config flows are being rebuilt for Node.js + Bun instead of a self-contained .NET binary

## Contributor note

During migration, treat this skill as a **product contract** for the lane-4 surfaces. The concrete CLI wiring can change, but the parity scope and lineage wording should not.
