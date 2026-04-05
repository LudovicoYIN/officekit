export const AGENT_TARGETS = [
  { aliases: ['claude', 'claude-code'], displayName: 'Claude Code', detectDir: '.claude', skillDir: '.claude/skills' },
  { aliases: ['copilot', 'github-copilot'], displayName: 'GitHub Copilot', detectDir: '.copilot', skillDir: '.copilot/skills' },
  { aliases: ['codex', 'openai-codex'], displayName: 'Codex CLI', detectDir: '.agents', skillDir: '.agents/skills' },
  { aliases: ['cursor'], displayName: 'Cursor', detectDir: '.cursor', skillDir: '.cursor/skills' },
  { aliases: ['windsurf'], displayName: 'Windsurf', detectDir: '.windsurf', skillDir: '.windsurf/skills' },
  { aliases: ['minimax', 'minimax-cli'], displayName: 'MiniMax CLI', detectDir: '.minimax', skillDir: '.minimax/skills' },
  { aliases: ['opencode'], displayName: 'OpenCode', detectDir: '.opencode', skillDir: '.opencode/skills' },
  { aliases: ['openclaw'], displayName: 'OpenClaw', detectDir: '.openclaw', skillDir: '.openclaw/skills' },
  { aliases: ['nanobot'], displayName: 'NanoBot', detectDir: '.nanobot/workspace', skillDir: '.nanobot/workspace/skills' },
  { aliases: ['zeroclaw'], displayName: 'ZeroClaw', detectDir: '.zeroclaw/workspace', skillDir: '.zeroclaw/workspace/skills' }
];

export const SKILL_CATALOG = [
  { name: 'officekit', folder: 'officekit', description: 'Base officekit skill with migration-aware usage guidance.' },
  { name: 'pptx', folder: 'officekit-pptx', description: 'PowerPoint-focused authoring guide.' },
  { name: 'word', folder: 'officekit-docx', description: 'Word-focused authoring guide.' },
  { name: 'excel', folder: 'officekit-xlsx', description: 'Excel-focused authoring guide.' },
  { name: 'pitch-deck', folder: 'officekit-pitch-deck', description: 'Presentation workflow guide.' },
  { name: 'academic-paper', folder: 'officekit-academic-paper', description: 'Document-heavy writing guide.' },
  { name: 'data-dashboard', folder: 'officekit-data-dashboard', description: 'Spreadsheet/dashboard guide.' },
  { name: 'financial-model', folder: 'officekit-financial-model', description: 'Spreadsheet modeling guide.' },
  { name: 'morph-ppt', folder: 'morph-ppt', description: 'Specialized morph-transition PowerPoint guide.' }
];

export function renderBaseSkill({ cliName = 'officekit' } = {}) {
  return `---
name: ${cliName}
description: Node.js + Bun migration of OfficeCLI for Word, Excel, PowerPoint, preview, install, and skill workflows.
---

# ${cliName}

${cliName} is migrated from OfficeCLI and targets OfficeCLI v1 parity except MCP. Use help instead of assuming command compatibility. Keep the lineage statement explicit in downstream docs.
`;
}

export function planSkillInstallTargets({ homeEntries = [] } = {}) {
  return AGENT_TARGETS.filter((target) => homeEntries.includes(target.detectDir)).map((target) => ({
    displayName: target.displayName,
    installDir: `${target.skillDir}/officekit`,
    aliases: target.aliases
  }));
}
