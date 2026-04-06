import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const moduleDir = dirname(fileURLToPath(import.meta.url));
const ledgerPath = resolve(moduleDir, '../fixtures/lane4-source-ledger.json');

export function loadLane4Ledger() {
  return JSON.parse(readFileSync(ledgerPath, 'utf8'));
}

export function getLineageStatement() {
  return loadLane4Ledger().lineageNote;
}

export function getAudienceSections() {
  return [
    {
      id: 'ai-agents',
      title: 'AI agents',
      promise: 'Keep skill-first onboarding and deep help navigation.'
    },
    {
      id: 'developers',
      title: 'Developers',
      promise: 'Keep preview/watch and differential parity evidence as product surfaces.'
    },
    {
      id: 'maintainers',
      title: 'Maintainers',
      promise: 'Track migration status with explicit scaffolded/implemented/verified states.'
    }
  ];
}

export function getCommandFamilySummary() {
  return ['create', 'view', 'watch', 'get', 'query', 'set', 'add', 'remove', 'raw', 'batch', 'import', 'check', 'install', 'skills', 'config', 'help'];
}

export function createHelpTopicIndex() {
  return [
    { format: 'docx', verbs: ['view', 'get', 'query', 'set', 'add', 'raw'] },
    { format: 'xlsx', verbs: ['view', 'get', 'query', 'set', 'add', 'raw'] },
    { format: 'pptx', verbs: ['view', 'get', 'query', 'set', 'add', 'raw', 'watch'] }
  ];
}

export function createLane4StatusReport() {
  const ledger = loadLane4Ledger();
  return ledger.families.map((family) => ({
    id: family.id,
    targetPackage: family.targetPackage,
    status: family.status,
    requiredBehaviors: family.requiredBehaviors.length,
    sourceFileCount: family.sourceFiles.length
  }));
}
