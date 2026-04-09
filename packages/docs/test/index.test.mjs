import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createHelpTopicIndex, createLane4StatusReport, getCommandFamilySummary, getLineageStatement, loadLane4Ledger } from '../src/index.mjs';

const repoRoot = resolve(import.meta.dirname, '../../..');

test('lane-4 ledger keeps explicit lineage and target coverage', () => {
  const ledger = loadLane4Ledger();
  assert.equal(ledger.sourceProduct, 'OfficeCLI');
  assert.match(getLineageStatement(), /migrated from OfficeCLI/i);
  assert.equal(ledger.families.length, 4);

  for (const family of ledger.families) {
    assert.ok(family.targetPackage.startsWith('@officekit/'));
    assert.ok(family.sourceFiles.length > 0, `missing source files for ${family.id}`);
    assert.ok(family.requiredBehaviors.length > 0, `missing required behaviors for ${family.id}`);
    assert.ok(family.verificationLayer.length > 0, `missing verification layers for ${family.id}`);
  }
});

test('docs helpers expose help topics and command families', () => {
  const topics = createHelpTopicIndex();
  assert.equal(topics.length, 3);
  assert.ok(topics.some((topic) => topic.format === 'pptx' && topic.verbs.includes('watch')));
  assert.ok(getCommandFamilySummary().includes('skills'));
  assert.ok(getCommandFamilySummary().includes('help'));
  assert.equal(createLane4StatusReport().every((entry) => entry.status === 'scaffolded'), true);
});

test('root README and SKILL have no officecli references', () => {
  const readme = readFileSync(resolve(repoRoot, 'README.md'), 'utf8');
  const skill = readFileSync(resolve(repoRoot, 'SKILL.md'), 'utf8');

  assert.ok(!/officecli/i.test(readme), 'README should not reference officecli');
  assert.ok(!/officecli/i.test(skill), 'SKILL should not reference officecli');
  assert.ok(readme.includes('officekit'), 'README should reference officekit');
  assert.ok(skill.includes('officekit'), 'SKILL should reference officekit');
});
