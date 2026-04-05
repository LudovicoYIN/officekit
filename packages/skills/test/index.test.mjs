import test from 'node:test';
import assert from 'node:assert/strict';
import { AGENT_TARGETS, SKILL_CATALOG, planSkillInstallTargets, renderBaseSkill } from '../src/index.mjs';

test('base officekit skill keeps lineage wording intact', () => {
  const skill = renderBaseSkill();
  assert.match(skill, /migrated from OfficeCLI/i);
  assert.match(skill, /Node\.js \+ Bun migration of OfficeCLI/i);
  assert.match(skill, /help instead of assuming command compatibility/i);
});

test('skill install targets preserve known agent directories', () => {
  const plan = planSkillInstallTargets({ homeEntries: ['.claude', '.agents', '.cursor'] });
  assert.equal(plan.length, 3);
  assert.deepEqual(plan.map((entry) => entry.displayName), ['Claude Code', 'Codex CLI', 'Cursor']);
  assert.ok(AGENT_TARGETS.some((target) => target.displayName === 'Codex CLI'));
  assert.ok(SKILL_CATALOG.some((entry) => entry.name === 'morph-ppt'));
});
