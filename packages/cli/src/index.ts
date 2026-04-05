import {
  buildExecutionPlan,
  normalizeError,
  parseCliInput,
  renderHelpText,
  renderPlanResult,
  summarizeParity,
} from "@officekit/core";

const VERSION = "0.1.0";

export function runCli(rawArgv: string[]) {
  const input = parseCliInput(rawArgv);
  const [command] = input.argv;

  if (input.version) {
    return { exitCode: 0, stdout: VERSION };
  }

  if (input.help || !command || command === "help") {
    return { exitCode: 0, stdout: renderHelpText() };
  }

  if (command === "about") {
    return {
      exitCode: 0,
      stdout: summarizeParity().lineage,
    };
  }

  if (command === "contracts") {
    return {
      exitCode: 0,
      stdout: JSON.stringify(summarizeParity(), null, input.json ? 2 : 0),
    };
  }

  try {
    const plan = buildExecutionPlan(input.argv);
    if (input.plan || input.json) {
      return {
        exitCode: 0,
        stdout: renderPlanResult(plan, input.json),
      };
    }

    return {
      exitCode: 2,
      stdout: renderPlanResult(plan, false),
      stderr: "Scaffold-only command contract: rerun with --plan or wait for the owning package implementation.",
    };
  } catch (error) {
    const normalized = normalizeError(error);
    const body = input.json
      ? JSON.stringify(
          {
            error: normalized.message,
            code: normalized.code,
            suggestion: normalized.suggestion,
          },
          null,
          2,
        )
      : [normalized.message, normalized.suggestion].filter(Boolean).join("\n");

    return {
      exitCode: 1,
      stderr: body,
    };
  }
}
