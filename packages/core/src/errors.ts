export class OfficekitError extends Error {
  readonly code: string;
  readonly suggestion?: string;

  constructor(message: string, code = "officekit_error", suggestion?: string) {
    super(message);
    this.name = "OfficekitError";
    this.code = code;
    this.suggestion = suggestion;
  }
}

export class UnsupportedCapabilityError extends OfficekitError {
  constructor(capability: string) {
    super(
      `Capability '${capability}' is intentionally excluded from officekit v1 parity scope.`,
      "capability_excluded",
      "Use the parity ledger to confirm v1 scope; MCP remains excluded by design.",
    );
    this.name = "UnsupportedCapabilityError";
  }
}

export class UsageError extends OfficekitError {
  constructor(message: string, suggestion?: string) {
    super(message, "usage_error", suggestion);
    this.name = "UsageError";
  }
}
