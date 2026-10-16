class ToolError extends Error {
  constructor(publicMessage, options = {}) {
    super(options.internalMessage || publicMessage);
    this.name = "ToolError";
    this.publicMessage = publicMessage;
    this.statusCode = options.statusCode || 400;
    this.code = options.code || "tool_error";
    this.hint = options.hint;
    this.details = options.details;
  }
}

function userError(publicMessage, options = {}) {
  return new ToolError(publicMessage, options);
}

function ok(text, structuredContent = {}) {
  return {
    content: [{ type: "text", text }],
    structuredContent,
  };
}

function errorResult(err, fallbackMessage = "I couldn't complete that SharePoint request.") {
  const publicMessage =
    err?.publicMessage ||
    (err?.statusCode === 401
      ? "Your SharePoint connection is missing or has expired. Please reconnect and try again."
      : err?.statusCode === 403
      ? "You don't have permission to do that in SharePoint."
      : err?.statusCode === 404
      ? "I couldn't find the SharePoint item you asked for."
      : fallbackMessage);

  const parts = [publicMessage];
  if (err?.hint) parts.push(err.hint);

  return {
    isError: true,
    content: [{ type: "text", text: parts.join("\n\n") }],
    structuredContent: {
      error: {
        code: err?.code || "tool_error",
        statusCode: err?.statusCode || 500,
      },
    },
  };
}

function requireString(params, key, label = key) {
  const value = params?.[key];
  if (typeof value !== "string" || value.trim().length === 0) {
    throw userError(`Please provide ${label}.`, {
      code: "missing_parameter",
      statusCode: 400,
    });
  }
  return value.trim();
}

function requireObject(params, key, label = key) {
  const value = params?.[key];
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw userError(`Please provide ${label} as an object.`, {
      code: "invalid_parameter",
      statusCode: 400,
    });
  }
  return value;
}

function optionalPositiveNumber(value, label, { integer = false, min = 1, max } = {}) {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== "number" || Number.isNaN(value)) {
    throw userError(`${label} must be a number.`, {
      code: "invalid_parameter",
      statusCode: 400,
    });
  }
  if (integer && !Number.isInteger(value)) {
    throw userError(`${label} must be a whole number.`, {
      code: "invalid_parameter",
      statusCode: 400,
    });
  }
  if (value < min || (typeof max === "number" && value > max)) {
    const range =
      typeof max === "number" ? `between ${min} and ${max}` : `at least ${min}`;
    throw userError(`${label} must be ${range}.`, {
      code: "invalid_parameter",
      statusCode: 400,
    });
  }
  return value;
}

function joinBlocks(blocks) {
  return blocks.filter(Boolean).join("\n\n");
}

function bulletList(items) {
  return items.map((item) => `- ${item}`).join("\n");
}

function formatCount(count, singular, plural = `${singular}s`) {
  return `${count} ${count === 1 ? singular : plural}`;
}

module.exports = {
  ToolError,
  userError,
  ok,
  errorResult,
  requireString,
  requireObject,
  optionalPositiveNumber,
  joinBlocks,
  bulletList,
  formatCount,
};
