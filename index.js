require("dotenv").config();
const express = require("express");
const { errorResult, userError } = require("./response");

const app = express();
app.use(express.json());

const { PORT = 8080 } = process.env;

function getBearerToken(req) {
  const auth = req.headers["authorization"] || "";
  return auth.startsWith("Bearer ") ? auth.slice(7) : null;
}

// Load tool modules
const sites = require("./tools/sites");
const drives = require("./tools/drives");
const lists = require("./tools/lists");
const permissions = require("./tools/permissions");
const search = require("./tools/search");
const users = require("./tools/users");
const files = require("./tools/files");
const globalSearch = require("./tools/global-search");

const allTools = [
  ...sites.tools,
  ...drives.tools,
  ...lists.tools,
  ...permissions.tools,
  ...search.tools,
  ...users.tools,
  ...files.tools,
  ...globalSearch.tools,
];

const handlers = {
  ...sites.handlers,
  ...drives.handlers,
  ...lists.handlers,
  ...permissions.handlers,
  ...search.handlers,
  ...users.handlers,
  ...files.handlers,
  ...globalSearch.handlers,
};

app.post("/mcp", async (req, res) => {
  const { id, method, params } = req.body || {};
  const accessToken = getBearerToken(req);

  console.log(`[${new Date().toISOString()}] /mcp called:`, { method, id });
  console.log(accessToken ? "Access token received." : "No access token provided.");

  if (method === "initialize") {
    return res.json({
      jsonrpc: "2.0",
      id,
      result: {
        protocolVersion: "2024-11-05",
        capabilities: { tools: {} },
        serverInfo: { name: "sharepoint-mcp", version: "1.0.0" },
      },
    });
  }

  if (method === "tools/list") {
    return res.json({
      jsonrpc: "2.0",
      id,
      result: { tools: allTools },
    });
  }

  if (!accessToken) {
    if (method === "tools/call") {
      return res.json({
        jsonrpc: "2.0",
        id,
        result: errorResult(
          userError("Your SharePoint connection is missing or has expired. Please reconnect and try again.", {
            statusCode: 401,
            code: "authentication_required",
          })
        ),
      });
    }

    return res.status(401).json({
      jsonrpc: "2.0",
      id,
      error: { code: 401, message: "Authentication required" },
    });
  }

  if (method === "tools/call") {
    const toolName = params?.name;
    const handler = handlers[toolName];

    if (!handler) {
      return res.json({
        jsonrpc: "2.0",
        id,
        result: errorResult(
          userError("That SharePoint action is not available right now.", {
            statusCode: 404,
            code: "tool_not_found",
          })
        ),
      });
    }

    try {
      const result = await handler({ accessToken, params: params?.arguments || {} });
      return res.json({ jsonrpc: "2.0", id, result });
    } catch (err) {
      console.error(
        `[${new Date().toISOString()}] Error in tool '${toolName}':`,
        err?.message || err
      );
      return res.json({ jsonrpc: "2.0", id, result: errorResult(err) });
    }
  }

  return res.json({
    jsonrpc: "2.0",
    id,
    error: { code: -32601, message: "Method not found" },
  });
});

app.listen(PORT, () => console.log(`SharePoint MCP running on http://localhost:${PORT}`));
