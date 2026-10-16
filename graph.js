const axios = require("axios");
const { ToolError } = require("./response");

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

function graphFailure(status, fallbackDetail) {
  const detail = fallbackDetail || "Unknown Microsoft Graph error.";

  if (status === 400) {
    return new ToolError("SharePoint could not process that request with the information provided.", {
      statusCode: 400,
      code: "invalid_request",
      internalMessage: detail,
    });
  }

  if (status === 401) {
    return new ToolError("Your SharePoint connection is missing or has expired. Please reconnect and try again.", {
      statusCode: 401,
      code: "authentication_required",
      internalMessage: detail,
    });
  }

  if (status === 403) {
    return new ToolError("You do not have permission to access that SharePoint content.", {
      statusCode: 403,
      code: "forbidden",
      internalMessage: detail,
    });
  }

  if (status === 404) {
    return new ToolError("I couldn't find the SharePoint item you asked for.", {
      statusCode: 404,
      code: "not_found",
      internalMessage: detail,
    });
  }

  return new ToolError("SharePoint couldn't complete that request right now. Please try again.", {
    statusCode: status || 500,
    code: "graph_error",
    internalMessage: detail,
  });
}

async function graphGet(accessToken, path, params = {}) {
  try {
    const res = await axios.get(`${GRAPH_BASE}${path}`, {
      headers: { Authorization: `Bearer ${accessToken}` },
      params,
    });
    return res.data;
  } catch (err) {
    const status = err.response?.status;
    const message = err.response?.data?.error?.message || err.message;
    throw graphFailure(status, message);
  }
}

async function graphPost(accessToken, path, body = {}, params = {}) {
  try {
    const res = await axios.post(`${GRAPH_BASE}${path}`, body, {
      headers: { Authorization: `Bearer ${accessToken}` },
      params,
    });
    return res.data;
  } catch (err) {
    const status = err.response?.status;
    const message = err.response?.data?.error?.message || err.message;
    throw graphFailure(status, message);
  }
}

async function graphPatch(accessToken, path, body = {}) {
  try {
    const res = await axios.patch(`${GRAPH_BASE}${path}`, body, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    return res.data;
  } catch (err) {
    const status = err.response?.status;
    const message = err.response?.data?.error?.message || err.message;
    throw graphFailure(status, message);
  }
}

async function graphGetRaw(accessToken, path, raw = false) {
  try {
    const response = await axios.get(`${GRAPH_BASE}${path}`, {
      headers: { Authorization: `Bearer ${accessToken}` },
      responseType: raw ? "arraybuffer" : "json",
    });

    return raw ? Buffer.from(response.data) : response.data;
  } catch (err) {
    const status = err.response?.status;
    const message = err.response?.data?.error?.message || err.message;
    throw graphFailure(status, message);
  }
}

module.exports = { graphGet, graphPost, graphPatch, graphGetRaw };
