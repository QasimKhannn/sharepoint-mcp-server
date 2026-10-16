const { graphGet } = require("../graph");
const { ok } = require("../response");

const tools = [
  {
    name: "get_current_user",
    description:
      "Get the profile of the currently connected Microsoft 365 user.",
    inputSchema: {
      type: "object",
      properties: {},
      required: [],
    },
  },
];

const handlers = {
  get_current_user: async ({ accessToken }) => {
    const me = await graphGet(accessToken, `/me`);

    return ok(
      [
        "Current Microsoft 365 user",
        `Name: ${me.displayName || "Unavailable"}`,
        `Email: ${me.mail || me.userPrincipalName || "Unavailable"}`,
        `Job title: ${me.jobTitle || "Unavailable"}`,
      ].join("\n"),
      { user: me }
    );
  },
};

module.exports = { tools, handlers };
