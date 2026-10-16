const { graphGet, graphPost } = require("../graph");
const { ok, requireString, userError } = require("../response");
const { formatPermission, listText } = require("./formatters");

const tools = [
  {
    name: "list_item_permissions",
    description:
      "Show who has access to a SharePoint file or folder and what level of access they have.",
    inputSchema: {
      type: "object",
      properties: {
        driveId: { type: "string", description: "The document library ID." },
        itemId: { type: "string", description: "The file or folder item ID." },
      },
      required: ["driveId", "itemId"],
    },
  },
  {
    name: "create_sharing_link",
    description:
      "Create a shareable link for a SharePoint file or folder. You can choose whether people can only view it or also edit it.",
    inputSchema: {
      type: "object",
      properties: {
        driveId: { type: "string", description: "The document library ID." },
        itemId: { type: "string", description: "The file or folder item ID." },
        type: {
          type: "string",
          enum: ["view", "edit"],
          default: "view",
          description: "Use 'view' for read-only access or 'edit' to allow changes.",
        },
        scope: {
          type: "string",
          enum: ["anonymous", "organization"],
          default: "organization",
          description: "Use 'organization' to limit the link to people in your company, or 'anonymous' for anyone with the link.",
        },
      },
      required: ["driveId", "itemId"],
    },
  },
];

const handlers = {
  list_item_permissions: async ({ accessToken, params }) => {
    const driveId = requireString(params, "driveId", "a document library ID");
    const itemId = requireString(params, "itemId", "a file or folder item ID");
    const data = await graphGet(accessToken, `/drives/${driveId}/items/${itemId}/permissions`);
    const permissions = data.value || [];

    return ok(
      listText(
        "Permissions for this SharePoint item",
        permissions.length,
        "permission entry",
        permissions,
        formatPermission,
        "No explicit permissions were returned for this item."
      ),
      { driveId, itemId, count: permissions.length, permissions }
    );
  },

  create_sharing_link: async ({ accessToken, params }) => {
    const driveId = requireString(params, "driveId", "a document library ID");
    const itemId = requireString(params, "itemId", "a file or folder item ID");
    const type = params?.type || "view";
    const scope = params?.scope || "organization";

    if (!["view", "edit"].includes(type)) {
      throw userError("Sharing link type must be either 'view' or 'edit'.", {
        code: "invalid_parameter",
        statusCode: 400,
      });
    }

    if (!["anonymous", "organization"].includes(scope)) {
      throw userError("Sharing link scope must be either 'anonymous' or 'organization'.", {
        code: "invalid_parameter",
        statusCode: 400,
      });
    }

    const link = await graphPost(accessToken, `/drives/${driveId}/items/${itemId}/createLink`, {
      type,
      scope,
    });

    return ok(
      [
        "Sharing link created",
        `Access: ${type === "edit" ? "Can edit" : "View only"}`,
        `Audience: ${scope === "organization" ? "People in your organization" : "Anyone with the link"}`,
        `URL: ${link?.link?.webUrl || "Unavailable"}`,
      ].join("\n"),
      { sharingLink: link }
    );
  },
};

module.exports = { tools, handlers };
