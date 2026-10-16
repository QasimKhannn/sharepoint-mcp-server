const { graphGet, graphPost, graphPatch } = require("../graph");
const { ok, requireObject, requireString } = require("../response");
const { formatListItem, listText } = require("./formatters");

const tools = [
  {
    name: "list_lists",
    description:
      "List the SharePoint lists available in a site, such as issue trackers, request lists, or custom business lists.",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "The SharePoint site ID.",
        },
      },
      required: ["siteId"],
    },
  },
  {
    name: "list_list_items",
    description:
      "List the items stored in a SharePoint list.",
    inputSchema: {
      type: "object",
      properties: {
        siteId: { type: "string", description: "The SharePoint site ID." },
        listId: { type: "string", description: "The SharePoint list ID." },
      },
      required: ["siteId", "listId"],
    },
  },
  {
    name: "create_list_item",
    description:
      "Create a new item in a SharePoint list by sending the field values to save.",
    inputSchema: {
      type: "object",
      properties: {
        siteId: { type: "string", description: "The SharePoint site ID." },
        listId: { type: "string", description: "The SharePoint list ID." },
        fields: {
          type: "object",
          description: "An object containing the field names and values for the new list item.",
        },
      },
      required: ["siteId", "listId", "fields"],
    },
  },
  {
    name: "update_list_item",
    description:
      "Update the field values for an existing SharePoint list item.",
    inputSchema: {
      type: "object",
      properties: {
        siteId: { type: "string", description: "The SharePoint site ID." },
        listId: { type: "string", description: "The SharePoint list ID." },
        itemId: { type: "string", description: "The SharePoint list item ID." },
        fields: {
          type: "object",
          description: "An object containing the field names and values to update.",
        },
      },
      required: ["siteId", "listId", "itemId", "fields"],
    },
  },
];

const handlers = {
  list_lists: async ({ accessToken, params }) => {
    const siteId = requireString(params, "siteId", "a SharePoint site ID");
    const data = await graphGet(accessToken, `/sites/${encodeURIComponent(siteId)}/lists`);
    const lists = data.value || [];

    return ok(
      listText(
        "Lists in this SharePoint site",
        lists.length,
        "list",
        lists,
        (list) => `${list.name || "Untitled list"} — List ID: ${list.id || "Unavailable"}`,
        "No lists were found for this site."
      ),
      { siteId, count: lists.length, lists }
    );
  },

  list_list_items: async ({ accessToken, params }) => {
    const siteId = requireString(params, "siteId", "a SharePoint site ID");
    const listId = requireString(params, "listId", "a SharePoint list ID");
    const data = await graphGet(
      accessToken,
      `/sites/${encodeURIComponent(siteId)}/lists/${listId}/items`,
      { expand: "fields" }
    );
    const items = data.value || [];

    return ok(
      listText(
        "Items in this SharePoint list",
        items.length,
        "item",
        items,
        formatListItem,
        "This list does not contain any items yet."
      ),
      { siteId, listId, count: items.length, items }
    );
  },

  create_list_item: async ({ accessToken, params }) => {
    const siteId = requireString(params, "siteId", "a SharePoint site ID");
    const listId = requireString(params, "listId", "a SharePoint list ID");
    const fields = requireObject(params, "fields", "list item fields");
    const created = await graphPost(
      accessToken,
      `/sites/${encodeURIComponent(siteId)}/lists/${listId}/items`,
      { fields }
    );

    return ok(
      [
        "List item created successfully",
        `Item ID: ${created.id || "Unavailable"}`,
        `Title: ${created.fields?.Title || created.fields?.title || "Unavailable"}`,
      ].join("\n"),
      { item: created }
    );
  },

  update_list_item: async ({ accessToken, params }) => {
    const siteId = requireString(params, "siteId", "a SharePoint site ID");
    const listId = requireString(params, "listId", "a SharePoint list ID");
    const itemId = requireString(params, "itemId", "a SharePoint list item ID");
    const fields = requireObject(params, "fields", "updated list item fields");
    const updated = await graphPatch(
      accessToken,
      `/sites/${encodeURIComponent(siteId)}/lists/${listId}/items/${itemId}/fields`,
      fields
    );

    return ok(
      ["List item updated successfully", `Item ID: ${itemId}`].join("\n"),
      { itemId, updatedFields: updated }
    );
  },
};

module.exports = { tools, handlers };
