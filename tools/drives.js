const { graphGet, graphPost } = require("../graph");
const { ok, requireString } = require("../response");
const { formatDriveItem, listText } = require("./formatters");

const tools = [
  {
    name: "list_drive_root_items",
    description:
      "List the top-level folders and files in a SharePoint document library.",
    inputSchema: {
      type: "object",
      properties: {
        driveId: {
          type: "string",
          description: "The document library ID.",
        },
      },
      required: ["driveId"],
    },
  },
  {
    name: "list_drive_children",
    description:
      "List the folders and files inside a specific SharePoint folder.",
    inputSchema: {
      type: "object",
      properties: {
        driveId: { type: "string", description: "The document library ID." },
        itemId: {
          type: "string",
          description: "The folder item ID whose contents you want to view.",
        },
      },
      required: ["driveId", "itemId"],
    },
  },
  {
    name: "get_drive_item",
    description:
      "Get details for a SharePoint file or folder, including its name, type, and web link.",
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
    name: "create_folder",
    description:
      "Create a new folder in a SharePoint document library. If a folder with the same name already exists, SharePoint will rename the new one.",
    inputSchema: {
      type: "object",
      properties: {
        driveId: { type: "string", description: "The document library ID." },
        parentItemId: {
          type: "string",
          description: "The parent folder item ID. Use 'root' to create the folder at the top level.",
        },
        name: { type: "string", description: "The folder name to create." },
      },
      required: ["driveId", "parentItemId", "name"],
    },
  },
  {
    name: "copy_drive_item",
    description:
      "Copy a SharePoint file or folder into another folder. SharePoint may complete the copy asynchronously.",
    inputSchema: {
      type: "object",
      properties: {
        driveId: { type: "string", description: "The source document library ID." },
        itemId: { type: "string", description: "The file or folder item ID to copy." },
        newParentItemId: {
          type: "string",
          description: "The destination folder item ID.",
        },
        newName: {
          type: "string",
          description: "Optional new name for the copied item.",
        },
      },
      required: ["driveId", "itemId", "newParentItemId"],
    },
  },
];

function sortByPurpleCheck(items) {
  return items
    .map((item) => {
      const moderationStatus = item.listItem?.fields?.["_ModerationStatus"];
      const isPurpleCheck = moderationStatus === 0 || moderationStatus === "0";
      return { ...item, isPurpleCheck };
    })
    .sort((a, b) => {
      if (a.isPurpleCheck === b.isPurpleCheck) return 0;
      return a.isPurpleCheck ? -1 : 1;
    });
}

const handlers = {
  list_drive_root_items: async ({ accessToken, params }) => {
    const driveId = requireString(params, "driveId", "a document library ID");
    const data = await graphGet(accessToken, `/drives/${driveId}/root/children`, {
      $expand: "listItem($expand=fields)",
    });
    const items = sortByPurpleCheck(data.value || []);

    return ok(
      listText(
        "Top-level items in this document library",
        items.length,
        "item",
        items,
        formatDriveItem,
        "This document library does not contain any top-level items yet."
      ),
      { driveId, count: items.length, items }
    );
  },

  list_drive_children: async ({ accessToken, params }) => {
    const driveId = requireString(params, "driveId", "a document library ID");
    const itemId = requireString(params, "itemId", "a folder item ID");
    const data = await graphGet(accessToken, `/drives/${driveId}/items/${itemId}/children`, {
      $expand: "listItem($expand=fields)",
    });
    const items = sortByPurpleCheck(data.value || []);

    return ok(
      listText(
        "Items in this folder",
        items.length,
        "item",
        items,
        formatDriveItem,
        "This folder is empty."
      ),
      { driveId, itemId, count: items.length, items }
    );
  },

  get_drive_item: async ({ accessToken, params }) => {
    const driveId = requireString(params, "driveId", "a document library ID");
    const itemId = requireString(params, "itemId", "a file or folder item ID");
    const item = await graphGet(accessToken, `/drives/${driveId}/items/${itemId}`);

    return ok(
      [
        "SharePoint item details",
        `Name: ${item.name || "Unavailable"}`,
        `Type: ${item.folder ? "Folder" : item.file ? "File" : "Item"}`,
        `Item ID: ${item.id || itemId}`,
        `URL: ${item.webUrl || "Unavailable"}`,
      ].join("\n"),
      { item }
    );
  },

  create_folder: async ({ accessToken, params }) => {
    const driveId = requireString(params, "driveId", "a document library ID");
    const parentItemId = requireString(params, "parentItemId", "a parent folder ID or 'root'");
    const name = requireString(params, "name", "a folder name");
    const parentPath =
      parentItemId === "root"
        ? `/drives/${driveId}/root/children`
        : `/drives/${driveId}/items/${parentItemId}/children`;

    const created = await graphPost(accessToken, parentPath, {
      name,
      folder: {},
      "@microsoft.graph.conflictBehavior": "rename",
    });

    return ok(
      [
        "Folder created successfully",
        `Name: ${created.name || name}`,
        `Item ID: ${created.id || "Unavailable"}`,
        `URL: ${created.webUrl || "Unavailable"}`,
      ].join("\n"),
      { folder: created }
    );
  },

  copy_drive_item: async ({ accessToken, params }) => {
    const driveId = requireString(params, "driveId", "a document library ID");
    const itemId = requireString(params, "itemId", "a file or folder item ID");
    const newParentItemId = requireString(params, "newParentItemId", "a destination folder ID");
    const newName =
      typeof params?.newName === "string" && params.newName.trim().length
        ? params.newName.trim()
        : undefined;

    const payload = {
      parentReference: { id: newParentItemId },
      ...(newName ? { name: newName } : {}),
    };

    const result = await graphPost(accessToken, `/drives/${driveId}/items/${itemId}/copy`, payload);

    return ok(
      [
        "Copy request submitted",
        "SharePoint accepted the copy request. Large items may take a little time to appear in the destination folder.",
        newName ? `Requested new name: ${newName}` : "",
      ]
        .filter(Boolean)
        .join("\n"),
      { driveId, itemId, destinationItemId: newParentItemId, result: result || { accepted: true } }
    );
  },
};

module.exports = { tools, handlers };
