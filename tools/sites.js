const { graphGet } = require("../graph");
const { ok, requireString } = require("../response");
const { formatDrive, formatSite, listText } = require("./formatters");

const tools = [
  {
    name: "list_sites",
    description:
      "Find SharePoint sites you can access. Use this when you need to locate a team site, department site, or project site before opening libraries or lists.",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "Optional site name or keyword to search for. Leave blank to list available sites.",
        },
      },
      required: [],
    },
  },
  {
    name: "get_site",
    description:
      "Get details for a specific SharePoint site when you already know the site ID.",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "The SharePoint site ID, usually in the format {hostname},{siteId},{webId}.",
        },
      },
      required: ["siteId"],
    },
  },
  {
    name: "list_site_drives",
    description:
      "List the document libraries available inside a SharePoint site, such as Documents or Shared Files.",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "The SharePoint site ID to inspect.",
        },
      },
      required: ["siteId"],
    },
  },
];

const handlers = {
  list_sites: async ({ accessToken, params }) => {
    const query = params?.query?.trim() || "*";
    const data = await graphGet(accessToken, `/sites`, { search: query });
    const sites = (data.value || []).map((site) => ({
      id: site.id,
      name: site.name,
      webUrl: site.webUrl,
    }));

    return ok(
      listText(
        query === "*" ? "Available SharePoint sites" : `SharePoint sites matching "${query}"`,
        sites.length,
        "site",
        sites,
        formatSite,
        query === "*"
          ? "No SharePoint sites were available for this account."
          : `No SharePoint sites matched "${query}".`
      ),
      { query, count: sites.length, sites }
    );
  },

  get_site: async ({ accessToken, params }) => {
    const siteId = requireString(params, "siteId", "a SharePoint site ID");
    const site = await graphGet(accessToken, `/sites/${encodeURIComponent(siteId)}`);

    return ok(
      [
        "SharePoint site details",
        `Name: ${site.name || "Unavailable"}`,
        `Site ID: ${site.id || siteId}`,
        `URL: ${site.webUrl || "Unavailable"}`,
      ].join("\n"),
      { site }
    );
  },

  list_site_drives: async ({ accessToken, params }) => {
    const siteId = requireString(params, "siteId", "a SharePoint site ID");
    const data = await graphGet(accessToken, `/sites/${encodeURIComponent(siteId)}/drives`);
    const drives = (data.value || []).map((drive) => ({
      id: drive.id,
      name: drive.name,
      webUrl: drive.webUrl,
    }));

    return ok(
      listText(
        "Document libraries in this site",
        drives.length,
        "library",
        drives,
        formatDrive,
        "No document libraries were found for this site."
      ),
      { siteId, count: drives.length, drives }
    );
  },
};

module.exports = { tools, handlers };
